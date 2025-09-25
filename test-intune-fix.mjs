#!/usr/bin/env node

/**
 * Test script to verify Intune authentication fix
 * This script tests the new resource-aware authentication for Intune endpoints
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import { createIntuneGraphClient, createStandardGraphClient, isIntuneEndpoint } from './src/utils/modern-graph-client.js';
import dotenv from 'dotenv';

// Load environment variables
dotenv.config();

const {
  AZURE_TENANT_ID,
  AZURE_CLIENT_ID,
  AZURE_CLIENT_SECRET
} = process.env;

if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET) {
  console.error('❌ Missing required environment variables:');
  console.error('   AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET');
  process.exit(1);
}

async function testIntuneAuthentication() {
  console.log('🔧 Testing Intune Authentication Fix\n');

  try {
    // Create credentials
    const credential = new ClientSecretCredential(
      AZURE_TENANT_ID,
      AZURE_CLIENT_ID,
      AZURE_CLIENT_SECRET
    );

    // Create authentication provider
    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: ['https://graph.microsoft.com/.default']
    });

    // Create base Graph client
    const graphClient = Client.initWithMiddleware({ authProvider });

    console.log('✅ Base Graph client created successfully');

    // Test 1: Verify endpoint detection
    console.log('\n📋 Test 1: Endpoint Detection');
    const testEndpoints = [
      '/deviceManagement/deviceConfigurations',
      '/deviceAppManagement/mobileApps',
      '/informationProtection/bitlocker/recoveryKeys',
      '/users',
      '/groups'
    ];

    testEndpoints.forEach(endpoint => {
      const isIntune = isIntuneEndpoint(endpoint);
      console.log(`   ${endpoint}: ${isIntune ? '🔒 Intune' : '📊 Standard'}`);
    });

    // Test 2: Create resource-specific clients
    console.log('\n📋 Test 2: Client Creation');
    const standardClient = createStandardGraphClient(graphClient);
    const intuneClient = createIntuneGraphClient(graphClient);

    console.log(`   Standard client resource: ${standardClient.getResource()}`);
    console.log(`   Intune client resource: ${intuneClient.getResource()}`);

    // Test 3: Test standard Graph API call (should work)
    console.log('\n📋 Test 3: Standard Graph API Test');
    try {
      const usersResponse = await standardClient.makeApiCall('/me', { 
        select: ['id', 'displayName', 'userPrincipalName'] 
      });
      console.log('   ✅ Standard Graph API call successful');
      console.log(`   📊 Response: ${JSON.stringify(usersResponse.data, null, 2)}`);
    } catch (error) {
      console.log(`   ⚠️ Standard Graph API call failed: ${error.message}`);
    }

    // Test 4: Test Intune API call (the main fix)
    console.log('\n📋 Test 4: Intune API Test (Device Configurations)');
    try {
      const intuneResponse = await intuneClient.makeApiCall('/deviceManagement/deviceConfigurations', {
        select: ['id', 'displayName', '@odata.type'],
        top: 5
      });
      console.log('   🎉 INTUNE API CALL SUCCESSFUL! The fix is working!');
      console.log(`   📊 Found ${intuneResponse.data.value?.length || 0} device configurations`);
      
      if (intuneResponse.data.value && intuneResponse.data.value.length > 0) {
        console.log('   📋 Sample configurations:');
        intuneResponse.data.value.slice(0, 3).forEach((config, index) => {
          console.log(`      ${index + 1}. ${config.displayName} (${config['@odata.type']})`);
        });
      }
    } catch (error) {
      console.log(`   ❌ Intune API call failed: ${error.message}`);
      console.log(`   🔍 Error details:`, error);
      
      // Additional debugging
      if (error.message.includes('403') || error.message.includes('Access denied')) {
        console.log('\n🔍 Debugging Information:');
        console.log('   - This might still be a token scope issue');
        console.log('   - Check if the app registration has the correct permissions');
        console.log('   - Verify admin consent has been granted');
        console.log('   - The resource parameter might need adjustment');
      }
    }

    // Test 5: Test other Intune endpoints
    console.log('\n📋 Test 5: Additional Intune Endpoints');
    const intuneEndpoints = [
      '/deviceManagement/deviceCompliancePolicies',
      '/deviceAppManagement/mobileApps',
      '/deviceManagement/managedDevices'
    ];

    for (const endpoint of intuneEndpoints) {
      try {
        const response = await intuneClient.makeApiCall(endpoint, { top: 1 });
        console.log(`   ✅ ${endpoint}: Success (${response.data.value?.length || 0} items)`);
      } catch (error) {
        console.log(`   ❌ ${endpoint}: Failed - ${error.message}`);
      }
    }

    // Test 6: Compare token acquisition
    console.log('\n📋 Test 6: Token Analysis');
    try {
      // Get tokens for both resources
      const standardToken = await credential.getToken(['https://graph.microsoft.com/.default']);
      const intuneToken = await credential.getToken(['https://manage.microsoft.com/.default']);
      
      console.log('   📊 Standard Graph token acquired:', !!standardToken);
      console.log('   🔒 Intune token acquired:', !!intuneToken);
      
      if (standardToken && intuneToken) {
        console.log('   🎯 Both tokens acquired successfully');
        
        // Decode token claims (basic info only)
        const decodeToken = (token) => {
          try {
            const payload = token.split('.')[1];
            const decoded = JSON.parse(Buffer.from(payload, 'base64').toString());
            return {
              aud: decoded.aud,
              roles: decoded.roles || [],
              scp: decoded.scp || '',
              appid: decoded.appid
            };
          } catch (e) {
            return { error: 'Could not decode token' };
          }
        };
        
        const standardClaims = decodeToken(standardToken.token);
        const intuneClaims = decodeToken(intuneToken.token);
        
        console.log('   📊 Standard token audience:', standardClaims.aud);
        console.log('   🔒 Intune token audience:', intuneClaims.aud);
        console.log('   📊 Standard token roles:', standardClaims.roles);
        console.log('   🔒 Intune token roles:', intuneClaims.roles);
      }
    } catch (error) {
      console.log(`   ⚠️ Token analysis failed: ${error.message}`);
    }

  } catch (error) {
    console.error('❌ Test failed with error:', error);
    console.error('Stack trace:', error.stack);
  }
}

// Run the test
testIntuneAuthentication().then(() => {
  console.log('\n🏁 Test completed');
}).catch(error => {
  console.error('💥 Test crashed:', error);
  process.exit(1);
});
