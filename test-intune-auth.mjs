#!/usr/bin/env node

/**
 * Test script to diagnose Intune authentication issues
 * This will help identify why deviceConfiguration endpoints are failing
 */

import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';

// Load environment variables
const MS_TENANT_ID = process.env.MS_TENANT_ID;
const MS_CLIENT_ID = process.env.MS_CLIENT_ID;
const MS_CLIENT_SECRET = process.env.MS_CLIENT_SECRET;

console.log('🔍 M365 Core MCP - Intune Authentication Diagnostic Tool');
console.log('==================================================');

// Validate environment variables
if (!MS_TENANT_ID || !MS_CLIENT_ID || !MS_CLIENT_SECRET) {
  console.error('❌ Missing required environment variables:');
  console.error(`   MS_TENANT_ID: ${MS_TENANT_ID ? '✅ Present' : '❌ Missing'}`);
  console.error(`   MS_CLIENT_ID: ${MS_CLIENT_ID ? '✅ Present' : '❌ Missing'}`);
  console.error(`   MS_CLIENT_SECRET: ${MS_CLIENT_SECRET ? '✅ Present' : '❌ Missing'}`);
  process.exit(1);
}

console.log('✅ Environment variables configured');
console.log(`   Tenant ID: ${MS_TENANT_ID}`);
console.log(`   Client ID: ${MS_CLIENT_ID}`);
console.log(`   Client Secret: ${MS_CLIENT_SECRET ? '***' + MS_CLIENT_SECRET.slice(-4) : 'Missing'}`);

// Test token acquisition
async function getAccessToken(scope = 'https://graph.microsoft.com/.default') {
  console.log(`\n🔐 Acquiring access token for scope: ${scope}`);
  
  const tokenEndpoint = `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append('client_id', MS_CLIENT_ID);
  params.append('client_secret', MS_CLIENT_SECRET);
  params.append('grant_type', 'client_credentials');
  params.append('scope', scope);

  try {
    const response = await fetch(tokenEndpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: params,
    });

    if (!response.ok) {
      const errorData = await response.text();
      console.error(`❌ Token acquisition failed: ${response.status} ${response.statusText}`);
      console.error(`   Error details: ${errorData}`);
      throw new Error(`Token acquisition failed: ${response.status}`);
    }

    const data = await response.json();
    console.log('✅ Token acquired successfully');
    console.log(`   Expires in: ${data.expires_in} seconds`);
    console.log(`   Token type: ${data.token_type}`);
    console.log(`   Scope: ${data.scope}`);
    
    // Parse the token to check claims
    if (data.access_token) {
      const tokenParts = data.access_token.split('.');
      if (tokenParts.length === 3) {
        try {
          const payload = JSON.parse(Buffer.from(tokenParts[1], 'base64').toString());
          console.log('   Token claims:');
          console.log(`     App ID: ${payload.appid || payload.azp || 'Not found'}`);
          console.log(`     Tenant ID: ${payload.tid || 'Not found'}`);
          console.log(`     Roles: ${payload.roles ? payload.roles.join(', ') : 'Not found'}`);
          console.log(`     Scopes: ${payload.scp || 'Not found'}`);
          console.log(`     Audience: ${payload.aud || 'Not found'}`);
        } catch (parseError) {
          console.log('   Could not parse token claims');
        }
      }
    }
    
    return data.access_token;
  } catch (error) {
    console.error(`❌ Token acquisition error: ${error.message}`);
    throw error;
  }
}

// Test Graph API endpoints
async function testGraphEndpoint(graphClient, endpoint, description) {
  console.log(`\n🧪 Testing: ${description}`);
  console.log(`   Endpoint: ${endpoint}`);
  
  try {
    const result = await graphClient.api(endpoint).get();
    console.log(`✅ Success: ${description}`);
    console.log(`   Result type: ${typeof result}`);
    console.log(`   Has value array: ${Array.isArray(result?.value)}`);
    console.log(`   Count: ${result?.value?.length || 'N/A'}`);
    return result;
  } catch (error) {
    console.error(`❌ Failed: ${description}`);
    console.error(`   Status: ${error.status || 'Unknown'}`);
    console.error(`   Code: ${error.code || 'Unknown'}`);
    console.error(`   Message: ${error.message || 'Unknown'}`);
    
    if (error.body) {
      try {
        const errorBody = typeof error.body === 'string' ? JSON.parse(error.body) : error.body;
        console.error(`   Error details: ${JSON.stringify(errorBody, null, 2)}`);
      } catch (parseError) {
        console.error(`   Raw error body: ${error.body}`);
      }
    }
    
    return null;
  }
}

// Main diagnostic function
async function runDiagnostics() {
  try {
    // Step 1: Get access token
    const token = await getAccessToken();
    
    // Step 2: Create Graph client
    console.log('\n🔧 Creating Microsoft Graph client...');
    const graphClient = Client.init({
      authProvider: (callback) => {
        callback(null, token);
      }
    });
    console.log('✅ Graph client created');
    
    // Step 3: Test basic Graph API
    await testGraphEndpoint(graphClient, '/me', 'Current user info (should fail for app-only)');
    
    // Step 4: Test organization info
    await testGraphEndpoint(graphClient, '/organization', 'Organization information');
    
    // Step 5: Test Intune device management endpoints
    console.log('\n📱 Testing Intune endpoints...');
    
    // Test managed devices (should work)
    await testGraphEndpoint(graphClient, '/deviceManagement/managedDevices', 'Managed devices');
    
    // Test device configurations (the problematic endpoint)
    await testGraphEndpoint(graphClient, '/deviceManagement/deviceConfigurations', 'Device configurations (PROBLEM ENDPOINT)');
    
    // Test compliance policies
    await testGraphEndpoint(graphClient, '/deviceManagement/deviceCompliancePolicies', 'Device compliance policies');
    
    // Test device configuration service endpoint (direct Intune API)
    console.log('\n🔧 Testing direct Intune service endpoints...');
    
    // Test with different API versions
    const intuneEndpoints = [
      '/deviceManagement/deviceConfigurations?api-version=2022-06-13',
      '/deviceManagement/deviceConfigurations?api-version=beta',
      '/deviceManagement/deviceConfigurations?$top=1'
    ];
    
    for (const endpoint of intuneEndpoints) {
      await testGraphEndpoint(graphClient, endpoint, `Device configurations with ${endpoint.includes('api-version') ? 'API version' : 'top filter'}`);
    }
    
    // Step 6: Test with explicit headers for Intune
    console.log('\n🔧 Testing with Intune-specific headers...');
    try {
      const result = await graphClient
        .api('/deviceManagement/deviceConfigurations')
        .header('Accept', 'application/json')
        .header('Content-Type', 'application/json')
        .header('User-Agent', 'M365-Core-MCP/1.0')
        .header('X-MS-Client-Request-Id', crypto.randomUUID())
        .get();
        
      console.log('✅ Success with explicit headers');
      console.log(`   Count: ${result?.value?.length || 'N/A'}`);
    } catch (error) {
      console.error('❌ Failed with explicit headers');
      console.error(`   Status: ${error.status}`);
      console.error(`   Message: ${error.message}`);
    }
    
    console.log('\n📊 Diagnostic Summary:');
    console.log('======================');
    console.log('If device configurations are failing but other endpoints work,');
    console.log('this suggests a specific permission or scope issue with Intune.');
    console.log('');
    console.log('Next steps:');
    console.log('1. Verify DeviceManagementConfiguration.ReadWrite.All is granted');
    console.log('2. Check if admin consent was provided');
    console.log('3. Verify the app registration has the correct permissions');
    console.log('4. Try using a different API endpoint or version');
    
  } catch (error) {
    console.error(`\n💥 Diagnostic failed: ${error.message}`);
    console.error('Stack trace:', error.stack);
  }
}

// Import crypto for UUID generation
import { randomUUID } from 'crypto';
const crypto = { randomUUID };

// Run the diagnostics
runDiagnostics();