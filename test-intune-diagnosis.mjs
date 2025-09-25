#!/usr/bin/env node

/**
 * Test script to diagnose Intune authorization using existing server infrastructure
 */

import { M365CoreServer } from './src/server.js';

async function testIntuneAuth() {
  console.log('🔍 Testing Intune Authentication with M365 Core Server');
  console.log('====================================================');

  try {
    // Create server instance
    const server = new M365CoreServer();
    
    // Test authentication first
    console.log('\n🔐 Testing authentication...');
    try {
      await server.ensureAuthenticated();
      console.log('✅ Authentication successful');
    } catch (error) {
      console.error(`❌ Authentication failed: ${error.message}`);
      return;
    }

    // Get the Graph client
    const graphClient = server.getGraphClient();
    
    // Test basic Graph API
    console.log('\n🧪 Testing basic Graph API...');
    try {
      const orgs = await graphClient.api('/organization').get();
      console.log(`✅ Organization API works (${orgs.value?.length || 0} orgs)`);
    } catch (error) {
      console.error(`❌ Organization API failed: ${error.message}`);
    }

    // Test Intune managed devices (should work)
    console.log('\n📱 Testing Intune managed devices...');
    try {
      const devices = await graphClient.api('/deviceManagement/managedDevices').top(1).get();
      console.log(`✅ Managed devices API works (${devices.value?.length || 0} devices)`);
    } catch (error) {
      console.error(`❌ Managed devices API failed: ${error.status} - ${error.message}`);
      if (error.body) {
        console.error(`   Details: ${JSON.stringify(error.body, null, 2)}`);
      }
    }

    // Test the problematic endpoint: deviceConfigurations
    console.log('\n⚠️  Testing device configurations (the problem endpoint)...');
    try {
      const configs = await graphClient.api('/deviceManagement/deviceConfigurations').top(1).get();
      console.log(`✅ Device configurations API works! (${configs.value?.length || 0} configs)`);
    } catch (error) {
      console.error(`❌ Device configurations API failed: ${error.status} - ${error.message}`);
      
      // Parse detailed error information
      if (error.body) {
        try {
          const errorBody = typeof error.body === 'string' ? JSON.parse(error.body) : error.body;
          console.error(`   Error details: ${JSON.stringify(errorBody, null, 2)}`);
        } catch (parseError) {
          console.error(`   Raw error body: ${error.body}`);
        }
      }
      
      // Test alternative approaches
      console.log('\n🔧 Trying alternative approaches...');
      
      // Try with beta endpoint
      try {
        console.log('   Testing beta endpoint...');
        const betaConfigs = await graphClient.api('/deviceManagement/deviceConfigurations').version('beta').top(1).get();
        console.log(`   ✅ Beta endpoint works! (${betaConfigs.value?.length || 0} configs)`);
      } catch (betaError) {
        console.error(`   ❌ Beta endpoint also failed: ${betaError.status} - ${betaError.message}`);
      }
      
      // Try with specific headers
      try {
        console.log('   Testing with specific headers...');
        const headerConfigs = await graphClient
          .api('/deviceManagement/deviceConfigurations')
          .header('ConsistencyLevel', 'eventual')
          .header('Prefer', 'return=minimal')
          .top(1)
          .get();
        console.log(`   ✅ Headers approach works! (${headerConfigs.value?.length || 0} configs)`);
      } catch (headerError) {
        console.error(`   ❌ Headers approach failed: ${headerError.status} - ${headerError.message}`);
      }
    }

    // Test device compliance policies
    console.log('\n📋 Testing device compliance policies...');
    try {
      const compliance = await graphClient.api('/deviceManagement/deviceCompliancePolicies').top(1).get();
      console.log(`✅ Compliance policies API works (${compliance.value?.length || 0} policies)`);
    } catch (error) {
      console.error(`❌ Compliance policies API failed: ${error.status} - ${error.message}`);
    }

    // Test using the actual Intune handler
    console.log('\n🔧 Testing with actual Intune Windows handler...');
    try {
      const { handleIntuneWindowsPolicies } = await import('./src/handlers/intune-windows-handler.js');
      
      const result = await handleIntuneWindowsPolicies(graphClient, {
        action: 'list',
        policyType: 'Configuration'
      });
      
      console.log('✅ Intune Windows handler works!');
      console.log(`   Result: ${JSON.stringify(result, null, 2)}`);
      
    } catch (handlerError) {
      console.error(`❌ Intune Windows handler failed: ${handlerError.message}`);
      if (handlerError.stack) {
        console.error(`   Stack: ${handlerError.stack}`);
      }
    }

    console.log('\n📊 Diagnostic Summary');
    console.log('===================');
    console.log('This test helps identify if the issue is:');
    console.log('1. Authentication/token problems');
    console.log('2. Permission scope issues');
    console.log('3. Endpoint-specific problems');
    console.log('4. Microsoft Graph SDK issues');

  } catch (error) {
    console.error(`💥 Test failed: ${error.message}`);
    if (error.stack) {
      console.error(`Stack trace: ${error.stack}`);
    }
  }
}

testIntuneAuth().catch(console.error);