#!/usr/bin/env node

/**
 * Focused test for Intune deviceConfiguration API permission issues
 * This tests the exact scenario from the error report
 */

import 'isomorphic-fetch';

// Environment variables
const MS_TENANT_ID = process.env.MS_TENANT_ID;
const MS_CLIENT_ID = process.env.MS_CLIENT_ID;
const MS_CLIENT_SECRET = process.env.MS_CLIENT_SECRET;

console.log('🔍 Intune Device Configuration Permission Test');
console.log('=============================================');

async function testIntunePermissions() {
  if (!MS_TENANT_ID || !MS_CLIENT_ID || !MS_CLIENT_SECRET) {
    console.error('❌ Missing environment variables');
    return;
  }

  try {
    // Get token
    console.log('🔐 Getting access token...');
    const tokenResponse = await fetch(`https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: MS_CLIENT_ID,
        client_secret: MS_CLIENT_SECRET,
        grant_type: 'client_credentials',
        scope: 'https://graph.microsoft.com/.default'
      })
    });

    if (!tokenResponse.ok) {
      const error = await tokenResponse.text();
      throw new Error(`Token request failed: ${error}`);
    }

    const tokenData = await tokenResponse.json();
    const token = tokenData.access_token;
    console.log('✅ Token acquired');

    // Decode token to check permissions
    const tokenParts = token.split('.');
    const payload = JSON.parse(Buffer.from(tokenParts[1], 'base64').toString());
    console.log('\n🎫 Token Details:');
    console.log(`   App ID: ${payload.appid}`);
    console.log(`   Tenant: ${payload.tid}`);
    console.log(`   Audience: ${payload.aud}`);
    console.log(`   Roles: ${payload.roles ? payload.roles.join(', ') : 'None'}`);

    // Test the exact failing endpoint
    console.log('\n🧪 Testing deviceManagement/deviceConfigurations...');
    
    const headers = {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      'User-Agent': 'M365-Core-MCP/1.0'
    };

    // Test 1: Basic request
    console.log('\n📍 Test 1: Basic GET request');
    try {
      const response = await fetch('https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations', {
        method: 'GET',
        headers
      });

      if (response.ok) {
        const data = await response.json();
        console.log(`✅ Success! Found ${data.value?.length || 0} configurations`);
      } else {
        const errorText = await response.text();
        console.error(`❌ Failed: ${response.status} ${response.statusText}`);
        console.error(`   Response: ${errorText}`);
        
        // Try to parse error details
        try {
          const errorObj = JSON.parse(errorText);
          if (errorObj.error) {
            console.error(`   Error code: ${errorObj.error.code}`);
            console.error(`   Error message: ${errorObj.error.message}`);
          }
        } catch (parseError) {
          // Error response wasn't JSON
        }
      }
    } catch (error) {
      console.error(`❌ Request failed: ${error.message}`);
    }

    // Test 2: With $top=1 to minimize response
    console.log('\n📍 Test 2: With $top=1 parameter');
    try {
      const response = await fetch('https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations?$top=1', {
        method: 'GET',
        headers
      });

      if (response.ok) {
        const data = await response.json();
        console.log(`✅ Success! Response type: ${typeof data}`);
      } else {
        const errorText = await response.text();
        console.error(`❌ Failed: ${response.status} ${response.statusText}`);
      }
    } catch (error) {
      console.error(`❌ Request failed: ${error.message}`);
    }

    // Test 3: Beta endpoint
    console.log('\n📍 Test 3: Beta endpoint');
    try {
      const response = await fetch('https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations', {
        method: 'GET',
        headers
      });

      if (response.ok) {
        const data = await response.json();
        console.log(`✅ Success with beta! Found ${data.value?.length || 0} configurations`);
      } else {
        const errorText = await response.text();
        console.error(`❌ Beta failed: ${response.status} ${response.statusText}`);
      }
    } catch (error) {
      console.error(`❌ Beta request failed: ${error.message}`);
    }

    // Test 4: Compare with working endpoint
    console.log('\n📍 Test 4: Compare with managedDevices (should work)');
    try {
      const response = await fetch('https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$top=1', {
        method: 'GET',
        headers
      });

      if (response.ok) {
        const data = await response.json();
        console.log(`✅ ManagedDevices works! Found ${data.value?.length || 0} devices`);
      } else {
        const errorText = await response.text();
        console.error(`❌ ManagedDevices failed: ${response.status} ${response.statusText}`);
      }
    } catch (error) {
      console.error(`❌ ManagedDevices request failed: ${error.message}`);
    }

    // Test 5: Check app permissions directly
    console.log('\n📍 Test 5: Check service principal permissions');
    try {
      const response = await fetch(`https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${MS_CLIENT_ID}'&$expand=oauth2PermissionGrants,appRoleAssignments`, {
        method: 'GET',
        headers
      });

      if (response.ok) {
        const data = await response.json();
        if (data.value && data.value.length > 0) {
          const sp = data.value[0];
          console.log(`✅ Service Principal found: ${sp.displayName}`);
          console.log(`   App Role Assignments: ${sp.appRoleAssignments?.length || 0}`);
          
          // Check for specific Intune permissions
          const roleAssignments = sp.appRoleAssignments || [];
          const intuneRoles = roleAssignments.filter(role => 
            role.principalDisplayName && role.principalDisplayName.includes('Microsoft Graph')
          );
          console.log(`   Graph role assignments: ${intuneRoles.length}`);
        }
      } else {
        console.error(`❌ Could not check service principal: ${response.status}`);
      }
    } catch (error) {
      console.error(`❌ Service principal check failed: ${error.message}`);
    }

  } catch (error) {
    console.error(`💥 Test failed: ${error.message}`);
  }
}

testIntunePermissions();