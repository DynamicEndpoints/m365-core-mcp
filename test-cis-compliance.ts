// Test script for CIS compliance functionality
import { handleCISCompliance } from './src/handlers/compliance-handler.js';
import { CISComplianceArgs } from './src/types/compliance-types.js';

// Mock Microsoft Graph Client for testing
const mockGraphClient = {
  api: (path: string) => ({
    get: async () => {
      console.log(`Mock API call: GET ${path}`);
      
      // Mock responses for different API endpoints
      if (path === '/policies/authenticationMethodsPolicy') {
        return {
          systemCredentialPreferences: {
            excludeTargets: [],
            includeTargets: []
          },
          authenticationMethodConfigurations: []
        };
      }
      
      if (path === '/identity/conditionalAccess/policies') {
        return {
          value: [
            {
              id: 'test-policy-1',
              displayName: 'Block legacy authentication',
              state: 'enabled',
              conditions: {
                clientAppTypes: ['exchangeActiveSync', 'other']
              },
              grantControls: {
                operator: 'OR',
                builtInControls: ['block']
              }
            }
          ]
        };
      }
      
      if (path === '/policies/identitySecurityDefaultsEnforcementPolicy') {
        return {
          isEnabled: true
        };
      }
      
      if (path === '/users') {
        return {
          value: [
            {
              id: 'user1',
              displayName: 'Test User 1',
              userPrincipalName: 'user1@contoso.com'
            }
          ]
        };
      }
      
      if (path === '/organization') {
        return {
          value: [
            {
              id: 'org1',
              displayName: 'Test Organization'
            }
          ]
        };
      }
      
      if (path === '/directoryRoles') {
        return {
          value: [
            {
              id: 'role1',
              displayName: 'Global Administrator'
            }
          ]
        };
      }
      
      return { value: [] };
    },
    top: (count: number) => ({
      get: async () => ({ value: [] })
    }),
    select: (fields: string) => ({
      top: (count: number) => ({
        get: async () => ({ value: [] })
      })
    }),
    version: (version: string) => ({
      get: async () => ({ value: [] })
    })
  })
} as any;

async function testCISCompliance() {
  console.log('Testing CIS Compliance Tool...\n');
  
  // Test 1: Get CIS Benchmark
  console.log('1. Testing get_benchmark action:');
  try {
    const benchmarkArgs: CISComplianceArgs = {
      action: 'get_benchmark',
      benchmark: 'office365'
    };
    
    const benchmarkResult = await handleCISCompliance(mockGraphClient, benchmarkArgs);
    console.log('✅ Get benchmark test passed');
    console.log('Benchmark details:', JSON.stringify(benchmarkResult.content[0].text.substring(0, 200) + '...', null, 2));
  } catch (error) {
    console.log('❌ Get benchmark test failed:', error);
  }
  
  console.log('\n2. Testing assess action:');
  try {
    const assessArgs: CISComplianceArgs = {
      action: 'assess',
      benchmark: 'office365',
      implementationGroup: '1',
      settings: {
        automated: true,
        generateRemediation: true,
        includeEvidence: true,
        riskPrioritization: true
      }
    };
    
    const assessResult = await handleCISCompliance(mockGraphClient, assessArgs);
    console.log('✅ Assessment test passed');
    console.log('Assessment summary:', JSON.stringify(assessResult.content[0].text.substring(0, 200) + '...', null, 2));
  } catch (error) {
    console.log('❌ Assessment test failed:', error);
  }
  
  console.log('\n3. Testing generate_report action:');
  try {
    const reportArgs: CISComplianceArgs = {
      action: 'generate_report',
      benchmark: 'office365',
      implementationGroup: '1'
    };
    
    const reportResult = await handleCISCompliance(mockGraphClient, reportArgs);
    console.log('✅ Report generation test passed');
    console.log('Report summary:', JSON.stringify(reportResult.content[0].text.substring(0, 200) + '...', null, 2));
  } catch (error) {
    console.log('❌ Report generation test failed:', error);
  }
  
  console.log('\n4. Testing configure_monitoring action:');
  try {
    const monitoringArgs: CISComplianceArgs = {
      action: 'configure_monitoring',
      benchmark: 'office365',
      settings: {
        automated: true
      }
    };
    
    const monitoringResult = await handleCISCompliance(mockGraphClient, monitoringArgs);
    console.log('✅ Configure monitoring test passed');
    console.log('Monitoring config:', JSON.stringify(monitoringResult.content[0].text.substring(0, 200) + '...', null, 2));
  } catch (error) {
    console.log('❌ Configure monitoring test failed:', error);
  }
  
  console.log('\n5. Testing remediate action:');
  try {
    const remediateArgs: CISComplianceArgs = {
      action: 'remediate',
      benchmark: 'office365',
      controlIds: ['CIS-2.1'],
      settings: {
        automated: true
      }
    };
    
    const remediateResult = await handleCISCompliance(mockGraphClient, remediateArgs);
    console.log('✅ Remediation test passed');
    console.log('Remediation result:', JSON.stringify(remediateResult.content[0].text.substring(0, 200) + '...', null, 2));
  } catch (error) {
    console.log('❌ Remediation test failed:', error);
  }
  
  console.log('\n🎉 CIS Compliance tool testing completed!');
}

// Run the test
testCISCompliance().catch(console.error);
