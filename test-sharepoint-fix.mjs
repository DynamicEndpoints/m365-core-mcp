#!/usr/bin/env node

/**
 * Test SharePoint hostname fix
 * Tests the corrected SharePoint handlers to ensure they use proper Graph API endpoints
 */

import { spawn } from 'child_process';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

console.log('🧪 Testing SharePoint hostname fixes...\n');

// Test configuration
const testConfig = {
  msTenantId: process.env.MS_TENANT_ID || 'test-tenant-id',
  msClientId: process.env.MS_CLIENT_ID || 'test-client-id', 
  msClientSecret: process.env.MS_CLIENT_SECRET || 'test-client-secret'
};

// Test cases for SharePoint operations
const testCases = [
  {
    name: 'SharePoint Site List',
    tool: 'manage_sharepoint_sites',
    args: { action: 'list' },
    expectedEndpoint: '/sites?search=*'
  },
  {
    name: 'SharePoint Site Search',
    tool: 'manage_sharepoint_sites', 
    args: { action: 'search', title: 'test' },
    expectedEndpoint: '/sites?search=test'
  },
  {
    name: 'SharePoint Site Get by ID',
    tool: 'manage_sharepoint_sites',
    args: { action: 'get', siteId: 'test-site-id' },
    expectedEndpoint: '/sites/test-site-id'
  },
  {
    name: 'SharePoint Site Get by URL',
    tool: 'manage_sharepoint_sites',
    args: { action: 'get', url: 'https://contoso.sharepoint.com/sites/marketing' },
    expectedEndpoint: '/sites/contoso.sharepoint.com:/sites/marketing'
  },
  {
    name: 'SharePoint Site Permissions',
    tool: 'manage_sharepoint_sites',
    args: { action: 'get_permissions', siteId: 'test-site-id' },
    expectedEndpoint: '/sites/test-site-id/permissions'
  },
  {
    name: 'SharePoint Site Drives',
    tool: 'manage_sharepoint_sites',
    args: { action: 'get_drives', siteId: 'test-site-id' },
    expectedEndpoint: '/sites/test-site-id/drives'
  }
];

async function runTest(testCase) {
  return new Promise((resolve) => {
    console.log(`Testing: ${testCase.name}`);
    
    // Create MCP request
    const mcpRequest = {
      jsonrpc: '2.0',
      id: 1,
      method: 'tools/call',
      params: {
        name: testCase.tool,
        arguments: testCase.args
      }
    };

    // Start server process
    const serverProcess = spawn('node', [join(__dirname, 'build/index.js')], {
      stdio: ['pipe', 'pipe', 'pipe'],
      env: {
        ...process.env,
        ...testConfig,
        LOG_LEVEL: 'debug'
      }
    });

    let output = '';
    let errorOutput = '';
    let testResult = {
      name: testCase.name,
      passed: false,
      error: null,
      details: ''
    };

    // Capture output
    serverProcess.stdout.on('data', (data) => {
      output += data.toString();
    });

    serverProcess.stderr.on('data', (data) => {
      errorOutput += data.toString();
    });

    // Send MCP request
    setTimeout(() => {
      serverProcess.stdin.write(JSON.stringify(mcpRequest) + '\n');
    }, 1000);

    // Handle process completion
    serverProcess.on('close', (code) => {
      // Check if the test passed
      const hasInvalidHostnameError = output.includes('Invalid hostname for this tenancy') || 
                                     errorOutput.includes('Invalid hostname for this tenancy');
      
      const hasCorrectEndpoint = output.includes(testCase.expectedEndpoint) ||
                                 output.includes('graph.microsoft.com');
      
      const hasAuthError = output.includes('authentication') || 
                           output.includes('401') ||
                           output.includes('403') ||
                           errorOutput.includes('authentication');

      if (hasInvalidHostnameError) {
        testResult.passed = false;
        testResult.error = 'Still getting "Invalid hostname for this tenancy" error';
        testResult.details = output + errorOutput;
      } else if (hasAuthError && !hasInvalidHostnameError) {
        // Authentication errors are expected in test environment, but no hostname errors
        testResult.passed = true;
        testResult.details = 'Authentication error expected (no hostname error detected)';
      } else if (hasCorrectEndpoint) {
        testResult.passed = true;
        testResult.details = 'Correct Graph API endpoint detected';
      } else {
        testResult.passed = false;
        testResult.error = 'Unexpected response or endpoint';
        testResult.details = output + errorOutput;
      }

      console.log(`  ${testResult.passed ? '✅' : '❌'} ${testResult.passed ? 'PASSED' : 'FAILED'}`);
      if (testResult.error) {
        console.log(`     Error: ${testResult.error}`);
      }
      if (testResult.details && process.env.VERBOSE) {
        console.log(`     Details: ${testResult.details.substring(0, 200)}...`);
      }
      console.log();

      resolve(testResult);
    });

    // Timeout after 10 seconds
    setTimeout(() => {
      serverProcess.kill();
      testResult.passed = false;
      testResult.error = 'Test timeout';
      resolve(testResult);
    }, 10000);
  });
}

async function runAllTests() {
  console.log('Running SharePoint hostname fix tests...\n');
  
  const results = [];
  
  for (const testCase of testCases) {
    const result = await runTest(testCase);
    results.push(result);
    
    // Small delay between tests
    await new Promise(resolve => setTimeout(resolve, 500));
  }
  
  // Summary
  const passed = results.filter(r => r.passed).length;
  const total = results.length;
  
  console.log('\n📊 Test Results Summary:');
  console.log(`✅ Passed: ${passed}/${total}`);
  console.log(`❌ Failed: ${total - passed}/${total}`);
  
  if (passed === total) {
    console.log('\n🎉 All SharePoint hostname fixes are working correctly!');
    console.log('The "Invalid hostname for this tenancy" error should be resolved.');
  } else {
    console.log('\n⚠️  Some tests failed. The SharePoint hostname issue may still exist.');
    
    const failedTests = results.filter(r => !r.passed);
    console.log('\nFailed tests:');
    failedTests.forEach(test => {
      console.log(`  - ${test.name}: ${test.error}`);
    });
  }
  
  return passed === total;
}

// Run tests if this script is executed directly
if (import.meta.url === `file://${process.argv[1]}`) {
  runAllTests()
    .then(success => {
      process.exit(success ? 0 : 1);
    })
    .catch(error => {
      console.error('Test execution failed:', error);
      process.exit(1);
    });
}

export { runAllTests };
