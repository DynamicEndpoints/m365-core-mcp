#!/usr/bin/env node

/**
 * Test script to verify HTTP transport functionality for M365 Core MCP Server
 * This script tests the HTTP endpoints and MCP protocol over HTTP
 */

import fetch from 'node-fetch';
import { spawn } from 'child_process';
import { setTimeout } from 'timers/promises';

const SERVER_PORT = 8080;
const SERVER_URL = `http://localhost:${SERVER_PORT}`;
const MCP_ENDPOINT = `${SERVER_URL}/mcp`;

let serverProcess = null;

async function startServer() {
  console.log('🚀 Starting M365 Core MCP Server in HTTP mode...');
  
  serverProcess = spawn('node', ['build/index.js'], {
    env: {
      ...process.env,
      USE_HTTP: 'true',
      PORT: SERVER_PORT.toString(),
      STATELESS: 'true', // Use stateless mode for testing
      LOG_LEVEL: 'info'
    },
    stdio: ['pipe', 'pipe', 'pipe']
  });

  serverProcess.stdout.on('data', (data) => {
    console.log(`[SERVER] ${data.toString().trim()}`);
  });

  serverProcess.stderr.on('data', (data) => {
    console.error(`[SERVER ERROR] ${data.toString().trim()}`);
  });

  // Wait for server to start
  await setTimeout(3000);
  console.log('✅ Server should be running now');
}

async function stopServer() {
  if (serverProcess) {
    console.log('🛑 Stopping server...');
    serverProcess.kill('SIGTERM');
    await setTimeout(1000);
  }
}

async function testHealthEndpoint() {
  console.log('\n📋 Test 1: Health Check Endpoint');
  try {
    const response = await fetch(`${SERVER_URL}/health`);
    const data = await response.json();
    
    console.log('   ✅ Health endpoint accessible');
    console.log(`   📊 Server: ${data.server}`);
    console.log(`   📊 Version: ${data.version}`);
    console.log(`   📊 Status: ${data.status}`);
    console.log(`   📊 Capabilities: ${Object.keys(data.capabilities).join(', ')}`);
    
    return true;
  } catch (error) {
    console.log(`   ❌ Health endpoint failed: ${error.message}`);
    return false;
  }
}

async function testCapabilitiesEndpoint() {
  console.log('\n📋 Test 2: Capabilities Endpoint');
  try {
    const response = await fetch(`${SERVER_URL}/capabilities`);
    const data = await response.json();
    
    console.log('   ✅ Capabilities endpoint accessible');
    console.log(`   📊 Protocol: ${data.protocol}`);
    console.log(`   📊 Tools available: ${data.tools.length}`);
    console.log(`   📊 Sample tools: ${data.tools.slice(0, 5).join(', ')}`);
    
    return true;
  } catch (error) {
    console.log(`   ❌ Capabilities endpoint failed: ${error.message}`);
    return false;
  }
}

async function testMCPInitialize() {
  console.log('\n📋 Test 3: MCP Initialize Request');
  try {
    const initializeRequest = {
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: {
        protocolVersion: '2024-11-05',
        capabilities: {
          roots: {
            listChanged: true
          },
          sampling: {}
        },
        clientInfo: {
          name: 'test-client',
          version: '1.0.0'
        }
      }
    };

    const response = await fetch(MCP_ENDPOINT, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json, text/event-stream',
      },
      body: JSON.stringify(initializeRequest)
    });

    const data = await response.json();
    
    if (data.result && data.result.capabilities) {
      console.log('   ✅ MCP Initialize successful');
      console.log(`   📊 Server capabilities: ${Object.keys(data.result.capabilities).join(', ')}`);
      console.log(`   📊 Server info: ${data.result.serverInfo?.name || 'Unknown'}`);
      return true;
    } else {
      console.log('   ❌ MCP Initialize failed - no capabilities in response');
      console.log(`   📊 Response: ${JSON.stringify(data, null, 2)}`);
      return false;
    }
  } catch (error) {
    console.log(`   ❌ MCP Initialize failed: ${error.message}`);
    return false;
  }
}

async function testMCPListTools() {
  console.log('\n📋 Test 4: MCP List Tools Request');
  try {
    const listToolsRequest = {
      jsonrpc: '2.0',
      id: 2,
      method: 'tools/list',
      params: {}
    };

    const response = await fetch(MCP_ENDPOINT, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json, text/event-stream',
      },
      body: JSON.stringify(listToolsRequest)
    });

    const data = await response.json();
    
    if (data.result && data.result.tools) {
      console.log('   ✅ MCP List Tools successful');
      console.log(`   📊 Tools found: ${data.result.tools.length}`);
      console.log(`   📊 Sample tools: ${data.result.tools.slice(0, 3).map(t => t.name).join(', ')}`);
      
      // Check for specific M365 tools
      const m365Tools = data.result.tools.filter(t => 
        t.name.includes('manage_') || t.name.includes('m365') || t.name.includes('azure')
      );
      console.log(`   📊 M365-specific tools: ${m365Tools.length}`);
      
      return true;
    } else {
      console.log('   ❌ MCP List Tools failed - no tools in response');
      console.log(`   📊 Response: ${JSON.stringify(data, null, 2)}`);
      return false;
    }
  } catch (error) {
    console.log(`   ❌ MCP List Tools failed: ${error.message}`);
    return false;
  }
}

async function testMCPHealthCheckTool() {
  console.log('\n📋 Test 5: MCP Health Check Tool (No Auth Required)');
  try {
    const toolCallRequest = {
      jsonrpc: '2.0',
      id: 3,
      method: 'tools/call',
      params: {
        name: 'health_check',
        arguments: {}
      }
    };

    const response = await fetch(MCP_ENDPOINT, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json, text/event-stream',
      },
      body: JSON.stringify(toolCallRequest)
    });

    const data = await response.json();
    
    if (data.result && data.result.content) {
      console.log('   ✅ Health Check Tool successful');
      const content = data.result.content[0];
      if (content && content.text) {
        const healthData = JSON.parse(content.text.split('\n\n')[1]);
        console.log(`   📊 Server Status: ${healthData.serverStatus}`);
        console.log(`   📊 Auth Configured: ${healthData.authentication.configured}`);
        console.log(`   📊 Auth Status: ${healthData.authentication.status}`);
      }
      return true;
    } else {
      console.log('   ❌ Health Check Tool failed');
      console.log(`   📊 Response: ${JSON.stringify(data, null, 2)}`);
      return false;
    }
  } catch (error) {
    console.log(`   ❌ Health Check Tool failed: ${error.message}`);
    return false;
  }
}

async function testCORSHeaders() {
  console.log('\n📋 Test 6: CORS Headers');
  try {
    const response = await fetch(MCP_ENDPOINT, {
      method: 'OPTIONS',
      headers: {
        'Origin': 'https://example.com',
        'Access-Control-Request-Method': 'POST',
        'Access-Control-Request-Headers': 'Content-Type'
      }
    });

    const corsHeaders = {
      'Access-Control-Allow-Origin': response.headers.get('Access-Control-Allow-Origin'),
      'Access-Control-Allow-Methods': response.headers.get('Access-Control-Allow-Methods'),
      'Access-Control-Allow-Headers': response.headers.get('Access-Control-Allow-Headers')
    };

    console.log('   ✅ CORS preflight successful');
    console.log(`   📊 Allow Origin: ${corsHeaders['Access-Control-Allow-Origin']}`);
    console.log(`   📊 Allow Methods: ${corsHeaders['Access-Control-Allow-Methods']}`);
    console.log(`   📊 Allow Headers: ${corsHeaders['Access-Control-Allow-Headers']}`);
    
    return response.status === 200;
  } catch (error) {
    console.log(`   ❌ CORS test failed: ${error.message}`);
    return false;
  }
}

async function runAllTests() {
  console.log('🧪 M365 Core MCP Server HTTP Transport Test Suite\n');

  try {
    // Start the server
    await startServer();

    // Run tests
    const results = [];
    results.push(await testHealthEndpoint());
    results.push(await testCapabilitiesEndpoint());
    results.push(await testMCPInitialize());
    results.push(await testMCPListTools());
    results.push(await testMCPHealthCheckTool());
    results.push(await testCORSHeaders());

    // Summary
    const passed = results.filter(r => r).length;
    const total = results.length;
    
    console.log('\n📊 Test Results Summary:');
    console.log(`   ✅ Passed: ${passed}/${total}`);
    console.log(`   ❌ Failed: ${total - passed}/${total}`);
    
    if (passed === total) {
      console.log('\n🎉 All tests passed! HTTP transport is working correctly.');
      console.log('\n🚀 Your M365 Core MCP Server is ready for HTTP deployment!');
      console.log('\n📋 Next steps:');
      console.log('   1. Configure your Azure AD credentials');
      console.log('   2. Deploy to Smithery or your preferred platform');
      console.log('   3. Test with real M365 operations');
    } else {
      console.log('\n⚠️  Some tests failed. Please check the server configuration.');
    }

  } catch (error) {
    console.error('💥 Test suite crashed:', error);
  } finally {
    await stopServer();
  }
}

// Handle cleanup on exit
process.on('SIGINT', async () => {
  console.log('\n🛑 Received SIGINT, cleaning up...');
  await stopServer();
  process.exit(0);
});

process.on('SIGTERM', async () => {
  console.log('\n🛑 Received SIGTERM, cleaning up...');
  await stopServer();
  process.exit(0);
});

// Run the tests
runAllTests().catch(error => {
  console.error('💥 Test execution failed:', error);
  process.exit(1);
});
