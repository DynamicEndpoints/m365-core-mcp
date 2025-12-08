#!/usr/bin/env node

/**
 * Test script to list all available MCP tools
 * This helps verify that document generation tools are registered
 */

import { spawn } from 'child_process';
import { resolve } from 'path';

async function listTools() {
  console.log('🔍 Listing all registered MCP tools...\n');

  const serverPath = resolve('build/index.js');
  
  return new Promise((resolvePromise, reject) => {
    const server = spawn('node', [serverPath], {
      stdio: ['pipe', 'pipe', 'pipe'],
      env: {
        ...process.env,
        MS_TENANT_ID: process.env.MS_TENANT_ID || 'test-tenant',
        MS_CLIENT_ID: process.env.MS_CLIENT_ID || 'test-client',
        MS_CLIENT_SECRET: process.env.MS_CLIENT_SECRET || 'test-secret'
      }
    });

    let output = '';
    let errorOutput = '';

    server.stdout.on('data', (data) => {
      output += data.toString();
    });

    server.stderr.on('data', (data) => {
      errorOutput += data.toString();
    });

    // Send initialize request
    const initRequest = {
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: {
        protocolVersion: '2024-11-05',
        capabilities: {},
        clientInfo: {
          name: 'test-client',
          version: '1.0.0'
        }
      }
    };

    server.stdin.write(JSON.stringify(initRequest) + '\n');

    // Send tools/list request after a short delay
    setTimeout(() => {
      const listRequest = {
        jsonrpc: '2.0',
        id: 2,
        method: 'tools/list',
        params: {}
      };
      server.stdin.write(JSON.stringify(listRequest) + '\n');
    }, 1000);

    // Process output after another delay
    setTimeout(() => {
      server.kill();
      
      if (errorOutput) {
        console.error('❌ Server Error:', errorOutput);
        reject(new Error(errorOutput));
        return;
      }

      // Parse JSON-RPC responses
      const lines = output.split('\n').filter(line => line.trim());
      const responses = lines.map(line => {
        try {
          return JSON.parse(line);
        } catch (e) {
          return null;
        }
      }).filter(r => r !== null);

      // Find tools/list response
      const toolsResponse = responses.find(r => r.id === 2 && r.result?.tools);
      
      if (toolsResponse) {
        const tools = toolsResponse.result.tools;
        console.log(`✅ Found ${tools.length} registered tools\n`);
        
        // Filter for document generation tools
        const docGenTools = tools.filter(t => 
          t.name.includes('generate_') || t.name.includes('oauth_')
        );
        
        console.log('📄 Document Generation Tools:');
        console.log('━'.repeat(80));
        
        if (docGenTools.length === 0) {
          console.log('⚠️  No document generation tools found!\n');
        } else {
          docGenTools.forEach(tool => {
            console.log(`\n✓ ${tool.name}`);
            if (tool.description) {
              console.log(`  ${tool.description.substring(0, 100)}...`);
            }
          });
          console.log('\n');
        }
        
        // List all tools for reference
        console.log('\n📋 All Available Tools:');
        console.log('━'.repeat(80));
        tools.forEach((tool, idx) => {
          console.log(`${idx + 1}. ${tool.name}`);
        });
        
        resolvePromise();
      } else {
        console.error('❌ No tools/list response found');
        console.log('Raw output:', output);
        reject(new Error('No tools response'));
      }
    }, 3000);
  });
}

listTools().catch(err => {
  console.error('Failed:', err.message);
  process.exit(1);
});
