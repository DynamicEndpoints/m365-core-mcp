#!/usr/bin/env node

// Quick startup test for Smithery deployment
import { spawn } from 'child_process';

console.log('Testing server startup...');
const start = Date.now();

const server = spawn('node', ['build/index.js'], {
  stdio: ['pipe', 'pipe', 'pipe'],
  env: {
    ...process.env,
    MS_TENANT_ID: 'test',
    MS_CLIENT_ID: 'test', 
    MS_CLIENT_SECRET: 'test'
  }
});

let initialized = false;

server.stdout.on('data', (data) => {
  const output = data.toString();
  console.log('Server output:', output);
  
  if (output.includes('Server running') || output.includes('Ready') || output.includes('listening')) {
    const duration = Date.now() - start;
    console.log(`✅ Server started in ${duration}ms`);
    initialized = true;
    server.kill();
  }
});

server.stderr.on('data', (data) => {
  console.log('Server stderr:', data.toString());
});

server.on('close', (code) => {
  const duration = Date.now() - start;
  if (!initialized) {
    console.log(`❌ Server exited with code ${code} after ${duration}ms`);
  } else {
    console.log(`✅ Server startup test completed in ${duration}ms`);
  }
});

// Timeout after 3 seconds
setTimeout(() => {
  if (!initialized) {
    console.log('❌ Server startup timeout (3s)');
    server.kill();
    process.exit(1);
  }
}, 3000);
