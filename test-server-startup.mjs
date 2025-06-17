#!/usr/bin/env node

// Set dummy environment variables to prevent validation errors
process.env.MS_TENANT_ID = 'test-tenant-id';
process.env.MS_CLIENT_ID = 'test-client-id';  
process.env.MS_CLIENT_SECRET = 'test-client-secret';

// Quick test to verify the server can start
import { M365CoreServer } from './build/server.js';

console.log('Testing M365CoreServer instantiation...');

try {
  const server = new M365CoreServer();
  console.log('✅ Server instance created successfully');
  console.log('✅ Server name:', server.server.name);
  console.log('✅ Server capabilities:', JSON.stringify(server.server.capabilities, null, 2));
  
  // Test if the server has tools registered
  console.log('✅ Server instance is ready');
  process.exit(0);
} catch (error) {
  console.error('❌ Server instantiation failed:', error.message);
  console.error('Stack trace:', error.stack);
  process.exit(1);
}
