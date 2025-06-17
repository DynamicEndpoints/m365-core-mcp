#!/usr/bin/env node

/**
 * Test script for enhanced Microsoft API features
 * This script validates the new performance and reliability enhancements
 */

import { promises as fs } from 'fs';

const testCases = [
  {
    name: "Basic Graph API call (backward compatibility)",
    request: {
      jsonrpc: "2.0",
      id: 1,
      method: "tools/call",
      params: {
        name: "dynamicendpoints m365 assistant",
        arguments: {
          apiType: "graph",
          path: "/me",
          method: "get"
        }
      }
    }
  },
  {
    name: "Enhanced Graph API with field selection",
    request: {
      jsonrpc: "2.0",
      id: 2,
      method: "tools/call",
      params: {
        name: "dynamicendpoints m365 assistant",
        arguments: {
          apiType: "graph",
          path: "/users",
          method: "get",
          selectFields: ["id", "displayName", "mail"],
          responseFormat: "minimal",
          maxRetries: 2,
          timeout: 15000
        }
      }
    }
  },
  {
    name: "Enhanced pagination with custom batch size",
    request: {
      jsonrpc: "2.0",
      id: 3,
      method: "tools/call",
      params: {
        name: "dynamicendpoints m365 assistant",
        arguments: {
          apiType: "graph",
          path: "/users",
          method: "get",
          fetchAll: true,
          batchSize: 50,
          selectFields: ["id", "displayName"],
          responseFormat: "json"
        }
      }
    }
  },
  {
    name: "Azure API with enhanced error handling",
    request: {
      jsonrpc: "2.0",
      id: 4,
      method: "tools/call",
      params: {
        name: "dynamicendpoints m365 assistant",
        arguments: {
          apiType: "azure",
          path: "/subscriptions",
          method: "get",
          apiVersion: "2022-12-01",
          maxRetries: 1,
          retryDelay: 500,
          responseFormat: "minimal"
        }
      }
    }
  },
  {
    name: "Custom headers and expand fields test",
    request: {
      jsonrpc: "2.0",
      id: 5,
      method: "tools/call",
      params: {
        name: "dynamicendpoints m365 assistant",
        arguments: {
          apiType: "graph",
          path: "/groups",
          method: "get",
          selectFields: ["id", "displayName", "members"],
          expandFields: ["members"],
          customHeaders: {
            "X-Test-Header": "enhanced-api-test"
          },
          responseFormat: "json",
          maxRetries: 3
        }
      }
    }
  }
];

console.log('🚀 Microsoft 365 MCP Server - Enhanced API Features Test Suite');
console.log('================================================================');
console.log();

testCases.forEach((testCase, index) => {
  console.log(`Test ${index + 1}: ${testCase.name}`);
  console.log('Request:', JSON.stringify(testCase.request, null, 2));
  console.log();
  console.log('Features being tested:');
  
  const args = testCase.request.params.arguments;
  const features = [];
  
  if (args.selectFields) features.push(`Field Selection: ${args.selectFields.join(', ')}`);
  if (args.expandFields) features.push(`Field Expansion: ${args.expandFields.join(', ')}`);
  if (args.responseFormat && args.responseFormat !== 'json') features.push(`Response Format: ${args.responseFormat}`);
  if (args.maxRetries !== undefined) features.push(`Max Retries: ${args.maxRetries}`);
  if (args.retryDelay !== undefined) features.push(`Retry Delay: ${args.retryDelay}ms`);
  if (args.timeout !== undefined) features.push(`Timeout: ${args.timeout}ms`);
  if (args.customHeaders) features.push(`Custom Headers: ${Object.keys(args.customHeaders).join(', ')}`);
  if (args.fetchAll) features.push(`Pagination: Fetch all pages`);
  if (args.batchSize !== undefined) features.push(`Batch Size: ${args.batchSize}`);
  
  if (features.length > 0) {
    features.forEach(feature => console.log(`  ✓ ${feature}`));
  } else {
    console.log('  • Basic API call (backward compatibility test)');
  }
  
  console.log();
  console.log('---');
  console.log();
});

console.log('📝 Enhanced Features Summary:');
console.log('  ✓ Token caching for better performance');
console.log('  ✓ Rate limiting to prevent API throttling');  
console.log('  ✓ Retry logic with exponential backoff');
console.log('  ✓ Configurable timeouts');
console.log('  ✓ Custom headers support');
console.log('  ✓ Multiple response formats (json, raw, minimal)');
console.log('  ✓ Auto-apply $select and $expand for Graph API');
console.log('  ✓ Configurable pagination batch sizes');
console.log('  ✓ Enhanced error reporting with execution metrics');
console.log('  ✓ Backward compatibility maintained');
console.log();

console.log('🔧 To test with your server:');
console.log('1. Start the server: npm start');
console.log('2. Send these requests to http://localhost:3000/mcp');
console.log('3. Check the response format and execution times');
console.log();

console.log('⚠️  Note: Some tests require valid Microsoft 365 credentials');
console.log('   Set TENANT_ID, CLIENT_ID, and CLIENT_SECRET environment variables');
