#!/usr/bin/env node

/**
 * Simple schema test for enhanced Intune tools
 */

import { enhancedIntuneTools } from './build/tool-definitions-intune-enhanced.js';

console.log('🧪 Testing Enhanced Schema\n');

const tool = enhancedIntuneTools[0];
const schema = tool.inputSchema;

// Test simple Windows policy
const simpleWindowsPolicy = {
  platform: 'windows',
  policyType: 'Compliance',
  displayName: 'Test Policy',
  description: 'Test description'
};

console.log('Testing simple Windows policy...');
const result1 = schema.safeParse(simpleWindowsPolicy);
console.log('Simple Windows policy:', result1.success ? '✅ Valid' : '❌ Invalid');
if (!result1.success) {
  console.log('Errors:', result1.error.errors);
}

// Test Windows policy with settings
const windowsPolicyWithSettings = {
  platform: 'windows',
  policyType: 'Compliance',
  displayName: 'Test Policy With Settings',
  description: 'Test description',
  settings: {
    policyType: 'Compliance',
    settings: {
      passwordRequired: true,
      passwordMinimumLength: 8
    }
  }
};

console.log('\nTesting Windows policy with settings...');
const result2 = schema.safeParse(windowsPolicyWithSettings);
console.log('Windows policy with settings:', result2.success ? '✅ Valid' : '❌ Invalid');
if (!result2.success) {
  console.log('Errors:', result2.error.errors);
}

// Test macOS policy
const macOSPolicy = {
  platform: 'macos',
  policyType: 'Configuration',
  displayName: 'Test macOS Policy',
  description: 'Test description',
  settings: {
    policyType: 'Configuration',
    settings: {
      customConfiguration: {
        payloadFileName: 'test.mobileconfig',
        payload: 'dGVzdA=='
      }
    }
  }
};

console.log('\nTesting macOS policy...');
const result3 = schema.safeParse(macOSPolicy);
console.log('macOS policy:', result3.success ? '✅ Valid' : '❌ Invalid');
if (!result3.success) {
  console.log('Errors:', result3.error.errors);
}
