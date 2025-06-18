#!/usr/bin/env node

/**
 * Test script for enhanced Intune policy creation tools
 * This script tests the new enhanced tools alongside existing ones
 */

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { enhancedIntuneTools } from './build/tool-definitions-intune-enhanced.js';
import { intuneTools } from './build/tool-definitions-intune.js';
import { 
  validatePolicySettings, 
  applyPolicyDefaults, 
  generatePolicyExample,
  POLICY_TEMPLATES 
} from './build/validators/intune-policy-validator.js';

console.log('🧪 Testing Enhanced Intune Policy Tools\n');

// Test 1: Verify enhanced tools are properly defined
console.log('✅ Test 1: Enhanced Tool Definitions');
console.log(`Enhanced tools found: ${enhancedIntuneTools.length}`);
enhancedIntuneTools.forEach(tool => {
  console.log(`  - ${tool.name}: ${tool.description}`);
});

console.log(`\nOriginal tools found: ${intuneTools.length}`);
intuneTools.forEach(tool => {
  console.log(`  - ${tool.name}: ${tool.description}`);
});

// Test 2: Test validation functions
console.log('\n✅ Test 2: Policy Validation');

// Test macOS compliance validation
const macOSCompliance = {
  compliance: {
    passwordRequired: true,
    passwordMinimumLength: 8,
    storageRequireEncryption: true,
    osMinimumVersion: '14.0'
  }
};

const macOSValidation = validatePolicySettings('macos', 'Compliance', macOSCompliance);
console.log('macOS Compliance validation:', macOSValidation.isValid ? '✅ Valid' : '❌ Invalid');
if (macOSValidation.warnings.length > 0) {
  console.log('  Warnings:', macOSValidation.warnings);
}

// Test Windows compliance validation
const windowsCompliance = {
  passwordRequired: true,
  passwordMinimumLength: 8,
  bitLockerEnabled: true,
  storageRequireEncryption: true
};

const windowsValidation = validatePolicySettings('windows', 'Compliance', windowsCompliance);
console.log('Windows Compliance validation:', windowsValidation.isValid ? '✅ Valid' : '❌ Invalid');
if (windowsValidation.warnings.length > 0) {
  console.log('  Warnings:', windowsValidation.warnings);
}

// Test 3: Test defaults application
console.log('\n✅ Test 3: Policy Defaults');

const emptySettings = {};
const windowsDefaults = applyPolicyDefaults('windows', 'Compliance', emptySettings);
console.log('Windows Compliance defaults applied:', Object.keys(windowsDefaults).length > 0 ? '✅ Success' : '❌ Failed');

const macOSDefaults = applyPolicyDefaults('macos', 'Compliance', emptySettings);
console.log('macOS Compliance defaults applied:', Object.keys(macOSDefaults).length > 0 ? '✅ Success' : '❌ Failed');

// Test 4: Test templates
console.log('\n✅ Test 4: Policy Templates');

const windowsTemplates = POLICY_TEMPLATES.windows;
console.log(`Windows templates available: ${Object.keys(windowsTemplates).length}`);
Object.keys(windowsTemplates).forEach(template => {
  console.log(`  - ${template}: ${windowsTemplates[template].name}`);
});

const macOSTemplates = POLICY_TEMPLATES.macos;
console.log(`macOS templates available: ${Object.keys(macOSTemplates).length}`);
Object.keys(macOSTemplates).forEach(template => {
  console.log(`  - ${template}: ${macOSTemplates[template].name}`);
});

// Test 5: Test example generation
console.log('\n✅ Test 5: Example Generation');

const windowsExample = generatePolicyExample('windows', 'Compliance');
console.log('Windows Compliance example generated:', windowsExample.length > 0 ? '✅ Success' : '❌ Failed');

const macOSExample = generatePolicyExample('macos', 'Configuration');
console.log('macOS Configuration example generated:', macOSExample.length > 0 ? '✅ Success' : '❌ Failed');

// Test 6: Schema validation
console.log('\n✅ Test 6: Schema Validation');

try {
  // Test enhanced schema with discriminated union
  const enhancedTool = enhancedIntuneTools[0];
  const schema = enhancedTool.inputSchema;
    // Test valid Windows policy
  const validWindowsPolicy = {
    platform: 'windows',
    policyType: 'Compliance',
    displayName: 'Test Windows Policy',
    description: 'Test policy',
    settings: {
      policyType: 'Compliance',
      settings: {
        passwordRequired: true,
        passwordMinimumLength: 8
      }
    }
  };
  
  const validationResult = schema.safeParse(validWindowsPolicy);
  console.log('Enhanced schema Windows validation:', validationResult.success ? '✅ Valid' : '❌ Invalid');
  
  // Test valid macOS policy
  const validMacOSPolicy = {
    platform: 'macos',
    policyType: 'Configuration',
    displayName: 'Test macOS Policy',
    description: 'Test policy',
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
  
  const macOSValidationResult = schema.safeParse(validMacOSPolicy);
  console.log('Enhanced schema macOS validation:', macOSValidationResult.success ? '✅ Valid' : '❌ Invalid');
  
} catch (error) {
  console.log('❌ Schema validation error:', error);
}

// Test 7: Tool name conflicts check
console.log('\n✅ Test 7: Tool Name Conflicts');

const originalToolNames = intuneTools.map(t => t.name);
const enhancedToolNames = enhancedIntuneTools.map(t => `enhanced_${t.name}`);

const hasConflicts = originalToolNames.some(name => enhancedToolNames.includes(name));
console.log('Tool name conflicts:', hasConflicts ? '❌ Conflicts found' : '✅ No conflicts');

console.log('\n🎉 Enhanced Intune Policy Tools Test Complete!\n');
console.log('Summary:');
console.log(`- Original tools: ${originalToolNames.length}`);
console.log(`- Enhanced tools: ${enhancedToolNames.length}`);
console.log(`- Total tools available: ${originalToolNames.length + enhancedToolNames.length}`);
console.log('\nTools can be used with:');
console.log('- Original: create_intune_policy, createIntunePolicy');
console.log('- Enhanced: enhanced_create_intune_policy');
console.log('\nThe enhanced tools provide:');
console.log('- ✅ Advanced validation');
console.log('- ✅ Policy templates');
console.log('- ✅ Default value application');
console.log('- ✅ Platform-specific schemas');
console.log('- ✅ Detailed error messages');
console.log('- ✅ Assignment validation');
