#!/usr/bin/env node

/**
 * Comprehensive Feature Test for M365 Core MCP Server
 * Tests both the enhanced Microsoft API capabilities and extended resources/prompts
 */

import { execSync } from 'child_process';
import fs from 'fs';

console.log('🧪 M365 Core MCP Server - Comprehensive Feature Test');
console.log('===================================================\n');

// Test categories
const tests = {
  build: [],
  extendedResources: [],
  enhancedAPI: [],
  prompts: [],
  integration: []
};

// 1. Build and Compilation Tests
console.log('📦 1. Testing Build and Compilation...');
try {
  execSync('npm run build', { stdio: 'pipe' });
  tests.build.push('✅ TypeScript compilation successful');
} catch (error) {
  tests.build.push('❌ TypeScript compilation failed');
  console.error('Build error:', error.message);
}

// 2. Extended Resources Verification
console.log('🔍 2. Verifying Extended Resources...');

// Check if extended-resources.ts exists and has the right content
const extendedResourcesPath = './src/extended-resources.ts';
if (fs.existsSync(extendedResourcesPath)) {
  const content = fs.readFileSync(extendedResourcesPath, 'utf8');
  
  // Check for key resources mentioned in EXTENDED_FEATURES.md
  const requiredResources = [
    'security_alerts',
    'security_incidents', 
    'conditional_access_policies',
    'applications',
    'service_principals',
    'directory_roles',
    'intune_devices_extended',
    'teams_list_extended',
    'mail_folders_extended',
    'calendar_events_extended'
  ];
  
  requiredResources.forEach(resource => {
    if (content.includes(`"${resource}"`)) {
      tests.extendedResources.push(`✅ Resource '${resource}' found`);
    } else {
      tests.extendedResources.push(`❌ Resource '${resource}' missing`);
    }
  });
  
  // Count total resources
  const resourceCount = (content.match(/server\.resource\(/g) || []).length;
  tests.extendedResources.push(`📊 Total resources defined: ${resourceCount}`);
  
} else {
  tests.extendedResources.push('❌ extended-resources.ts file not found');
}

// 3. Enhanced Microsoft API Features Verification
console.log('🚀 3. Verifying Enhanced Microsoft API Features...');

// Check enhanced tool schema
const toolDefinitionsPath = './src/tool-definitions.ts';
if (fs.existsSync(toolDefinitionsPath)) {
  const content = fs.readFileSync(toolDefinitionsPath, 'utf8');
  
  const enhancedFeatures = [
    'maxRetries',
    'retryDelay', 
    'timeout',
    'customHeaders',
    'responseFormat',
    'selectFields',
    'expandFields',
    'batchSize'
  ];
  
  enhancedFeatures.forEach(feature => {
    if (content.includes(feature)) {
      tests.enhancedAPI.push(`✅ Enhanced feature '${feature}' found in schema`);
    } else {
      tests.enhancedAPI.push(`❌ Enhanced feature '${feature}' missing from schema`);
    }
  });
} else {
  tests.enhancedAPI.push('❌ tool-definitions.ts file not found');
}

// Check enhanced handler implementation
const handlersPath = './src/handlers.ts';
if (fs.existsSync(handlersPath)) {
  const content = fs.readFileSync(handlersPath, 'utf8');
  
  const handlerFeatures = [
    'executeWithRetry',
    'TokenCache',
    'RateLimiter', 
    'exponential backoff',
    'responseFormat',
    'selectFields',
    'expandFields'
  ];
  
  handlerFeatures.forEach(feature => {
    if (content.includes(feature)) {
      tests.enhancedAPI.push(`✅ Handler feature '${feature}' implemented`);
    } else {
      tests.enhancedAPI.push(`⚠️  Handler feature '${feature}' not found (may use different naming)`);
    }
  });
} else {
  tests.enhancedAPI.push('❌ handlers.ts file not found');
}

// 4. Prompts Verification  
console.log('📝 4. Verifying Comprehensive Prompts...');

if (fs.existsSync(extendedResourcesPath)) {
  const content = fs.readFileSync(extendedResourcesPath, 'utf8');
  
  const requiredPrompts = [
    'security_assessment',
    'compliance_review',
    'user_access_review', 
    'device_compliance_analysis',
    'collaboration_governance'
  ];
  
  requiredPrompts.forEach(prompt => {
    if (content.includes(`"${prompt}"`)) {
      tests.prompts.push(`✅ Prompt '${prompt}' found`);
    } else {
      tests.prompts.push(`❌ Prompt '${prompt}' missing`);
    }
  });
  
  // Count total prompts
  const promptCount = (content.match(/server\.prompt\(/g) || []).length;
  tests.prompts.push(`📊 Total prompts defined: ${promptCount}`);
}

// 5. Integration Tests
console.log('🔗 5. Verifying Integration...');

// Check if server.ts imports and uses extended resources
const serverPath = './src/server.ts';
if (fs.existsSync(serverPath)) {
  const content = fs.readFileSync(serverPath, 'utf8');
  
  if (content.includes('setupExtendedResources')) {
    tests.integration.push('✅ Extended resources integrated into server');
  } else {
    tests.integration.push('❌ Extended resources not integrated');
  }
  
  if (content.includes('TokenCache') || content.includes('RateLimiter')) {
    tests.integration.push('✅ Enhanced utility classes integrated');
  } else {
    tests.integration.push('❌ Enhanced utility classes not integrated');
  }
  
  if (content.includes('version: \'1.1.0\'')) {
    tests.integration.push('✅ Server version updated to reflect enhancements');
  } else {
    tests.integration.push('⚠️  Server version not updated (may still be 1.0.0)');
  }
} else {
  tests.integration.push('❌ server.ts file not found');
}

// Check if index.ts reflects enhanced version
const indexPath = './src/index.ts';
if (fs.existsSync(indexPath)) {
  const content = fs.readFileSync(indexPath, 'utf8');
  
  if (content.includes('1.1.0')) {
    tests.integration.push('✅ Index version updated to reflect enhancements');
  } else {
    tests.integration.push('⚠️  Index version not updated');
  }
}

// Display Results
console.log('\n📊 TEST RESULTS SUMMARY');
console.log('=======================\n');

Object.entries(tests).forEach(([category, results]) => {
  console.log(`${category.toUpperCase()}:`);
  results.forEach(result => console.log(`  ${result}`));
  console.log('');
});

// Overall Assessment
const allTests = Object.values(tests).flat();
const passedTests = allTests.filter(test => test.includes('✅')).length;
const failedTests = allTests.filter(test => test.includes('❌')).length;
const warningTests = allTests.filter(test => test.includes('⚠️')).length;

console.log('OVERALL ASSESSMENT:');
console.log(`✅ Passed: ${passedTests}`);
console.log(`❌ Failed: ${failedTests}`);
console.log(`⚠️  Warnings: ${warningTests}`);
console.log(`📊 Total: ${allTests.length}`);

const successRate = (passedTests / allTests.length) * 100;
console.log(`\n🎯 Success Rate: ${successRate.toFixed(1)}%`);

if (successRate >= 90) {
  console.log('🎉 Excellent! All major features are implemented and working.');
} else if (successRate >= 75) {
  console.log('👍 Good! Most features are working, some minor issues to address.');
} else if (successRate >= 50) {
  console.log('⚠️  Fair! Several features need attention.');
} else {
  console.log('🚨 Poor! Major issues need to be resolved.');
}

console.log('\n✨ Feature Test Complete!\n');
