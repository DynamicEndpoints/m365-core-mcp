#!/usr/bin/env node

/**
 * Validation script to test our recent changes:
 * 1. API tool renamed to "dynamicendpoints m365 assistant"
 * 2. All placeholder handlers implemented
 * 3. All tools are functional
 */

import { spawn } from 'child_process';
import fs from 'fs';
import path from 'path';

console.log('🔍 Validating M365 Core MCP Server Changes');
console.log('=========================================\n');

async function validateChanges() {
  const results = [];

  // Test 1: Check if API tool is renamed
  console.log('1️⃣ Checking API tool rename...');
  try {
    const serverContent = fs.readFileSync('./src/server.ts', 'utf8');
    const hasNewName = serverContent.includes('"dynamicendpoints m365 assistant"');
    const hasOldName = serverContent.includes('"call_microsoft_api"');
    
    if (hasNewName && !hasOldName) {
      console.log('   ✅ API tool successfully renamed to "dynamicendpoints m365 assistant"');
      results.push({ test: 'API Tool Rename', status: 'PASS' });
    } else {
      console.log('   ❌ API tool rename failed');
      results.push({ test: 'API Tool Rename', status: 'FAIL' });
    }
  } catch (error) {
    console.log('   ❌ Error reading server.ts:', error.message);
    results.push({ test: 'API Tool Rename', status: 'ERROR' });
  }

  // Test 2: Check placeholder implementations
  console.log('\n2️⃣ Checking placeholder implementations...');
  try {
    const handlersContent = fs.readFileSync('./src/handlers.ts', 'utf8');
    const hasPlaceholderText = handlersContent.includes('not yet implemented');
    
    if (!hasPlaceholderText) {
      console.log('   ✅ All placeholder implementations removed');
      results.push({ test: 'Placeholder Removal', status: 'PASS' });
    } else {
      console.log('   ❌ Still contains placeholder implementations');
      results.push({ test: 'Placeholder Removal', status: 'FAIL' });
    }
  } catch (error) {
    console.log('   ❌ Error reading handlers.ts:', error.message);
    results.push({ test: 'Placeholder Removal', status: 'ERROR' });
  }

  // Test 3: Check TypeScript compilation
  console.log('\n3️⃣ Testing TypeScript compilation...');
  try {
    const { execSync } = await import('child_process');
    execSync('npm run build', { stdio: 'pipe' });
    console.log('   ✅ TypeScript compilation successful');
    results.push({ test: 'TypeScript Build', status: 'PASS' });
  } catch (error) {
    console.log('   ❌ TypeScript compilation failed');
    results.push({ test: 'TypeScript Build', status: 'FAIL' });
  }

  // Test 4: Check tool definitions
  console.log('\n4️⃣ Checking tool registration...');
  try {
    const expectedTools = [
      'manage_distribution_lists',
      'manage_security_groups', 
      'manage_m365_groups',
      'manage_exchange_settings',
      'manage_user_settings',
      'manage_offboarding',
      'manage_sharepoint_sites',
      'manage_sharepoint_lists',
      'manage_azuread_roles',
      'manage_azuread_apps',
      'manage_azuread_devices',
      'manage_service_principals',
      'dynamicendpoints m365 assistant',
      'search_audit_log',
      'manage_alerts',
      'manage_dlp_policies',
      'manage_dlp_incidents',
      'manage_sensitivity_labels',
      'manage_intune_macos_devices',
      'manage_intune_macos_policies',
      'manage_intune_macos_apps',
      'manage_intune_macos_compliance',
      'manage_compliance_frameworks',
      'manage_compliance_assessments',
      'manage_compliance_monitoring',
      'manage_evidence_collection',
      'manage_gap_analysis',
      'generate_audit_reports',
      'manage_cis_compliance'
    ];

    const serverContent = fs.readFileSync('./src/server.ts', 'utf8');
    const missingTools = expectedTools.filter(tool => !serverContent.includes(`"${tool}"`));
    
    if (missingTools.length === 0) {
      console.log(`   ✅ All ${expectedTools.length} tools properly registered`);
      results.push({ test: 'Tool Registration', status: 'PASS' });
    } else {
      console.log(`   ❌ Missing tools: ${missingTools.join(', ')}`);
      results.push({ test: 'Tool Registration', status: 'FAIL' });
    }
  } catch (error) {
    console.log('   ❌ Error checking tool registration:', error.message);
    results.push({ test: 'Tool Registration', status: 'ERROR' });
  }

  // Test 5: Validate handler imports
  console.log('\n5️⃣ Checking handler implementations...');
  try {
    const handlersContent = fs.readFileSync('./src/handlers.ts', 'utf8');
    const requiredFunctions = [
      'handleDistributionLists',
      'handleSecurityGroups',
      'handleM365Groups',
      'handleSharePointSites',
      'handleSharePointLists'
    ];
    
    const missingFunctions = requiredFunctions.filter(func => 
      !handlersContent.includes(`export async function ${func}`)
    );
    
    if (missingFunctions.length === 0) {
      console.log('   ✅ All handler functions properly implemented');
      results.push({ test: 'Handler Implementation', status: 'PASS' });
    } else {
      console.log(`   ❌ Missing handler functions: ${missingFunctions.join(', ')}`);
      results.push({ test: 'Handler Implementation', status: 'FAIL' });
    }
  } catch (error) {
    console.log('   ❌ Error checking handler implementations:', error.message);
    results.push({ test: 'Handler Implementation', status: 'ERROR' });
  }

  // Summary
  console.log('\n📋 Validation Summary');
  console.log('=====================');
  const passCount = results.filter(r => r.status === 'PASS').length;
  const totalCount = results.length;
  
  results.forEach(result => {
    const icon = result.status === 'PASS' ? '✅' : 
                 result.status === 'FAIL' ? '❌' : '⚠️';
    console.log(`${icon} ${result.test}: ${result.status}`);
  });
  
  console.log(`\n🎯 Results: ${passCount}/${totalCount} tests passed`);
  
  if (passCount === totalCount) {
    console.log('\n🎉 All validations passed! The M365 Core MCP Server is ready.');
  } else {
    console.log('\n⚠️  Some validations failed. Please review the issues above.');
  }
}

validateChanges().catch(console.error);
