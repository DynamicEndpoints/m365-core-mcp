#!/usr/bin/env node

/**
 * Verification script for new prompts and resources
 * Checks that document generation and policy management prompts/resources are registered
 */

import { readFileSync } from 'fs';

console.log('\n=== VERIFYING PROMPTS & RESOURCES ===\n');

// Read compiled prompts file
const promptsContent = readFileSync('./build/prompts.js', 'utf-8');

// Check for new prompts
const newPrompts = [
  'generate_client_report',
  'policy_management_guide'
];

console.log('📝 Checking Prompts:\n');
let promptsFound = 0;
for (const prompt of newPrompts) {
  const found = promptsContent.includes(`name: '${prompt}'`) || 
                promptsContent.includes(`name: "${prompt}"`);
  console.log(`  ${found ? '✓' : '✗'} ${prompt}`);
  if (found) promptsFound++;
}

// Read compiled resources file
const resourcesContent = readFileSync('./build/resources.js', 'utf-8');

// Check for new resources
const newResources = [
  'm365://documents/presentations',
  'm365://documents/word-documents',
  'm365://documents/reports',
  'm365://documents/templates',
  'm365://policies/conditional-access',
  'm365://policies/retention',
  'm365://policies/information-protection',
  'm365://policies/defender',
  'm365://policies/teams',
  'm365://policies/exchange',
  'm365://policies/sharepoint-governance',
  'm365://policies/security-alerts',
  'm365://policies/overview'
];

console.log('\n📚 Checking Resources:\n');
let resourcesFound = 0;
for (const resource of newResources) {
  const found = resourcesContent.includes(resource);
  console.log(`  ${found ? '✓' : '✗'} ${resource}`);
  if (found) resourcesFound++;
}

// Summary
console.log('\n=== SUMMARY ===\n');
console.log(`Prompts: ${promptsFound}/${newPrompts.length} found`);
console.log(`Resources: ${resourcesFound}/${newResources.length} found`);

if (promptsFound === newPrompts.length && resourcesFound === newResources.length) {
  console.log('\n✅ All prompts and resources registered successfully!\n');
  process.exit(0);
} else {
  console.log('\n❌ Some prompts or resources are missing!\n');
  process.exit(1);
}
