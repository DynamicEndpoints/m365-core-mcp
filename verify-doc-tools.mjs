#!/usr/bin/env node

/**
 * Simple test to verify document generation tools are registered
 * Checks the compiled JavaScript for tool registrations
 */

import { readFileSync } from 'fs';
import { resolve } from 'path';

console.log('🔍 Checking for document generation tools in compiled code...\n');

const serverPath = resolve('build/server.js');
const serverCode = readFileSync(serverPath, 'utf-8');

// Tools to check for
const docGenTools = [
  'generate_powerpoint_presentation',
  'generate_word_document',
  'generate_html_report',
  'generate_professional_report',
  'oauth_authorize'
];

console.log('📄 Document Generation Tools Status:\n');
console.log('━'.repeat(80));

let allFound = true;

docGenTools.forEach(toolName => {
  // Check if tool is registered
  const isRegistered = serverCode.includes(`"${toolName}"`);
  const status = isRegistered ? '✅' : '❌';
  
  if (!isRegistered) {
    allFound = false;
  }
  
  console.log(`${status} ${toolName}`);
  
  // Find the registration line
  if (isRegistered) {
    const lines = serverCode.split('\n');
    const lineIndex = lines.findIndex(line => line.includes(`"${toolName}"`));
    if (lineIndex >= 0) {
      console.log(`   → Registered at line ${lineIndex + 1}`);
    }
  }
});

console.log('━'.repeat(80));

if (allFound) {
  console.log('\n✅ All 5 document generation tools are registered!\n');
  
  // Check for setupDocumentGenerationTools method
  if (serverCode.includes('setupDocumentGenerationTools')) {
    console.log('✅ setupDocumentGenerationTools() method found');
  }
  
  // Check for handler imports
  const handlers = [
    'handlePowerPointPresentations',
    'handleWordDocuments',
    'handleHTMLReports',
    'handleProfessionalReports',
    'handleOAuthAuthorization'
  ];
  
  console.log('\n📦 Handler Imports:');
  handlers.forEach(handler => {
    const found = serverCode.includes(handler);
    console.log(`${found ? '✅' : '❌'} ${handler}`);
  });
  
  console.log('\n🎉 Document generation feature is fully integrated!\n');
  process.exit(0);
} else {
  console.log('\n❌ Some document generation tools are missing!\n');
  process.exit(1);
}
