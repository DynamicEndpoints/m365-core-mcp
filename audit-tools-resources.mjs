#!/usr/bin/env node

/**
 * Audit script to verify tools and resources match the README documentation
 */

import fs from 'fs';
import path from 'path';

console.log('🔍 Auditing M365 Core MCP Server Tools & Resources');
console.log('================================================\n');

// Extract tools from server.ts
function extractToolsFromServer() {
  const serverContent = fs.readFileSync('./src/server.ts', 'utf8');
  const toolMatches = serverContent.match(/this\.server\.tool\(\s*"([^"]+)"/g);
  
  if (!toolMatches) return [];
  
  return toolMatches.map(match => {
    const toolName = match.match(/"([^"]+)"/)[1];
    return toolName;
  }).sort();
}

// Extract resources from server.ts
function extractResourcesFromServer() {
  const serverContent = fs.readFileSync('./src/server.ts', 'utf8');
  const resourceMatches = serverContent.match(/this\.server\.resource\(\s*'([^']+)'/g);
  
  if (!resourceMatches) return [];
  
  return resourceMatches.map(match => {
    const resourceName = match.match(/'([^']+)'/)[1];
    return resourceName;
  }).sort();
}

// Extract tools mentioned in README
function extractToolsFromReadme() {
  const readmeContent = fs.readFileSync('./README.md', 'utf8');
  
  // Look for tool examples like await callTool('tool_name', ...)
  const toolExamples = readmeContent.match(/callTool\('([^']+)'/g);
  const exampleTools = toolExamples ? toolExamples.map(match => match.match(/'([^']+)'/)[1]) : [];
  
  // Look for bullet points mentioning tools like - `tool_name`: description
  const toolBullets = readmeContent.match(/- `([^`]+)`:/g);
  const bulletTools = toolBullets ? toolBullets.map(match => match.match(/`([^`]+)`/)[1]) : [];
  
  return [...new Set([...exampleTools, ...bulletTools])].sort();
}

// Extract resources mentioned in README
function extractResourcesFromReadme() {
  const readmeContent = fs.readFileSync('./README.md', 'utf8');
  
  // Look for resource examples like m365://something
  const resourceMatches = readmeContent.match(/- `(m365:\/\/[^`]+)`/g);
  
  if (!resourceMatches) return [];
  
  return resourceMatches.map(match => {
    const resourceUri = match.match(/`([^`]+)`/)[1];
    return resourceUri;
  }).sort();
}

// Main audit function
async function auditToolsAndResources() {
  console.log('1️⃣ Extracting tools from server implementation...');
  const serverTools = extractToolsFromServer();
  console.log(`   Found ${serverTools.length} tools in server.ts`);
  
  console.log('\n2️⃣ Extracting tools from README documentation...');
  const readmeTools = extractToolsFromReadme();
  console.log(`   Found ${readmeTools.length} tools in README.md`);
  
  console.log('\n3️⃣ Extracting resources from server implementation...');
  const serverResources = extractResourcesFromServer();
  console.log(`   Found ${serverResources.length} resources in server.ts`);
  
  console.log('\n4️⃣ Extracting resources from README documentation...');
  const readmeResources = extractResourcesFromReadme();
  console.log(`   Found ${readmeResources.length} resources in README.md`);
  
  console.log('\n📋 TOOLS ANALYSIS');
  console.log('==================');
  console.log('\n🔧 Implemented Tools:');
  serverTools.forEach(tool => console.log(`   ✅ ${tool}`));
  
  console.log('\n📚 Documented Tools:');
  readmeTools.forEach(tool => console.log(`   📖 ${tool}`));
  
  // Find tools in server but not in README
  const undocumentedTools = serverTools.filter(tool => !readmeTools.includes(tool));
  if (undocumentedTools.length > 0) {
    console.log('\n⚠️  Tools implemented but not documented:');
    undocumentedTools.forEach(tool => console.log(`   🔍 ${tool}`));
  }
  
  // Find tools in README but not in server
  const unimplementedTools = readmeTools.filter(tool => !serverTools.includes(tool));
  if (unimplementedTools.length > 0) {
    console.log('\n❌ Tools documented but not implemented:');
    unimplementedTools.forEach(tool => console.log(`   📝 ${tool}`));
  }
  
  console.log('\n📋 RESOURCES ANALYSIS');
  console.log('=====================');
  console.log('\n🔧 Implemented Resources:');
  serverResources.forEach(resource => console.log(`   ✅ ${resource}`));
  
  console.log('\n📚 Documented Resources:');
  readmeResources.forEach(resource => console.log(`   📖 ${resource}`));
  
  // Find resources in server but not in README
  const undocumentedResources = serverResources.filter(resource => !readmeResources.some(r => r.includes(resource)));
  if (undocumentedResources.length > 0) {
    console.log('\n⚠️  Resources implemented but not documented:');
    undocumentedResources.forEach(resource => console.log(`   🔍 ${resource}`));
  }
  
  // Find resources in README but not in server
  const unimplementedResources = readmeResources.filter(resource => !serverResources.some(r => resource.includes(r)));
  if (unimplementedResources.length > 0) {
    console.log('\n❌ Resources documented but not implemented:');
    unimplementedResources.forEach(resource => console.log(`   📝 ${resource}`));
  }
  
  console.log('\n📊 SUMMARY');
  console.log('==========');
  
  const toolsMatch = serverTools.length === readmeTools.length && undocumentedTools.length === 0 && unimplementedTools.length === 0;
  const resourcesMatch = undocumentedResources.length === 0 && unimplementedResources.length === 0;
  
  console.log(`Tools: ${toolsMatch ? '✅ MATCH' : '❌ MISMATCH'}`);
  console.log(`Resources: ${resourcesMatch ? '✅ MATCH' : '❌ MISMATCH'}`);
  
  if (toolsMatch && resourcesMatch) {
    console.log('\n🎉 All tools and resources are properly documented!');
  } else {
    console.log('\n🔧 Some tools or resources need documentation updates.');
  }
  
  return {
    toolsMatch,
    resourcesMatch,
    serverTools,
    readmeTools,
    serverResources,
    readmeResources,
    undocumentedTools,
    unimplementedTools,
    undocumentedResources,
    unimplementedResources
  };
}

// Run the audit
auditToolsAndResources().catch(console.error);
