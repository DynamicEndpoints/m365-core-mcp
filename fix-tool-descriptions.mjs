#!/usr/bin/env node

/**
 * Script to add proper descriptions to all MCP tools
 */

import fs from 'fs';

console.log('🔧 Adding proper descriptions to all MCP tools...');

const serverPath = './src/server.ts';
let content = fs.readFileSync(serverPath, 'utf8');

// Define tool descriptions
const toolDescriptions = {
  'manage_intune_macos_policies': 'Create, update, delete, and manage Intune macOS configuration policies including device restrictions and compliance',
  'manage_intune_macos_apps': 'Deploy, update, remove, and manage macOS applications through Microsoft Intune including assignment and monitoring',
  'manage_intune_macos_compliance': 'Configure and manage macOS device compliance policies in Intune including requirements and actions',
  'manage_intune_windows_devices': 'Manage Intune Windows devices including enrollment, compliance, configuration, and remote actions',
  'manage_intune_windows_policies': 'Create, update, delete, and manage Intune Windows configuration policies including device restrictions and security',
  'manage_intune_windows_apps': 'Deploy, update, remove, and manage Windows applications through Microsoft Intune including assignment and monitoring',
  'manage_intune_windows_compliance': 'Configure and manage Windows device compliance policies in Intune including requirements and actions',
  'manage_compliance_frameworks': 'Assess and manage compliance against various frameworks (SOC2, ISO27001, NIST, GDPR, HIPAA)',
  'manage_compliance_assessments': 'Create, run, and manage compliance assessments with automated scoring and gap analysis',
  'manage_compliance_monitoring': 'Monitor compliance status in real-time with alerts, reporting, and automated remediation workflows',
  'manage_evidence_collection': 'Collect, organize, and manage compliance evidence including automated evidence gathering and validation',
  'manage_gap_analysis': 'Perform gap analysis against compliance frameworks with prioritized remediation recommendations',
  'manage_audit_reports': 'Generate comprehensive audit reports with evidence mapping, findings, and executive summaries',
  'manage_cis_compliance': 'Assess and manage CIS (Center for Internet Security) compliance benchmarks and controls'
};

// Replace each tool registration with proper description
Object.entries(toolDescriptions).forEach(([toolName, description]) => {
  // Find the pattern: this.server.tool(\n      "toolName",\n      schemaName.shape,
  const pattern = new RegExp(`(this\\.server\\.tool\\(\\s*"${toolName}",\\s*)([^\\s].*?\\.shape,)`, 'gs');
  
  content = content.replace(pattern, `$1"${description}",\n      $2`);
});

// Write the updated content
fs.writeFileSync(serverPath, content);

console.log('✅ Tool descriptions added successfully!');

// Verify by counting the number of descriptions
const descriptionCount = (content.match(/this\.server\.tool\([^,]+,\s*"[^"]*",\s*[^,]+\.shape,/g) || []).length;
console.log(`📊 Total tools with descriptions: ${descriptionCount}`);
