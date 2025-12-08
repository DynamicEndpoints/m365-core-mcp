#!/usr/bin/env node
/**
 * Script to update all tool registrations with descriptions and annotations
 */

import { readFileSync, writeFileSync } from 'fs';
import { toolMetadata } from './build/tool-metadata.js';

const serverPath = './src/server.ts';
let serverContent = readFileSync(serverPath, 'utf-8');

// Tool name to schema name mapping
const toolToSchema = {
  manage_distribution_lists: 'distributionListSchema',
  manage_security_groups: 'securityGroupSchema',
  manage_m365_groups: 'm365GroupSchema',
  manage_sharepoint_sites: 'sharePointSiteSchema',
  manage_sharepoint_lists: 'sharePointListSchema',
  manage_user_settings: 'userManagementSchema',
  manage_offboarding: 'offboardingSchema',
  manage_exchange_settings: 'exchangeSettingsSchema',
  manage_azure_ad_roles: 'azureAdRoleSchema',
  manage_azure_ad_apps: 'azureAdAppSchema',
  manage_azure_ad_devices: 'azureAdDeviceSchema',
  manage_service_principals: 'azureAdSpSchema',
  call_microsoft_api: 'callMicrosoftApiSchema',
  search_audit_log: 'auditLogSchema',
  manage_alerts: 'alertSchema',
  manage_dlp_policies: 'dlpPolicySchema',
  manage_dlp_incidents: 'dlpIncidentSchema',
  manage_sensitivity_labels: 'sensitivityLabelSchema',
  manage_intune_macos_devices: 'intuneMacOSDeviceSchema',
  manage_intune_macos_policies: 'intuneMacOSPolicySchema',
  manage_intune_macos_apps: 'intuneMacOSAppSchema',
  manage_intune_macos_compliance: 'intuneMacOSComplianceSchema',
  manage_intune_windows_devices: 'intuneWindowsDeviceSchema',
  manage_intune_windows_policies: 'intuneWindowsPolicySchema',
  manage_intune_windows_apps: 'intuneWindowsAppSchema',
  manage_intune_windows_compliance: 'intuneWindowsComplianceSchema',
  manage_compliance_frameworks: 'complianceFrameworkSchema',
  manage_compliance_assessments: 'complianceAssessmentSchema',
  manage_compliance_monitoring: 'complianceMonitoringSchema',
  manage_evidence_collection: 'evidenceCollectionSchema',
  manage_gap_analysis: 'gapAnalysisSchema',
  generate_audit_reports: 'auditReportSchema',
  manage_cis_compliance: 'cisComplianceSchema',
  execute_graph_batch: 'batchRequestSchema',
  execute_delta_query: 'deltaQuerySchema',
  manage_graph_subscriptions: 'webhookSubscriptionSchema',
  execute_graph_search: 'searchQuerySchema',
  manage_retention_policies: 'retentionPolicyArgsSchema',
  manage_conditional_access_policies: 'conditionalAccessPolicyArgsSchema',
  manage_information_protection_policies: 'informationProtectionPolicyArgsSchema',
  manage_defender_policies: 'defenderPolicyArgsSchema',
  manage_teams_policies: 'teamsPolicyArgsSchema',
  manage_exchange_policies: 'exchangePolicyArgsSchema',
  manage_sharepoint_governance_policies: 'sharePointGovernancePolicyArgsSchema',
  manage_security_alert_policies: 'securityAlertPolicyArgsSchema',
  generate_powerpoint_presentation: 'powerPointPresentationArgsSchema',
  generate_word_document: 'wordDocumentArgsSchema',
  generate_html_report: 'htmlReportArgsSchema',
  generate_professional_report: 'professionalReportArgsSchema',
  oauth_authorize: 'oauthAuthorizationArgsSchema'
};

console.log('Updating tool registrations...\n');

let updatedCount = 0;
for (const [toolName, schemaName] of Object.entries(toolToSchema)) {
  // Skip if already updated
  if (toolName === 'manage_distribution_lists') {
    console.log(`✓ ${toolName} (already updated)`);
    continue;
  }
  
  const metadata = toolMetadata[toolName];
  if (!metadata) {
    console.log(`⚠ ${toolName} - no metadata found`);
    continue;
  }

  // Pattern to match: this.server.tool("toolName", schemaName.shape,
  const oldPattern = new RegExp(
    `this\\.server\\.tool\\(\\s*"${toolName}",\\s*${schemaName}\\.shape,`,
    'g'
  );

  // New pattern with description and annotations
  const newPattern = `this.server.tool(\n      "${toolName}",\n      "${metadata.description}",\n      ${schemaName}.shape,\n      ${JSON.stringify(metadata.annotations || {})},`;

  const beforeCount = (serverContent.match(oldPattern) || []).length;
  serverContent = serverContent.replace(oldPattern, newPattern);
  const afterCount = (serverContent.match(oldPattern) || []).length;

  if (beforeCount > afterCount) {
    updatedCount++;
    console.log(`✓ ${toolName}`);
  } else {
    console.log(`✗ ${toolName} - pattern not found`);
  }
}

writeFileSync(serverPath, serverContent, 'utf-8');

console.log(`\n✅ Updated ${updatedCount} tool registrations`);
console.log(`\nRun 'npm run build' to verify the changes.`);
