import { z } from 'zod';

// Define proper Zod schemas for MCP tool discovery and Smithery registry

// SharePoint Site Management
export const sharePointSiteSchema = z.object({
  action: z.enum(['get', 'create', 'update', 'delete', 'add_users', 'remove_users']).describe('Action to perform on SharePoint site'),
  siteId: z.string().optional().describe('SharePoint site ID for existing site operations'),
  url: z.string().optional().describe('URL for the SharePoint site'),
  title: z.string().optional().describe('Title for the SharePoint site'),
  description: z.string().optional().describe('Description of the SharePoint site'),
  template: z.string().optional().describe('Web template ID for site creation (e.g., STS#3 for Modern Team Site)'),
  owners: z.array(z.string()).optional().describe('List of owner email addresses'),
  members: z.array(z.string()).optional().describe('List of member email addresses'),
  settings: z.object({
    isPublic: z.boolean().optional().describe('Whether the site is public'),
    allowSharing: z.boolean().optional().describe('Allow external sharing'),
    storageQuota: z.number().optional().describe('Storage quota in MB'),
  }).optional().describe('Site configuration settings'),
});

// SharePoint List Management
export const sharePointListSchema = z.object({
  action: z.enum(['get', 'create', 'update', 'delete', 'add_items', 'get_items']).describe('Action to perform on SharePoint list'),
  siteId: z.string().describe('SharePoint site ID containing the list'),
  listId: z.string().optional().describe('SharePoint list ID for existing list operations'),
  title: z.string().optional().describe('Title for the SharePoint list'),
  description: z.string().optional().describe('Description of the SharePoint list'),
  template: z.string().optional().describe('Template to use for list creation'),
  columns: z.array(z.object({
    name: z.string().describe('Column name'),
    type: z.string().describe('Column type (Text, Number, DateTime, etc.)'),
    required: z.boolean().optional().describe('Whether the column is required'),
    defaultValue: z.any().optional().describe('Default value for the column'),
  })).optional().describe('List column definitions'),
  items: z.array(z.record(z.any())).optional().describe('Items to add to the list'),
});

// Distribution List Management
export const distributionListSchema = z.object({
  action: z.enum(['get', 'create', 'update', 'delete', 'add_members', 'remove_members']).describe('Action to perform on distribution list'),
  listId: z.string().optional().describe('Distribution list ID for existing list operations'),
  displayName: z.string().optional().describe('Display name for the distribution list'),
  emailAddress: z.string().optional().describe('Email address for the distribution list'),
  members: z.array(z.string()).optional().describe('List of member email addresses'),
  settings: z.object({
    hideFromGAL: z.boolean().optional().describe('Hide from Global Address List'),
    requireSenderAuthentication: z.boolean().optional().describe('Require sender authentication'),
    moderatedBy: z.array(z.string()).optional().describe('List of moderator email addresses'),
  }).optional().describe('Distribution list settings'),
});

// Security Group Management
export const securityGroupSchema = z.object({
  action: z.enum(['get', 'create', 'update', 'delete', 'add_members', 'remove_members']).describe('Action to perform on security group'),
  groupId: z.string().optional().describe('Security group ID for existing group operations'),
  displayName: z.string().optional().describe('Display name for the security group'),
  description: z.string().optional().describe('Description of the security group'),
  members: z.array(z.string()).optional().describe('List of member email addresses'),
  settings: z.object({
    securityEnabled: z.boolean().optional().describe('Whether security is enabled'),
    mailEnabled: z.boolean().optional().describe('Whether mail is enabled'),
  }).optional().describe('Security group settings'),
});

// Microsoft 365 Group Management
export const m365GroupSchema = z.object({
  action: z.enum(['get', 'create', 'update', 'delete', 'add_members', 'remove_members']).describe('Action to perform on M365 group'),
  groupId: z.string().optional().describe('M365 group ID for existing group operations'),
  displayName: z.string().optional().describe('Display name for the M365 group'),
  description: z.string().optional().describe('Description of the M365 group'),
  owners: z.array(z.string()).optional().describe('List of owner email addresses'),
  members: z.array(z.string()).optional().describe('List of member email addresses'),
  settings: z.object({
    visibility: z.enum(['Private', 'Public']).optional().describe('Group visibility setting'),
    allowExternalSenders: z.boolean().optional().describe('Allow external senders'),
    autoSubscribeNewMembers: z.boolean().optional().describe('Auto-subscribe new members'),
  }).optional().describe('M365 group settings'),
});

// Exchange Online Settings Management
export const exchangeSettingsSchema = z.object({
  action: z.enum(['get', 'update']).describe('Action to perform on Exchange settings'),
  settingType: z.enum(['mailbox', 'transport', 'organization', 'retention']).describe('Type of Exchange settings to manage'),
  target: z.string().optional().describe('User/Group ID for mailbox settings'),
  settings: z.object({
    automateProcessing: z.object({
      autoForwardEnabled: z.boolean().optional().describe('Enable auto-forwarding'),
      autoReplyEnabled: z.boolean().optional().describe('Enable auto-reply'),
    }).optional().describe('Automated processing settings'),
    rules: z.array(z.object({
      name: z.string().describe('Rule name'),
      conditions: z.record(z.any()).describe('Rule conditions'),
      actions: z.record(z.any()).describe('Rule actions'),
    })).optional().describe('Transport rules'),
    sharingPolicy: z.object({
      enabled: z.boolean().optional().describe('Enable sharing policy'),
      domains: z.array(z.string()).optional().describe('Allowed domains'),
    }).optional().describe('Sharing policy settings'),
    retentionTags: z.array(z.object({
      name: z.string().describe('Retention tag name'),
      type: z.string().describe('Retention tag type'),
      retentionDays: z.number().describe('Retention period in days'),
    })).optional().describe('Retention tag definitions'),
  }).optional().describe('Exchange configuration settings'),
});

// User Management Settings
export const userManagementSchema = z.object({
  action: z.enum(['get', 'update']).describe('Action to perform on user settings'),
  userId: z.string().describe('User ID or UPN'),
  settings: z.record(z.unknown()).optional().describe('User settings to update'),
});

// User Offboarding Process
export const offboardingSchema = z.object({
  action: z.enum(['start', 'check', 'complete']).describe('Offboarding process action'),
  userId: z.string().describe('User ID or UPN to offboard'),
  options: z.object({
    revokeAccess: z.boolean().optional().describe('Revoke all access immediately'),
    retainMailbox: z.boolean().optional().describe('Retain user mailbox'),
    convertToShared: z.boolean().optional().describe('Convert mailbox to shared'),
    backupData: z.boolean().optional().describe('Backup user data'),
  }).optional().describe('Offboarding options'),
});

// Azure AD Role Management
export const azureAdRoleSchema = z.object({
  action: z.enum(['list_roles', 'list_role_assignments', 'assign_role', 'remove_role_assignment']).describe('Azure AD role management action'),
  roleId: z.string().optional().describe('ID of the directory role'),
  principalId: z.string().optional().describe('ID of the principal (user, group, SP)'),
  assignmentId: z.string().optional().describe('ID of the role assignment to remove'),
  filter: z.string().optional().describe('OData filter string'),
});

// Azure AD Application Management
export const azureAdAppSchema = z.object({
  action: z.enum(['list_apps', 'get_app', 'update_app', 'add_owner', 'remove_owner']).describe('Azure AD application management action'),
  appId: z.string().optional().describe('Object ID of the application'),
  ownerId: z.string().optional().describe('Object ID of the user to add/remove as owner'),
  appDetails: z.object({
    displayName: z.string().optional().describe('Application display name'),
    signInAudience: z.string().optional().describe('Sign-in audience setting'),
  }).optional().describe('Application details for updates'),
  filter: z.string().optional().describe('OData filter string'),
});

// Azure AD Device Management
export const azureAdDeviceSchema = z.object({
  action: z.enum(['list_devices', 'get_device', 'enable_device', 'disable_device', 'delete_device']).describe('Azure AD device management action'),
  deviceId: z.string().optional().describe('Object ID of the device'),
  filter: z.string().optional().describe('OData filter string'),
});

// Service Principal Management
export const azureAdSpSchema = z.object({
  action: z.enum(['list_sps', 'get_sp', 'add_owner', 'remove_owner']).describe('Service principal management action'),
  spId: z.string().optional().describe('Object ID of the Service Principal'),
  ownerId: z.string().optional().describe('Object ID of the user to add/remove as owner'),
  filter: z.string().optional().describe('OData filter string'),
});

// Dynamic Microsoft API Calling
export const callMicrosoftApiSchema = z.object({
  apiType: z.enum(['graph', 'azure']).describe('API type: Microsoft Graph or Azure Resource Management'),
  path: z.string().describe('API URL path (e.g., \'/users\')'),
  method: z.enum(['get', 'post', 'put', 'patch', 'delete']).describe('HTTP method'),
  apiVersion: z.string().optional().describe('Azure API version (required for Azure APIs)'),
  subscriptionId: z.string().optional().describe('Azure Subscription ID (for Azure APIs)'),
  queryParams: z.record(z.string()).optional().describe('Query parameters'),
  body: z.record(z.any()).optional().describe('Request body (for POST, PUT, PATCH)'),
  consistencyLevel: z.string().optional().describe('Consistency level header (eventual, strong)'),
});

// Audit Log Search
export const auditLogSchema = z.object({
  filter: z.string().optional().describe('OData filter string (e.g., \'activityDateTime ge 2024-01-01T00:00:00Z\')'),
  top: z.number().optional().describe('Maximum number of records to return'),
});

// Security Alerts Management
export const alertSchema = z.object({
  action: z.enum(['list_alerts', 'get_alert']).describe('Alert management action'),
  alertId: z.string().optional().describe('ID of the alert (required for get_alert)'),
  filter: z.string().optional().describe('OData filter string (e.g., \'status eq \\\'new\\\'\')'),
  top: z.number().optional().describe('Maximum number of alerts to return'),
});

// DLP Policy Management
export const dlpPolicySchema = z.object({
  action: z.enum(['list', 'get', 'create', 'update', 'delete']).describe('DLP policy management action'),
  policyId: z.string().optional().describe('DLP policy ID'),
  name: z.string().optional().describe('Policy name'),
  description: z.string().optional().describe('Policy description'),
  mode: z.enum(['Test', 'TestWithoutNotifications', 'Enforce']).optional().describe('Policy mode'),
  locations: z.array(z.string()).optional().describe('Policy locations (Exchange, SharePoint, OneDrive, etc.)'),
  rules: z.array(z.record(z.any())).optional().describe('Policy rules configuration'),
});

// DLP Incident Management
export const dlpIncidentSchema = z.object({
  action: z.enum(['list', 'get', 'update']).describe('DLP incident management action'),
  incidentId: z.string().optional().describe('DLP incident ID'),
  status: z.enum(['Active', 'Resolved', 'Dismissed']).optional().describe('Incident status'),
  filter: z.string().optional().describe('Filter criteria'),
});

// Sensitivity Label Management
export const sensitivityLabelSchema = z.object({
  action: z.enum(['list', 'get', 'create', 'update', 'delete', 'apply']).describe('Sensitivity label action'),
  labelId: z.string().optional().describe('Sensitivity label ID'),
  name: z.string().optional().describe('Label name'),
  description: z.string().optional().describe('Label description'),
  targetId: z.string().optional().describe('Target resource ID for label application'),
  settings: z.record(z.any()).optional().describe('Label settings and policies'),
});

// Intune macOS Management
export const intuneMacOSDeviceSchema = z.object({
  action: z.enum(['list', 'get', 'sync', 'wipe', 'retire']).describe('Intune macOS device management action'),
  deviceId: z.string().optional().describe('Device ID for device-specific operations'),
  filter: z.string().optional().describe('OData filter for device listing'),
});

export const intuneMacOSPolicySchema = z.object({
  action: z.enum(['list', 'get', 'create', 'update', 'delete', 'assign']).describe('Intune macOS policy management action'),
  policyId: z.string().optional().describe('Policy ID for policy-specific operations'),
  policyType: z.enum(['configuration', 'compliance', 'security']).optional().describe('Type of macOS policy'),
  name: z.string().optional().describe('Policy name'),
  settings: z.record(z.any()).optional().describe('Policy configuration settings'),
  assignments: z.array(z.string()).optional().describe('Group IDs for policy assignment'),
});

export const intuneMacOSAppSchema = z.object({
  action: z.enum(['list', 'get', 'deploy', 'update', 'remove']).describe('Intune macOS app management action'),
  appId: z.string().optional().describe('App ID for app-specific operations'),
  bundleId: z.string().optional().describe('macOS app bundle identifier'),
  name: z.string().optional().describe('Application name'),
  version: z.string().optional().describe('Application version'),
  assignmentGroups: z.array(z.string()).optional().describe('Target groups for app deployment'),
});

export const intuneMacOSComplianceSchema = z.object({
  action: z.enum(['assess', 'report', 'remediate']).describe('Intune macOS compliance action'),
  deviceId: z.string().optional().describe('Device ID for compliance assessment'),
  complianceType: z.enum(['security', 'configuration', 'update']).optional().describe('Type of compliance check'),
  policies: z.array(z.string()).optional().describe('Specific policy IDs to assess'),
});

// Compliance Framework Management
export const complianceFrameworkSchema = z.object({
  action: z.enum(['list', 'get', 'assess', 'report']).describe('Compliance framework management action'),
  framework: z.enum(['SOC2', 'ISO27001', 'NIST', 'CIS', 'PCI-DSS', 'HIPAA']).optional().describe('Compliance framework type'),
  scope: z.string().optional().describe('Assessment scope (organization, specific systems)'),
  includeEvidence: z.boolean().optional().describe('Include compliance evidence in reports'),
});

export const complianceAssessmentSchema = z.object({
  action: z.enum(['start', 'status', 'results', 'export']).describe('Compliance assessment action'),
  assessmentId: z.string().optional().describe('Assessment ID for tracking'),
  framework: z.enum(['SOC2', 'ISO27001', 'NIST', 'CIS', 'PCI-DSS', 'HIPAA']).describe('Framework to assess against'),
  scope: z.array(z.string()).optional().describe('Systems or services to include in assessment'),
  automated: z.boolean().optional().describe('Whether to run automated checks'),
});

export const complianceMonitoringSchema = z.object({
  action: z.enum(['setup', 'status', 'alerts', 'report']).describe('Compliance monitoring action'),
  framework: z.enum(['SOC2', 'ISO27001', 'NIST', 'CIS', 'PCI-DSS', 'HIPAA']).describe('Framework to monitor'),
  frequency: z.enum(['daily', 'weekly', 'monthly']).optional().describe('Monitoring frequency'),
  alertThreshold: z.enum(['low', 'medium', 'high', 'critical']).optional().describe('Alert threshold level'),
});

export const evidenceCollectionSchema = z.object({
  action: z.enum(['collect', 'list', 'export', 'validate']).describe('Evidence collection action'),
  evidenceType: z.enum(['configuration', 'logs', 'policies', 'certificates', 'reports']).optional().describe('Type of evidence to collect'),
  timeRange: z.object({
    start: z.string().describe('Start date (ISO format)'),
    end: z.string().describe('End date (ISO format)'),
  }).optional().describe('Time range for evidence collection'),
  systems: z.array(z.string()).optional().describe('Specific systems to collect evidence from'),
});

export const gapAnalysisSchema = z.object({
  action: z.enum(['analyze', 'report', 'recommendations']).describe('Gap analysis action'),
  framework: z.enum(['SOC2', 'ISO27001', 'NIST', 'CIS', 'PCI-DSS', 'HIPAA']).describe('Framework for gap analysis'),
  currentState: z.record(z.any()).optional().describe('Current compliance state data'),
  targetState: z.enum(['basic', 'intermediate', 'advanced']).optional().describe('Target compliance level'),
});

export const auditReportSchema = z.object({
  action: z.enum(['generate', 'schedule', 'list', 'export']).describe('Audit report action'),
  reportType: z.enum(['compliance', 'security', 'activity', 'configuration']).describe('Type of audit report'),
  timeRange: z.object({
    start: z.string().describe('Start date (ISO format)'),
    end: z.string().describe('End date (ISO format)'),
  }).optional().describe('Report time range'),
  format: z.enum(['pdf', 'xlsx', 'csv', 'json']).optional().describe('Report output format'),
  includeEvidence: z.boolean().optional().describe('Include supporting evidence'),
});

export const cisComplianceSchema = z.object({
  action: z.enum(['assess', 'report', 'remediate', 'monitor']).describe('CIS compliance action'),
  benchmark: z.enum(['Windows', 'macOS', 'Linux', 'Azure', 'M365']).describe('CIS benchmark to assess'),
  level: z.enum(['Level1', 'Level2']).optional().describe('CIS benchmark level'),
  systems: z.array(z.string()).optional().describe('Target systems for assessment'),
  autoRemediate: z.boolean().optional().describe('Automatically remediate findings'),
});

// Core M365 Tools Collection
export const m365CoreTools = {
  distributionLists: distributionListSchema,
  securityGroups: securityGroupSchema,
  m365Groups: m365GroupSchema,
  exchangeSettings: exchangeSettingsSchema,
  userManagement: userManagementSchema,
  offboarding: offboardingSchema,
  sharePointSites: sharePointSiteSchema,
  sharePointLists: sharePointListSchema,
  azureAdRoles: azureAdRoleSchema,
  azureAdApps: azureAdAppSchema,
  azureAdDevices: azureAdDeviceSchema,
  servicePrincipals: azureAdSpSchema,
  apiCalls: callMicrosoftApiSchema,
  auditLogs: auditLogSchema,
  alerts: alertSchema,
  dlpPolicies: dlpPolicySchema,
  dlpIncidents: dlpIncidentSchema,
  sensitivityLabels: sensitivityLabelSchema,
  intuneMacOSDevices: intuneMacOSDeviceSchema,
  intuneMacOSPolicies: intuneMacOSPolicySchema,
  intuneMacOSApps: intuneMacOSAppSchema,
  intuneMacOSCompliance: intuneMacOSComplianceSchema,
  complianceFrameworks: complianceFrameworkSchema,
  complianceAssessments: complianceAssessmentSchema,
  complianceMonitoring: complianceMonitoringSchema,
  evidenceCollection: evidenceCollectionSchema,
  gapAnalysis: gapAnalysisSchema,
  auditReports: auditReportSchema,
  cisCompliance: cisComplianceSchema,
} as const;

export { z };
