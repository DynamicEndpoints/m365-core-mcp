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
  action: z.enum(['list', 'get', 'create', 'update', 'delete', 'test']).describe('DLP policy management action'),
  policyId: z.string().optional().describe('DLP policy ID'),
  name: z.string().optional().describe('Policy name'),
  description: z.string().optional().describe('Policy description'),
  locations: z.array(z.enum(['Exchange', 'SharePoint', 'OneDrive', 'Teams', 'Endpoint'])).optional().describe('Policy locations'),  rules: z.array(z.object({
    name: z.string().describe('Rule name'),
    conditions: z.array(z.object({
      type: z.enum(['ContentContains', 'SensitiveInfoType', 'DocumentProperty', 'MessageProperty']).describe('Condition type'),
      value: z.string().describe('Condition value'),
      operator: z.enum(['Equals', 'Contains', 'StartsWith', 'EndsWith', 'RegexMatch']).optional().describe('Condition operator'),
      caseSensitive: z.boolean().optional().describe('Case sensitive matching'),
    })).describe('Rule conditions'),
    actions: z.array(z.object({
      type: z.enum(['Block', 'BlockWithOverride', 'Notify', 'Audit', 'Quarantine']).describe('Action type'),
      settings: z.object({
        notificationMessage: z.string().optional().describe('Notification message'),
        blockMessage: z.string().optional().describe('Block message'),
        allowOverride: z.boolean().optional().describe('Allow override'),
        overrideJustificationRequired: z.boolean().optional().describe('Override justification required'),
      }).optional().describe('Action settings'),
    })).describe('Rule actions'),
    enabled: z.boolean().optional().describe('Whether rule is enabled'),
    priority: z.number().optional().describe('Rule priority'),
  })).optional().describe('Policy rules configuration'),
  settings: z.object({
    mode: z.enum(['Test', 'TestWithNotifications', 'Enforce']).optional().describe('Policy mode'),
    priority: z.number().optional().describe('Policy priority'),
    enabled: z.boolean().optional().describe('Whether policy is enabled'),
  }).optional().describe('Policy settings'),
});

// DLP Incident Management
export const dlpIncidentSchema = z.object({
  action: z.enum(['list', 'get', 'resolve', 'escalate']).describe('DLP incident management action'),
  incidentId: z.string().optional().describe('DLP incident ID'),
  dateRange: z.object({
    startDate: z.string().describe('Start date'),
    endDate: z.string().describe('End date'),
  }).optional().describe('Date range filter'),
  severity: z.enum(['Low', 'Medium', 'High', 'Critical']).optional().describe('Incident severity'),
  status: z.enum(['Active', 'Resolved', 'InProgress', 'Dismissed']).optional().describe('Incident status'),
  policyId: z.string().optional().describe('Associated policy ID'),
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
  action: z.enum(['list', 'get', 'enroll', 'retire', 'wipe', 'restart', 'sync', 'remote_lock', 'collect_logs']).describe('Intune macOS device management action'),
  deviceId: z.string().optional().describe('Device ID for device-specific operations'),
  filter: z.string().optional().describe('OData filter for device listing'),
  enrollmentType: z.enum(['UserEnrollment', 'DeviceEnrollment', 'AutomaticDeviceEnrollment']).optional().describe('Enrollment type'),
  assignmentTarget: z.object({
    groupIds: z.array(z.string()).optional().describe('Target group IDs'),
    userIds: z.array(z.string()).optional().describe('Target user IDs'),
    deviceIds: z.array(z.string()).optional().describe('Target device IDs'),
  }).optional().describe('Assignment target'),
});

export const intuneMacOSPolicySchema = z.object({
  action: z.enum(['list', 'get', 'create', 'update', 'delete', 'assign', 'deploy']).describe('Intune macOS policy management action'),
  policyId: z.string().optional().describe('Policy ID for policy-specific operations'),
  policyType: z.enum(['Configuration', 'Compliance', 'Security', 'Update', 'AppProtection']).describe('Type of macOS policy'),
  name: z.string().optional().describe('Policy name'),
  description: z.string().optional().describe('Policy description'),
  settings: z.record(z.any()).optional().describe('Policy configuration settings'),
  assignments: z.array(z.object({
    target: z.object({
      deviceAndAppManagementAssignmentFilterId: z.string().optional().describe('Filter ID'),
      deviceAndAppManagementAssignmentFilterType: z.enum(['none', 'include', 'exclude']).optional().describe('Filter type'),
      groupId: z.string().optional().describe('Group ID'),
      collectionId: z.string().optional().describe('Collection ID'),
    }).describe('Assignment target'),
    intent: z.enum(['apply', 'remove']).optional().describe('Assignment intent'),
    settings: z.record(z.any()).optional().describe('Assignment settings'),
  })).optional().describe('Policy assignments'),
  deploymentSettings: z.object({
    installBehavior: z.enum(['doNotInstall', 'installAsManaged', 'installAsUnmanaged']).optional().describe('Install behavior'),
    uninstallOnDeviceRemoval: z.boolean().optional().describe('Uninstall on device removal'),
    installAsManaged: z.boolean().optional().describe('Install as managed'),
  }).optional().describe('Deployment settings'),
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
  action: z.enum(['get_status', 'get_details', 'update_policy', 'force_evaluation']).describe('Intune macOS compliance action'),
  deviceId: z.string().optional().describe('Device ID for compliance assessment'),
  complianceType: z.enum(['security', 'configuration', 'update']).optional().describe('Type of compliance check'),
  policies: z.array(z.string()).optional().describe('Specific policy IDs to assess'),
});

// Intune Windows Management
export const intuneWindowsDeviceSchema = z.object({
  action: z.enum(['list', 'get', 'enroll', 'retire', 'wipe', 'restart', 'sync', 'remote_lock', 'collect_logs', 'bitlocker_recovery', 'autopilot_reset']).describe('Intune Windows device management action'),
  deviceId: z.string().optional().describe('Device ID for device-specific operations'),
  filter: z.string().optional().describe('OData filter for device listing'),
  enrollmentType: z.enum(['AzureADJoin', 'HybridAzureADJoin', 'AutoPilot', 'BulkEnrollment']).optional().describe('Windows enrollment type'),
  assignmentTarget: z.object({
    groupIds: z.array(z.string()).optional().describe('Target group IDs'),
    userIds: z.array(z.string()).optional().describe('Target user IDs'),
    deviceIds: z.array(z.string()).optional().describe('Target device IDs'),
  }).optional().describe('Assignment target'),
  bitlockerSettings: z.object({
    requireBitlockerEncryption: z.boolean().optional().describe('Require BitLocker encryption'),
    allowBitlockerRecoveryKeyBackup: z.boolean().optional().describe('Allow recovery key backup'),
  }).optional().describe('BitLocker configuration'),
});

export const intuneWindowsPolicySchema = z.object({
  action: z.enum(['list', 'get', 'create', 'update', 'delete', 'assign', 'deploy']).describe('Intune Windows policy management action'),
  policyId: z.string().optional().describe('Policy ID for policy-specific operations'),
  policyType: z.enum(['Configuration', 'Compliance', 'Security', 'Update', 'AppProtection', 'EndpointSecurity']).describe('Type of Windows policy'),
  name: z.string().optional().describe('Policy name'),
  description: z.string().optional().describe('Policy description'),
  settings: z.record(z.any()).optional().describe('Policy configuration settings'),
  assignments: z.array(z.object({
    target: z.object({
      deviceAndAppManagementAssignmentFilterId: z.string().optional().describe('Filter ID'),
      deviceAndAppManagementAssignmentFilterType: z.enum(['none', 'include', 'exclude']).optional().describe('Filter type'),
      groupId: z.string().optional().describe('Group ID'),
      collectionId: z.string().optional().describe('Collection ID'),
    }).describe('Assignment target'),
    intent: z.enum(['apply', 'remove']).optional().describe('Assignment intent'),
    settings: z.record(z.any()).optional().describe('Assignment settings'),
  })).optional().describe('Policy assignments'),
  deploymentSettings: z.object({
    installBehavior: z.enum(['doNotInstall', 'installAsManaged', 'installAsUnmanaged']).optional().describe('Install behavior'),
    uninstallOnDeviceRemoval: z.boolean().optional().describe('Uninstall on device removal'),
    installAsManaged: z.boolean().optional().describe('Install as managed'),
  }).optional().describe('Deployment settings'),
});

export const intuneWindowsAppSchema = z.object({
  action: z.enum(['list', 'get', 'deploy', 'update', 'remove', 'sync_status']).describe('Intune Windows app management action'),
  appId: z.string().optional().describe('App ID for app-specific operations'),
  appType: z.enum(['win32LobApp', 'microsoftStoreForBusinessApp', 'webApp', 'officeSuiteApp', 'microsoftEdgeApp']).optional().describe('Windows app type'),
  name: z.string().optional().describe('Application name'),
  version: z.string().optional().describe('Application version'),
  assignmentGroups: z.array(z.string()).optional().describe('Target groups for app deployment'),
  assignment: z.object({
    groupIds: z.array(z.string()).describe('Target group IDs'),
    installIntent: z.enum(['available', 'required', 'uninstall', 'availableWithoutEnrollment']).describe('Installation intent'),
    deliveryOptimizationPriority: z.enum(['notConfigured', 'foreground']).optional().describe('Delivery optimization priority'),
  }).optional().describe('App assignment configuration'),
  appInfo: z.object({
    displayName: z.string().describe('Application display name'),
    description: z.string().optional().describe('Application description'),
    publisher: z.string().describe('Application publisher'),
    fileName: z.string().optional().describe('Installation file name'),
    setupFilePath: z.string().optional().describe('Setup file path'),
    installCommandLine: z.string().optional().describe('Install command line'),
    uninstallCommandLine: z.string().optional().describe('Uninstall command line'),
    minimumSupportedOperatingSystem: z.string().optional().describe('Minimum OS version'),
    packageFilePath: z.string().optional().describe('Package file path'),
  }).optional().describe('Application information'),
});

export const intuneWindowsComplianceSchema = z.object({
  action: z.enum(['get_status', 'get_details', 'update_policy', 'force_evaluation', 'get_bitlocker_keys']).describe('Intune Windows compliance action'),
  deviceId: z.string().optional().describe('Device ID for compliance assessment'),
  complianceType: z.enum(['security', 'configuration', 'update', 'bitlocker']).optional().describe('Type of compliance check'),
  policies: z.array(z.string()).optional().describe('Specific policy IDs to assess'),
  complianceData: z.object({
    passwordCompliant: z.boolean().optional().describe('Password compliance status'),
    encryptionCompliant: z.boolean().optional().describe('Encryption compliance status'),
    osVersionCompliant: z.boolean().optional().describe('OS version compliance status'),
    threatProtectionCompliant: z.boolean().optional().describe('Threat protection compliance status'),
    bitlockerCompliant: z.boolean().optional().describe('BitLocker compliance status'),
    firewallCompliant: z.boolean().optional().describe('Firewall compliance status'),
    antivirusCompliant: z.boolean().optional().describe('Antivirus compliance status'),
  }).optional().describe('Compliance assessment data'),
});

// Compliance Framework Management
export const complianceFrameworkSchema = z.object({
  action: z.enum(['list', 'configure', 'status', 'assess', 'activate', 'deactivate']).describe('Compliance framework management action'),
  framework: z.enum(['hitrust', 'iso27001', 'soc2', 'cis']).describe('Compliance framework type'),
  scope: z.array(z.string()).optional().describe('Assessment scope (organization, specific systems)'),
  settings: z.record(z.unknown()).optional().describe('Framework settings'),
});

export const complianceAssessmentSchema = z.object({
  action: z.enum(['create', 'update', 'execute', 'schedule', 'cancel', 'get_results']).describe('Compliance assessment action'),
  assessmentId: z.string().optional().describe('Assessment ID for tracking'),
  framework: z.enum(['hitrust', 'iso27001', 'soc2']).describe('Framework to assess against'),
  scope: z.record(z.unknown()).describe('Assessment scope'),
  settings: z.record(z.unknown()).optional().describe('Assessment settings'),
});

export const complianceMonitoringSchema = z.object({
  action: z.enum(['get_status', 'get_alerts', 'get_trends', 'configure_monitoring']).describe('Compliance monitoring action'),
  framework: z.enum(['hitrust', 'iso27001', 'soc2']).optional().describe('Framework to monitor'),
  filters: z.record(z.unknown()).optional().describe('Monitoring filters'),
  monitoringSettings: z.record(z.unknown()).optional().describe('Monitoring settings'),
});

export const evidenceCollectionSchema = z.object({
  action: z.enum(['get_status', 'schedule', 'collect', 'download']).describe('Evidence collection action'),
  evidenceType: z.enum(['configuration', 'logs', 'policies', 'certificates', 'reports']).optional().describe('Type of evidence to collect'),
  timeRange: z.object({
    start: z.string().describe('Start date (ISO format)'),
    end: z.string().describe('End date (ISO format)'),
  }).optional().describe('Time range for evidence collection'),
  systems: z.array(z.string()).optional().describe('Specific systems to collect evidence from'),
});

export const gapAnalysisSchema = z.object({
  action: z.enum(['generate', 'get_results', 'export']).describe('Gap analysis action'),
  analysisId: z.string().optional().describe('Analysis ID'),
  framework: z.enum(['hitrust', 'iso27001', 'soc2']).describe('Framework for gap analysis'),
  targetFramework: z.enum(['hitrust', 'iso27001', 'soc2']).optional().describe('Target framework for cross-framework mapping'),
  scope: z.object({
    controlIds: z.array(z.string()).optional().describe('Control IDs'),
    categories: z.array(z.string()).optional().describe('Categories'),
  }).optional().describe('Analysis scope'),
  settings: z.object({
    includeRecommendations: z.boolean().describe('Include recommendations'),
    prioritizeByRisk: z.boolean().describe('Prioritize by risk'),
    includeTimeline: z.boolean().describe('Include timeline'),
    includeCostEstimate: z.boolean().describe('Include cost estimate'),
  }).optional().describe('Analysis settings'),
});

export const auditReportSchema = z.object({
  framework: z.enum(['hitrust', 'iso27001', 'soc2', 'cis']).describe('Compliance framework'),
  reportType: z.enum(['full', 'summary', 'gaps', 'evidence', 'executive', 'control_matrix', 'risk_assessment']).describe('Type of audit report'),
  dateRange: z.object({
    startDate: z.string().describe('Start date (ISO format)'),
    endDate: z.string().describe('End date (ISO format)'),
  }).describe('Report time range'),
  format: z.enum(['csv', 'html', 'pdf', 'xlsx']).describe('Report output format'),
  includeEvidence: z.boolean().describe('Include supporting evidence'),
  outputPath: z.string().optional().describe('Output file path'),
  customTemplate: z.string().optional().describe('Custom template path'),
  filters: z.object({
    controlIds: z.array(z.string()).optional().describe('Specific control IDs'),
    riskLevels: z.array(z.enum(['low', 'medium', 'high', 'critical'])).optional().describe('Risk levels to include'),
    implementationStatus: z.array(z.enum(['implemented', 'partiallyImplemented', 'notImplemented', 'notApplicable'])).optional().describe('Implementation status filter'),
    testingStatus: z.array(z.enum(['passed', 'failed', 'notTested', 'inProgress'])).optional().describe('Testing status filter'),
    owners: z.array(z.string()).optional().describe('Control owners'),
  }).optional().describe('Report filters'),
});

export const cisComplianceSchema = z.object({
  action: z.enum(['assess', 'get_benchmark', 'generate_report', 'configure_monitoring', 'remediate']).describe('CIS compliance action'),
  benchmark: z.enum(['windows-10', 'windows-11', 'windows-server-2019', 'windows-server-2022', 'office365', 'azure', 'intune']).optional().describe('CIS benchmark to assess'),
  implementationGroup: z.enum(['1', '2', '3']).optional().describe('Implementation group'),
  controlIds: z.array(z.string()).optional().describe('Specific control IDs'),
  scope: z.object({
    devices: z.array(z.string()).optional().describe('Target devices'),
    users: z.array(z.string()).optional().describe('Target users'),
    policies: z.array(z.string()).optional().describe('Target policies'),
  }).optional().describe('Assessment scope'),
  settings: z.object({
    automated: z.boolean().optional().describe('Automated assessment'),
    generateRemediation: z.boolean().optional().describe('Generate remediation plans'),
    includeEvidence: z.boolean().optional().describe('Include evidence'),
    riskPrioritization: z.boolean().optional().describe('Risk-based prioritization'),
  }).optional().describe('Assessment settings'),
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
  intuneWindowsDevices: intuneWindowsDeviceSchema,
  intuneWindowsPolicies: intuneWindowsPolicySchema,
  intuneWindowsApps: intuneWindowsAppSchema,
  intuneWindowsCompliance: intuneWindowsComplianceSchema,
  complianceFrameworks: complianceFrameworkSchema,
  complianceAssessments: complianceAssessmentSchema,
  complianceMonitoring: complianceMonitoringSchema,
  evidenceCollection: evidenceCollectionSchema,
  gapAnalysis: gapAnalysisSchema,
  auditReports: auditReportSchema,
  cisCompliance: cisComplianceSchema,
} as const;

export { z };
