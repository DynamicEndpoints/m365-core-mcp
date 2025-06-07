import { z } from 'zod';

// Define Zod schemas for validation
export const sharePointSiteSchema = {
  action: z.enum(['get', 'create', 'update', 'delete', 'add_users', 'remove_users']),
  siteId: z.string().optional(),
  url: z.string().optional(),
  title: z.string().optional(),
  description: z.string().optional(),
  template: z.string().optional(),
  owners: z.array(z.string()).optional(),
  members: z.array(z.string()).optional(),
  settings: z.object({
    isPublic: z.boolean().optional(),
    allowSharing: z.boolean().optional(),
    storageQuota: z.number().optional(),
  }).optional(),
};

export const sharePointListSchema = {
  action: z.enum(['get', 'create', 'update', 'delete', 'add_items', 'get_items']),
  siteId: z.string(),
  listId: z.string().optional(),
  title: z.string().optional(),
  description: z.string().optional(),
  template: z.string().optional(),
  columns: z.array(z.object({
    name: z.string(),
    type: z.string(),
    required: z.boolean().optional(),
    defaultValue: z.any().optional(),
  })).optional(),
  items: z.array(z.record(z.any())).optional(),
};

export const distributionListSchema = {
  action: z.enum(['get', 'create', 'update', 'delete', 'add_members', 'remove_members']),
  listId: z.string().optional(),
  displayName: z.string().optional(),
  emailAddress: z.string().optional(),
  members: z.array(z.string()).optional(),
  settings: z.object({
    hideFromGAL: z.boolean().optional(),
    requireSenderAuthentication: z.boolean().optional(),
    moderatedBy: z.array(z.string()).optional(),
  }).optional(),
};

export const securityGroupSchema = {
  action: z.enum(['get', 'create', 'update', 'delete', 'add_members', 'remove_members']),
  groupId: z.string().optional(),
  displayName: z.string().optional(),
  description: z.string().optional(),
  members: z.array(z.string()).optional(),
  settings: z.object({
    securityEnabled: z.boolean().optional(),
    mailEnabled: z.boolean().optional(),
  }).optional(),
};

export const m365GroupSchema = {
  action: z.enum(['get', 'create', 'update', 'delete', 'add_members', 'remove_members']),
  groupId: z.string().optional(),
  displayName: z.string().optional(),
  description: z.string().optional(),
  owners: z.array(z.string()).optional(),
  members: z.array(z.string()).optional(),
  settings: z.object({
    visibility: z.enum(['Private', 'Public']).optional(),
    allowExternalSenders: z.boolean().optional(),
    autoSubscribeNewMembers: z.boolean().optional(),
  }).optional(),
};

export const exchangeSettingsSchema = {
  action: z.enum(['get', 'update']),
  settingType: z.enum(['mailbox', 'transport', 'organization', 'retention']),
  target: z.string().optional(),
  settings: z.object({
    automateProcessing: z.object({
      autoReplyEnabled: z.boolean().optional(),
      autoForwardEnabled: z.boolean().optional(),
    }).optional(),
    rules: z.array(z.object({
      name: z.string(),
      conditions: z.record(z.unknown()),
      actions: z.record(z.unknown()),
    })).optional(),
    sharingPolicy: z.object({
      domains: z.array(z.string()),
      enabled: z.boolean(),
    }).optional(),
    retentionTags: z.array(z.object({
      name: z.string(),
      type: z.string(),
      retentionDays: z.number(),
    })).optional(),
  }).optional(),
};

export const userManagementSchema = {
  action: z.enum(['get', 'update']),
  userId: z.string(),
  settings: z.record(z.unknown()).optional(),
};

export const offboardingSchema = {
  action: z.enum(['start', 'check', 'complete']),
  userId: z.string(),
  options: z.object({
    revokeAccess: z.boolean().optional(),
    retainMailbox: z.boolean().optional(),
    convertToShared: z.boolean().optional(),
    backupData: z.boolean().optional(),
  }).optional(),
};

// --- Azure AD Schemas ---
export const azureAdRoleSchema = {
  action: z.enum(['list_roles', 'list_role_assignments', 'assign_role', 'remove_role_assignment']),
  roleId: z.string().optional(), // ID of the directoryRole
  principalId: z.string().optional(), // ID of the user, group, or SP
  assignmentId: z.string().optional(), // ID of the role assignment
  filter: z.string().optional(), // OData filter
};

export const azureAdAppSchema = {
  action: z.enum(['list_apps', 'get_app', 'update_app', 'add_owner', 'remove_owner']),
  appId: z.string().optional(), // Object ID of the application
  ownerId: z.string().optional(), // Object ID of the user to add/remove as owner
  appDetails: z.object({ // Details for update_app
    displayName: z.string().optional(),
    signInAudience: z.string().optional(), // e.g., AzureADMyOrg, AzureADMultipleOrgs, AzureADandPersonalMicrosoftAccount, PersonalMicrosoftAccount
    // Add other updatable properties as needed
  }).optional(),
  filter: z.string().optional(), // OData filter for list_apps
};

export const azureAdDeviceSchema = {
  action: z.enum(['list_devices', 'get_device', 'enable_device', 'disable_device', 'delete_device']),
  deviceId: z.string().optional(), // Object ID of the device
  filter: z.string().optional(), // OData filter for list_devices
};

export const azureAdSpSchema = {
  action: z.enum(['list_sps', 'get_sp', 'add_owner', 'remove_owner']),
  spId: z.string().optional(), // Object ID of the Service Principal
  ownerId: z.string().optional(), // Object ID of the user to add/remove as owner
  filter: z.string().optional(), // OData filter for list_sps
};

export const callMicrosoftApiSchema = {
  apiType: z.enum(["graph", "azure"]).describe("Type of Microsoft API: 'graph' or 'azure'."),
  path: z.string().describe("API URL path (e.g., '/users', '/subscriptions/{subId}/resourceGroups')."),
  method: z.enum(["get", "post", "put", "patch", "delete"]).describe("HTTP method."),
  apiVersion: z.string().optional().describe("Azure API version (required for 'azure')."),
  subscriptionId: z.string().optional().describe("Azure Subscription ID (required for most 'azure' paths)."),
  queryParams: z.record(z.string()).optional().describe("Query parameters as key-value pairs."),
  body: z.any().optional().describe("Request body (for POST, PUT, PATCH)."),
  graphApiVersion: z.enum(["v1.0", "beta"]).optional().default("v1.0").describe("Microsoft Graph API version to use (default: v1.0)."),
  fetchAll: z.boolean().optional().default(false).describe("Set to true to automatically fetch all pages for list results (e.g., users, groups). Default is false."),
  consistencyLevel: z.string().optional().describe("Graph API ConsistencyLevel header. Set to 'eventual' for Graph GET requests using advanced query parameters ($filter, $count, $search, $orderby)."),
};

// --- Security & Compliance Schemas ---
export const auditLogSchema = {
  filter: z.string().optional().describe("OData filter string (e.g., 'activityDateTime ge 2024-01-01T00:00:00Z and initiatedBy/user/id eq \\'...'')"),
  top: z.number().int().positive().optional().describe("Maximum number of records to return."),
};

export const alertSchema = {
  action: z.enum(['list_alerts', 'get_alert']).describe("Action to perform."),
  alertId: z.string().optional().describe("ID of the alert (required for get_alert)."),
  filter: z.string().optional().describe("OData filter string (e.g., 'status eq \\'new\\'')."),
  top: z.number().int().positive().optional().describe("Maximum number of alerts to return."),
};

// DLP Schemas
export const dlpPolicySchema = {
  action: z.enum(['list', 'get', 'create', 'update', 'delete', 'test']),
  policyId: z.string().optional(),
  name: z.string().optional(),
  description: z.string().optional(),
  locations: z.array(z.enum(["Exchange", "SharePoint", "OneDrive", "Teams", "Endpoint"])).optional(),
  settings: z.object({
    mode: z.enum(["Test", "TestWithNotifications", "Enforce"]).optional(),
    priority: z.number().optional(),
    enabled: z.boolean().optional()
  }).optional()
};

export const dlpIncidentSchema = {
  action: z.enum(['list', 'get', 'resolve', 'escalate']),
  incidentId: z.string().optional(),
  filter: z.string().optional(),
  top: z.number().optional()
};

export const sensitivityLabelSchema = {
  action: z.enum(['list', 'get', 'create', 'update', 'delete']),
  labelId: z.string().optional(),
  name: z.string().optional(),
  description: z.string().optional(),
  color: z.string().optional(),
  tooltip: z.string().optional(),
  isActive: z.boolean().optional()
};

// Intune macOS Schemas
export const intuneMacOSDeviceSchema = {
  action: z.enum(['list', 'get', 'wipe', 'restart', 'sync', 'enroll', 'retire', 'remote_lock', 'collect_logs']),
  deviceId: z.string().optional(),
  filter: z.string().optional(),
  top: z.number().optional()
};

export const intuneMacOSPolicySchema = {
  action: z.enum(['list', 'get', 'create', 'update', 'delete', 'assign']),
  policyType: z.enum(['Configuration', 'Compliance', 'Security', 'Update', 'AppProtection']),
  policyId: z.string().optional(),
  name: z.string().optional(),
  settings: z.record(z.unknown()).optional()
};

export const intuneMacOSAppSchema = {
  action: z.enum(['list', 'get', 'deploy', 'update', 'remove', 'sync_status']),
  appId: z.string().optional(),
  appType: z.enum(['webApp', 'officeSuiteApp', 'microsoftEdgeApp', 'microsoftDefenderApp', 'managedIOSApp', 'managedAndroidApp', 'managedMobileLobApp', 'macOSLobApp', 'macOSMicrosoftEdgeApp', 'macOSMicrosoftDefenderApp', 'macOSOfficeSuiteApp', 'macOSWebClip', 'managedApp']).optional(),
  assignment: z.object({
    groupIds: z.array(z.string()),
    installIntent: z.enum(['available', 'required', 'uninstall', 'availableWithoutEnrollment']),
    deliveryOptimizationPriority: z.enum(['notConfigured', 'foreground']).optional()
  }).optional(),
  appInfo: z.object({
    displayName: z.string(),
    description: z.string().optional(),
    publisher: z.string(),
    bundleId: z.string().optional(),
    buildNumber: z.string().optional(),
    versionNumber: z.string().optional(),
    packageFilePath: z.string().optional(),
    minimumSupportedOperatingSystem: z.string().optional(),
    ignoreVersionDetection: z.boolean().optional(),
    installAsManaged: z.boolean().optional()
  }).optional()
};

export const intuneMacOSComplianceSchema = {
  action: z.enum(['get_status', 'get_details', 'update_policy', 'force_evaluation']),
  deviceId: z.string().optional(),
  policyId: z.string().optional(),
  complianceData: z.object({
    passwordCompliant: z.boolean().optional(),
    encryptionCompliant: z.boolean().optional(),
    osVersionCompliant: z.boolean().optional(),
    threatProtectionCompliant: z.boolean().optional(),
    systemIntegrityCompliant: z.boolean().optional(),
    firewallCompliant: z.boolean().optional(),
    gatekeeperCompliant: z.boolean().optional(),
    jailbrokenCompliant: z.boolean().optional()
  }).optional()
};

// Compliance Framework Schemas
export const complianceFrameworkSchema = {
  action: z.enum(['list', 'configure', 'status', 'assess', 'activate', 'deactivate']),
  framework: z.enum(['hitrust', 'iso27001', 'soc2']),
  scope: z.array(z.string()).optional(),
  settings: z.record(z.unknown()).optional()
};

export const complianceAssessmentSchema = {
  action: z.enum(['create', 'update', 'execute', 'schedule', 'cancel', 'get_results']),
  framework: z.enum(['hitrust', 'iso27001', 'soc2']),
  scope: z.record(z.unknown()),
  assessmentId: z.string().optional(),
  settings: z.record(z.unknown()).optional()
};

export const complianceMonitoringSchema = {
  action: z.enum(['get_status', 'get_alerts', 'get_trends', 'configure_monitoring']),
  framework: z.enum(['hitrust', 'iso27001', 'soc2']).optional(),
  filters: z.record(z.unknown()).optional(),
  monitoringSettings: z.record(z.unknown()).optional()
};

export const evidenceCollectionSchema = {
  action: z.enum(['collect', 'schedule', 'get_status', 'download']),
  collectionId: z.string().optional(),
  framework: z.enum(['hitrust', 'iso27001', 'soc2']).optional(),
  controlIds: z.array(z.string()).optional(),
  evidenceTypes: z.array(z.enum(['configuration', 'logs', 'policies', 'screenshots', 'documents'])).optional(),
  settings: z.object({
    automated: z.boolean(),
    scheduledTime: z.string().optional(),
    retention: z.number(),
    encryption: z.boolean(),
    compression: z.boolean()
  }).optional()
};

export const gapAnalysisSchema = {
  action: z.enum(['generate', 'get_results', 'export']),
  framework: z.enum(['hitrust', 'iso27001', 'soc2']),
  analysisId: z.string().optional(),
  targetFramework: z.enum(['hitrust', 'iso27001', 'soc2']).optional(),
  scope: z.object({
    controlIds: z.array(z.string()).optional(),
    categories: z.array(z.string()).optional()
  }).optional(),
  settings: z.object({
    includeRecommendations: z.boolean(),
    prioritizeByRisk: z.boolean(),
    includeTimeline: z.boolean(),
    includeCostEstimate: z.boolean()
  }).optional()
};

export const auditReportSchema = {
  framework: z.enum(['hitrust', 'iso27001', 'soc2']),
  reportType: z.enum(['full', 'summary', 'gaps', 'evidence', 'executive', 'control_matrix', 'risk_assessment']),
  dateRange: z.object({
    startDate: z.string(),
    endDate: z.string()
  }),
  format: z.enum(['csv', 'html', 'pdf', 'xlsx']),
  includeEvidence: z.boolean(),
  outputPath: z.string().optional(),
  customTemplate: z.string().optional(),
  filters: z.object({
    controlIds: z.array(z.string()).optional(),
    riskLevels: z.array(z.enum(['low', 'medium', 'high', 'critical'])).optional(),
    implementationStatus: z.array(z.enum(['implemented', 'partiallyImplemented', 'notImplemented', 'notApplicable'])).optional(),
    testingStatus: z.array(z.enum(['passed', 'failed', 'notTested', 'inProgress'])).optional(),
    owners: z.array(z.string()).optional()
  }).optional()
};


// Define tools with descriptions
export const m365CoreTools = [
  // DLP Management Tools
  {
    name: "manage_dlp_policies",
    description: "Manage Data Loss Prevention (DLP) policies in Microsoft 365",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ['list', 'get', 'create', 'update', 'delete', 'test'],
          description: "Action to perform on DLP policies"
        },
        policyId: {
          type: "string",
          description: "DLP policy ID for operations on existing policies"
        },
        name: {
          type: "string",
          description: "Name of the DLP policy"
        },
        description: {
          type: "string",
          description: "Description of the DLP policy"
        },
        locations: {
          type: "array",
          items: {
            type: "string",
            enum: ["Exchange", "SharePoint", "OneDrive", "Teams", "Endpoint"]
          },
          description: "Locations where the policy applies"
        },
        settings: {
          type: "object",
          properties: {
            mode: {
              type: "string",
              enum: ["Test", "TestWithNotifications", "Enforce"]
            },
            priority: { type: "number" },
            enabled: { type: "boolean" }
          }
        }
      },
      required: ["action"]
    }
  },
  {
    name: "manage_dlp_incidents",
    description: "Manage DLP policy violations and incidents",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ['list', 'get', 'resolve', 'escalate'],
          description: "Action to perform on DLP incidents"
        },
        incidentId: {
          type: "string",
          description: "DLP incident ID"
        },
        dateRange: {
          type: "object",
          properties: {
            startDate: { type: "string" },
            endDate: { type: "string" }
          }
        },
        severity: {
          type: "string",
          enum: ["Low", "Medium", "High", "Critical"]
        },
        status: {
          type: "string",
          enum: ["Active", "Resolved", "InProgress", "Dismissed"]
        },
        policyId: {
          type: "string",
          description: "Filter by specific DLP policy"
        }
      },
      required: ["action"]
    }
  },
  {
    name: "manage_sensitivity_labels",
    description: "Manage Microsoft Information Protection sensitivity labels",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ['list', 'get', 'create', 'update', 'delete', 'apply', 'remove'],
          description: "Action to perform on sensitivity labels"
        },
        labelId: {
          type: "string",
          description: "Sensitivity label ID"
        },
        name: {
          type: "string",
          description: "Name of the sensitivity label"
        },
        description: {
          type: "string",
          description: "Description of the sensitivity label"
        },
        settings: {
          type: "object",
          properties: {
            color: { type: "string" },
            sensitivity: { type: "number" },
            protectionSettings: { type: "object" },
            markingSettings: { type: "object" },
            autoLabelingSettings: { type: "object" }
          }
        },
        targetResource: {
          type: "object",
          properties: {
            resourceType: {
              type: "string",
              enum: ["Email", "Document", "Site", "Container"]
            },
            resourceId: { type: "string" }
          }
        }
      },
      required: ["action"]
    }
  },
  // Intune macOS Management Tools
  {
    name: "manage_intune_macos_devices",
    description: "Manage macOS devices in Microsoft Intune",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ['list', 'get', 'enroll', 'retire', 'wipe', 'restart', 'sync', 'remote_lock', 'collect_logs'],
          description: "Action to perform on macOS devices"
        },
        deviceId: {
          type: "string",
          description: "Device ID for operations on specific devices"
        },
        filter: {
          type: "string",
          description: "OData filter for device queries"
        },
        enrollmentType: {
          type: "string",
          enum: ["UserEnrollment", "DeviceEnrollment", "AutomaticDeviceEnrollment"]
        },
        assignmentTarget: {
          type: "object",
          properties: {
            groupIds: {
              type: "array",
              items: { type: "string" }
            },
            userIds: {
              type: "array",
              items: { type: "string" }
            },
            deviceIds: {
              type: "array",
              items: { type: "string" }
            }
          }
        }
      },
      required: ["action"]
    }
  },
  {
    name: "manage_intune_macos_policies",
    description: "Manage macOS configuration and compliance policies in Intune",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ['list', 'get', 'create', 'update', 'delete', 'assign', 'deploy'],
          description: "Action to perform on macOS policies"
        },
        policyId: {
          type: "string",
          description: "Policy ID for operations on existing policies"
        },
        policyType: {
          type: "string",
          enum: ["Configuration", "Compliance", "Security", "Update", "AppProtection"],
          description: "Type of policy to manage"
        },
        name: {
          type: "string",
          description: "Name of the policy"
        },
        description: {
          type: "string",
          description: "Description of the policy"
        },
        settings: {
          type: "object",
          description: "Policy-specific settings and configurations"
        },
        assignments: {
          type: "array",
          items: { type: "object" },
          description: "Policy assignment targets"
        },
        deploymentSettings: {
          type: "object",
          properties: {
            installBehavior: {
              type: "string",
              enum: ["doNotInstall", "installAsManaged", "installAsUnmanaged"]
            },
            uninstallOnDeviceRemoval: { type: "boolean" },
            installAsManaged: { type: "boolean" }
          }
        }
      },
      required: ["action", "policyType"]
    }
  },
  {
    name: "manage_intune_macos_apps",
    description: "Manage macOS applications in Microsoft Intune",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ['list', 'get', 'deploy', 'update', 'remove', 'sync_status'],
          description: "Action to perform on macOS apps"
        },
        appId: {
          type: "string",
          description: "Application ID"
        },
        appType: {
          type: "string",
          enum: ["webApp", "officeSuiteApp", "microsoftEdgeApp", "microsoftDefenderApp", "macOSLobApp", "macOSMicrosoftEdgeApp", "macOSMicrosoftDefenderApp", "macOSOfficeSuiteApp", "macOSWebClip", "managedApp"]
        },
        assignment: {
          type: "object",
          properties: {
            groupIds: {
              type: "array",
              items: { type: "string" }
            },
            installIntent: {
              type: "string",
              enum: ["available", "required", "uninstall", "availableWithoutEnrollment"]
            },
            deliveryOptimizationPriority: {
              type: "string",
              enum: ["notConfigured", "foreground"]
            }
          }
        },
        appInfo: {
          type: "object",
          properties: {
            displayName: { type: "string" },
            description: { type: "string" },
            publisher: { type: "string" },
            bundleId: { type: "string" },
            buildNumber: { type: "string" },
            versionNumber: { type: "string" },
            packageFilePath: { type: "string" },
            minimumSupportedOperatingSystem: { type: "string" },
            ignoreVersionDetection: { type: "boolean" },
            installAsManaged: { type: "boolean" }
          }
        }
      },
      required: ["action"]
    }
  },
  {
    name: "manage_intune_macos_compliance",
    description: "Monitor and manage macOS device compliance in Intune",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ['get_status', 'get_details', 'update_policy', 'force_evaluation'],
          description: "Action to perform for compliance monitoring"
        },
        deviceId: {
          type: "string",
          description: "Device ID for device-specific operations"
        },
        policyId: {
          type: "string",
          description: "Compliance policy ID"
        },
        complianceData: {
          type: "object",
          properties: {
            passwordCompliant: { type: "boolean" },
            encryptionCompliant: { type: "boolean" },
            osVersionCompliant: { type: "boolean" },
            threatProtectionCompliant: { type: "boolean" },
            systemIntegrityCompliant: { type: "boolean" },
            firewallCompliant: { type: "boolean" },
            gatekeeperCompliant: { type: "boolean" },
            jailbrokenCompliant: { type: "boolean" }
          }
        }
      },
      required: ["action"]
    }
  },
  // Compliance Framework Management Tools
  {
    name: "manage_compliance_frameworks",
    description: "Manage compliance frameworks (HITRUST, ISO 27001, SOC 2)",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ['list', 'configure', 'status', 'assess', 'activate', 'deactivate'],
          description: "Action to perform on compliance frameworks"
        },
        framework: {
          type: "string",
          enum: ["hitrust", "iso27001", "soc2"],
          description: "Compliance framework to manage"
        },
        scope: {
          type: "array",
          items: { type: "string" },
          description: "Scope of controls or domains to include"
        },
        settings: {
          type: "object",
          properties: {
            assessmentPeriod: { type: "string" },
            reportingContacts: {
              type: "array",
              items: { type: "string" }
            },
            customControls: { type: "array" },
            assessmentSettings: {
              type: "object",
              properties: {
                automaticTesting: { type: "boolean" },
                testingFrequency: {
                  type: "string",
                  enum: ["daily", "weekly", "monthly", "quarterly", "annually"]
                },
                evidenceCollection: { type: "boolean" },
                riskAssessment: { type: "boolean" },
                complianceThreshold: { type: "number" }
              }
            }
          }
        }
      },
      required: ["action"]
    }
  },
  {
    name: "manage_compliance_assessments",
    description: "Create and manage compliance assessments",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ['create', 'update', 'execute', 'schedule', 'cancel', 'get_results'],
          description: "Action to perform on assessments"
        },
        assessmentId: {
          type: "string",
          description: "Assessment ID for operations on existing assessments"
        },
        framework: {
          type: "string",
          enum: ["hitrust", "iso27001", "soc2"],
          description: "Compliance framework for the assessment"
        },
        scope: {
          type: "object",
          properties: {
            controlIds: {
              type: "array",
              items: { type: "string" }
            },
            categories: {
              type: "array",
              items: { type: "string" }
            },
            riskLevels: {
              type: "array",
              items: {
                type: "string",
                enum: ["low", "medium", "high", "critical"]
              }
            }
          }
        },
        settings: {
          type: "object",
          properties: {
            assessmentType: {
              type: "string",
              enum: ["full", "partial", "targeted"]
            },
            scheduledDate: { type: "string" },
            automated: { type: "boolean" },
            evidenceCollection: { type: "boolean" },
            notificationSettings: {
              type: "object",
              properties: {
                onCompletion: { type: "boolean" },
                onFailure: { type: "boolean" },
                recipients: {
                  type: "array",
                  items: { type: "string" }
                }
              }
            }
          }
        }
      },
      required: ["action"]
    }
  },
  {
    name: "manage_compliance_monitoring",
    description: "Monitor compliance status and configure alerts",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ['get_status', 'get_alerts', 'get_trends', 'configure_monitoring'],
          description: "Action to perform for compliance monitoring"
        },
        framework: {
          type: "string",
          enum: ["hitrust", "iso27001", "soc2"]
        },
        filters: {
          type: "object",
          properties: {
            riskLevel: {
              type: "array",
              items: {
                type: "string",
                enum: ["low", "medium", "high", "critical"]
              }
            },
            controlDomains: {
              type: "array",
              items: { type: "string" }
            },
            timeRange: {
              type: "object",
              properties: {
                startDate: { type: "string" },
                endDate: { type: "string" }
              }
            }
          }
        },
        monitoringSettings: {
          type: "object",
          properties: {
            enabled: { type: "boolean" },
            frequency: {
              type: "string",
              enum: ["realtime", "hourly", "daily", "weekly"]
            },
            alertThresholds: { type: "array" },
            notifications: {
              type: "object",
              properties: {
                email: { type: "boolean" },
                teams: { type: "boolean" },
                webhook: { type: "string" },
                recipients: {
                  type: "array",
                  items: { type: "string" }
                }
              }
            }
          }
        }
      },
      required: ["action"]
    }
  },
  {
    name: "manage_evidence_collection",
    description: "Collect and manage compliance evidence",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ['collect', 'schedule', 'get_status', 'download'],
          description: "Action to perform for evidence collection"
        },
        collectionId: {
          type: "string",
          description: "Evidence collection ID"
        },
        framework: {
          type: "string",
          enum: ["hitrust", "iso27001", "soc2"]
        },
        controlIds: {
          type: "array",
          items: { type: "string" },
          description: "Specific controls to collect evidence for"
        },
        evidenceTypes: {
          type: "array",
          items: {
            type: "string",
            enum: ["configuration", "logs", "policies", "screenshots", "documents"]
          }
        },
        settings: {
          type: "object",
          properties: {
            automated: { type: "boolean" },
            scheduledTime: { type: "string" },
            retention: { type: "number" },
            encryption: { type: "boolean" },
            compression: { type: "boolean" }
          }
        }
      },
      required: ["action"]
    }
  },
  {
    name: "manage_gap_analysis",
    description: "Perform compliance gap analysis",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ['generate', 'get_results', 'export'],
          description: "Action to perform for gap analysis"
        },
        analysisId: {
          type: "string",
          description: "Gap analysis ID"
        },
        framework: {
          type: "string",
          enum: ["hitrust", "iso27001", "soc2"],
          description: "Primary compliance framework"
        },
        targetFramework: {
          type: "string",
          enum: ["hitrust", "iso27001", "soc2"],
          description: "Target framework for cross-framework mapping"
        },
        scope: {
          type: "object",
          properties: {
            controlIds: {
              type: "array",
              items: { type: "string" }
            },
            categories: {
              type: "array",
              items: { type: "string" }
            }
          }
        },
        settings: {
          type: "object",
          properties: {
            includeRecommendations: { type: "boolean" },
            prioritizeByRisk: { type: "boolean" },
            includeTimeline: { type: "boolean" },
            includeCostEstimate: { type: "boolean" }
          }
        }
      },
      required: ["action"]
    }
  },
  {
    name: "generate_audit_reports",
    description: "Generate comprehensive audit and compliance reports",
    inputSchema: {
      type: "object",
      properties: {
        framework: {
          type: "string",
          enum: ["hitrust", "iso27001", "soc2"],
          description: "Compliance framework for the report"
        },
        reportType: {
          type: "string",
          enum: ["full", "summary", "gaps", "evidence", "executive", "control_matrix", "risk_assessment"],
          description: "Type of report to generate"
        },
        dateRange: {
          type: "object",
          properties: {
            startDate: { type: "string" },
            endDate: { type: "string" }
          },
          description: "Date range for the report"
        },
        format: {
          type: "string",
          enum: ["csv", "html", "pdf", "xlsx"],
          description: "Output format for the report"
        },
        includeEvidence: {
          type: "boolean",
          description: "Include evidence attachments in the report"
        },
        outputPath: {
          type: "string",
          description: "Custom output path for the report file"
        },
        customTemplate: {
          type: "string",
          description: "Custom template to use for report generation"
        },
        filters: {
          type: "object",
          properties: {
            controlIds: {
              type: "array",
              items: { type: "string" }
            },
            riskLevels: {
              type: "array",
              items: {
                type: "string",
                enum: ["low", "medium", "high", "critical"]
              }
            },
            implementationStatus: {
              type: "array",
              items: {
                type: "string",
                enum: ["implemented", "partiallyImplemented", "notImplemented", "notApplicable"]
              }
            },
            testingStatus: {
              type: "array",
              items: {
                type: "string",
                enum: ["passed", "failed", "notTested", "inProgress"]
              }
            },
            owners: {
              type: "array",
              items: { type: "string" }
            }
          }
        }
      },
      required: ["framework", "reportType", "dateRange", "format", "includeEvidence"]
    }
  },
  {
    name: "manage_azure_ad_roles",
    description: "Manage Azure AD directory roles and assignments",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ['list_roles', 'list_role_assignments', 'assign_role', 'remove_role_assignment'],
          description: "Action to perform"
        },
        roleId: {
          type: "string",
          description: "ID of the directory role (required for assign/remove)"
        },
        principalId: {
          type: "string",
          description: "ID of the principal (user, group, SP) to assign/remove role for (required for assign/remove)"
        },
        assignmentId: {
          type: "string",
          description: "ID of the role assignment to remove (required for remove)"
        },
        filter: {
          type: "string",
          description: "OData filter string (optional for list actions)"
        }
      },
      required: ["action"]
    }
  },
  {
    name: "manage_azure_ad_apps",
    description: "Manage Azure AD application registrations",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ['list_apps', 'get_app', 'update_app', 'add_owner', 'remove_owner'],
          description: "Action to perform"
        },
        appId: {
          type: "string",
          description: "Object ID of the application (required for get, update, add/remove owner)"
        },
        ownerId: {
          type: "string",
          description: "Object ID of the user to add/remove as owner (required for add/remove owner)"
        },
        appDetails: {
          type: "object",
          properties: {
             displayName: { type: "string" },
             signInAudience: { type: "string" }
             // Add other properties here
          },
          description: "Details for updating the application (required for update_app)"
        },
        filter: {
          type: "string",
          description: "OData filter string (optional for list_apps)"
        }
      },
      required: ["action"]
    }
  },
  {
    name: "manage_azure_ad_devices",
    description: "Manage Azure AD device objects",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ['list_devices', 'get_device', 'enable_device', 'disable_device', 'delete_device'],
          description: "Action to perform"
        },
        deviceId: {
          type: "string",
          description: "Object ID of the device (required for get, enable, disable, delete)"
        },
        filter: {
          type: "string",
          description: "OData filter string (optional for list_devices)"
        }
      },
      required: ["action"]
    }
  },
  {
    name: "Dynamicendpoint_automation_assistant",
    description: "Acts as a versatile assistant to call any Microsoft Graph or Azure Resource Management API endpoint. Use this for managing users, groups, applications, devices, policies (Conditional Access, Intune Configuration/Compliance), security alerts, audit logs, SharePoint, Exchange, and more.",
    inputSchema: {
      type: "object",
      properties: {
        apiType: { type: "string", enum: ["graph", "azure"], description: "API type: 'graph' or 'azure'." },
        path: { type: "string", description: "API URL path (e.g., '/users')." },
        method: { type: "string", enum: ["get", "post", "put", "patch", "delete"], description: "HTTP method." },
        apiVersion: { type: "string", description: "Azure API version (required for 'azure')." },
        subscriptionId: { type: "string", description: "Azure Subscription ID (for 'azure')." },
        queryParams: { type: "object", additionalProperties: { type: "string" }, description: "Query parameters." },
        body: { type: "object", description: "Request body (for POST, PUT, PATCH)." }, // Representing 'any' as object for schema
        graphApiVersion: { type: "string", enum: ["v1.0", "beta"], description: "Microsoft Graph API version to use (default: v1.0)." },
        fetchAll: { type: "boolean", description: "Set to true to automatically fetch all pages for list results (e.g., users, groups). Default is false." },
        consistencyLevel: { type: "string", description: "Graph API ConsistencyLevel header. Set to 'eventual' for Graph GET requests using advanced query parameters ($filter, $count, $search, $orderby)." }
      },
      required: ["apiType", "path", "method"]
    }
  },
  {
    name: "search_audit_log",
    description: "Search the Azure AD Unified Audit Log.",
    inputSchema: {
      type: "object",
      properties: {
        filter: { type: "string", description: "OData filter string (e.g., 'activityDateTime ge 2024-01-01T00:00:00Z')." },
        top: { type: "number", description: "Maximum number of records." }
      },
      required: [] // Filter is technically optional, though usually needed
    }
  },
  {
    name: "manage_alerts",
    description: "List and view security alerts from Microsoft security products.",
    inputSchema: {
      type: "object",
      properties: {
        action: { type: "string", enum: ['list_alerts', 'get_alert'], description: "Action: list_alerts or get_alert." },
        alertId: { type: "string", description: "ID of the alert (required for get_alert)." },
        filter: { type: "string", description: "OData filter string (e.g., 'status eq \\'new\\'')." },
        top: { type: "number", description: "Maximum number of alerts." }
      },
      required: ["action"]
    }
  },
   {
    name: "manage_service_principals",
    description: "Manage Azure AD Service Principals",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ['list_sps', 'get_sp', 'add_owner', 'remove_owner'],
          description: "Action to perform"
        },
        spId: {
          type: "string",
          description: "Object ID of the Service Principal (required for get, add/remove owner)"
        },
        ownerId: {
          type: "string",
          description: "Object ID of the user to add/remove as owner (required for add/remove owner)"
        },
        filter: {
          type: "string",
          description: "OData filter string (optional for list_sps)"
        }
      },
      required: ["action"]
    }
  },
  {
    name: "manage_sharepoint_sites",
    description: "Manage SharePoint sites",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ["get", "create", "update", "delete", "add_users", "remove_users"],
          description: "Action to perform"
        },
        siteId: {
          type: "string",
          description: "SharePoint site ID for existing site operations"
        },
        url: {
          type: "string",
          description: "URL for the SharePoint site"
        },
        title: {
          type: "string",
          description: "Title for the SharePoint site"
        },
        description: {
          type: "string",
          description: "Description of the SharePoint site"
        },
        template: {
          type: "string",
          description: "Web template ID for site creation. Examples: 'STS#3' (Modern Team Site), 'SITEPAGEPUBLISHING#0' (Communication Site), 'STS#0' (Classic Team Site - default if omitted)."
        },
        owners: {
          type: "array",
          items: { type: "string" },
          description: "List of owner email addresses"
        },
        members: {
          type: "array",
          items: { type: "string" },
          description: "List of member email addresses"
        },
        settings: {
          type: "object",
          properties: {
            isPublic: { type: "boolean" },
            allowSharing: { type: "boolean" },
            storageQuota: { type: "number" }
          }
        }
      },
      required: ["action"]
    }
  },
  {
    name: "manage_sharepoint_lists",
    description: "Manage SharePoint lists",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ["get", "create", "update", "delete", "add_items", "get_items"],
          description: "Action to perform"
        },
        siteId: {
          type: "string",
          description: "SharePoint site ID"
        },
        listId: {
          type: "string",
          description: "SharePoint list ID for existing list operations"
        },
        title: {
          type: "string",
          description: "Title for the SharePoint list"
        },
        description: {
          type: "string",
          description: "Description of the SharePoint list"
        },
        template: {
          type: "string",
          description: "Template to use for list creation"
        },
        columns: {
          type: "array",
          items: {
            type: "object",
            properties: {
              name: { type: "string" },
              type: { type: "string" },
              required: { type: "boolean" },
              defaultValue: {} // Removed invalid type: "any"
            }
          },
          description: "Columns for the SharePoint list"
        },
        items: {
          type: "array",
          items: { type: "object" },
          description: "Items to add to the list"
        }
      },
      required: ["action", "siteId"]
    }
  },
  {
    name: "manage_distribution_lists",
    description: "Manage Microsoft 365 distribution lists",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ["get", "create", "update", "delete", "add_members", "remove_members"],
          description: "Action to perform"
        },
        listId: {
          type: "string",
          description: "Distribution list ID for existing list operations"
        },
        displayName: {
          type: "string",
          description: "Display name for the distribution list"
        },
        emailAddress: {
          type: "string",
          description: "Email address for the distribution list"
        },
        members: {
          type: "array",
          items: { type: "string" },
          description: "List of member email addresses"
        },
        settings: {
          type: "object",
          properties: {
            hideFromGAL: { type: "boolean" },
            requireSenderAuthentication: { type: "boolean" },
            moderatedBy: {
              type: "array",
              items: { type: "string" }
            }
          }
        }
      },
      required: ["action"]
    }
  },
  {
    name: "manage_security_groups",
    description: "Manage Microsoft 365 security groups",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ["get", "create", "update", "delete", "add_members", "remove_members"],
          description: "Action to perform"
        },
        groupId: {
          type: "string",
          description: "Security group ID for existing group operations"
        },
        displayName: {
          type: "string",
          description: "Display name for the security group"
        },
        description: {
          type: "string",
          description: "Description of the security group"
        },
        members: {
          type: "array",
          items: { type: "string" },
          description: "List of member email addresses"
        },
        settings: {
          type: "object",
          properties: {
            securityEnabled: { type: "boolean" },
            mailEnabled: { type: "boolean" }
          }
        }
      },
      required: ["action"]
    }
  },
  {
    name: "manage_m365_groups",
    description: "Manage Microsoft 365 groups",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ["get", "create", "update", "delete", "add_members", "remove_members"],
          description: "Action to perform"
        },
        groupId: {
          type: "string",
          description: "M365 group ID for existing group operations"
        },
        displayName: {
          type: "string",
          description: "Display name for the M365 group"
        },
        description: {
          type: "string",
          description: "Description of the M365 group"
        },
        owners: {
          type: "array",
          items: { type: "string" },
          description: "List of owner email addresses"
        },
        members: {
          type: "array",
          items: { type: "string" },
          description: "List of member email addresses"
        },
        settings: {
          type: "object",
          properties: {
            visibility: {
              type: "string",
              enum: ["Private", "Public"]
            },
            allowExternalSenders: { type: "boolean" },
            autoSubscribeNewMembers: { type: "boolean" }
          }
        }
      },
      required: ["action"]
    }
  },
  {
    name: "manage_exchange_settings",
    description: "Manage Exchange Online settings",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ["get", "update"],
          description: "Action to perform"
        },
        settingType: {
          type: "string",
          enum: ["mailbox", "transport", "organization", "retention"],
          description: "Type of Exchange settings to manage"
        },
        target: {
          type: "string",
          description: "User/Group ID for mailbox settings"
        },
        settings: {
          type: "object",
          properties: {
            automateProcessing: {
              type: "object",
              properties: {
                autoReplyEnabled: { type: "boolean" },
                autoForwardEnabled: { type: "boolean" }
              }
            },
            rules: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  name: { type: "string" },
                  conditions: { type: "object" },
                  actions: { type: "object" }
                }
              }
            },
            sharingPolicy: {
              type: "object",
              properties: {
                domains: {
                  type: "array",
                  items: { type: "string" }
                },
                enabled: { type: "boolean" }
              }
            },
            retentionTags: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  name: { type: "string" },
                  type: { type: "string" },
                  retentionDays: { type: "number" }
                }
              }
            }
          }
        }
      },
      required: ["action", "settingType"]
    }
  },
  {
    name: "manage_user_settings",
    description: "Manage Microsoft 365 user settings and configurations",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ["get", "update"],
          description: "Action to perform"
        },
        userId: {
          type: "string",
          description: "User ID or UPN"
        },
        settings: {
          type: "object",
          description: "User settings to update"
        }
      },
      required: ["action", "userId"]
    }
  },
  {
    name: "manage_offboarding",
    description: "Manage user offboarding processes",
    inputSchema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ["start", "check", "complete"],
          description: "Action to perform"
        },
        userId: {
          type: "string",
          description: "User ID or UPN to offboard"
        },
        options: {
          type: "object",
          properties: {
            revokeAccess: {
              type: "boolean",
              description: "Revoke all access immediately"
            },
            retainMailbox: {
              type: "boolean",
              description: "Retain user mailbox"
            },
            convertToShared: {
              type: "boolean",
              description: "Convert mailbox to shared"
            },
            backupData: {
              type: "boolean",
              description: "Backup user data"
            }
          }
        }
      },
      required: ["action", "userId"]
    }
  }
];
