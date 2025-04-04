import { z } from 'zod';

// Define Zod schemas for validation
export const sharePointSiteSchema = z.object({
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
});

export const sharePointListSchema = z.object({
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
});

export const distributionListSchema = z.object({
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
});

export const securityGroupSchema = z.object({
  action: z.enum(['get', 'create', 'update', 'delete', 'add_members', 'remove_members']),
  groupId: z.string().optional(),
  displayName: z.string().optional(),
  description: z.string().optional(),
  members: z.array(z.string()).optional(),
  settings: z.object({
    securityEnabled: z.boolean().optional(),
    mailEnabled: z.boolean().optional(),
  }).optional(),
});

export const m365GroupSchema = z.object({
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
});

export const exchangeSettingsSchema = z.object({
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
});

export const userManagementSchema = z.object({
  action: z.enum(['get', 'update']),
  userId: z.string(),
  settings: z.record(z.unknown()).optional(),
});

export const offboardingSchema = z.object({
  action: z.enum(['start', 'check', 'complete']),
  userId: z.string(),
  options: z.object({
    revokeAccess: z.boolean().optional(),
    retainMailbox: z.boolean().optional(),
    convertToShared: z.boolean().optional(),
    backupData: z.boolean().optional(),
  }).optional(),
});

// --- Azure AD Schemas ---
export const azureAdRoleSchema = z.object({
  action: z.enum(['list_roles', 'list_role_assignments', 'assign_role', 'remove_role_assignment']),
  roleId: z.string().optional(), // ID of the directoryRole
  principalId: z.string().optional(), // ID of the user, group, or SP
  assignmentId: z.string().optional(), // ID of the role assignment
  filter: z.string().optional(), // OData filter
});

export const azureAdAppSchema = z.object({
  action: z.enum(['list_apps', 'get_app', 'update_app', 'add_owner', 'remove_owner']),
  appId: z.string().optional(), // Object ID of the application
  ownerId: z.string().optional(), // Object ID of the user to add/remove as owner
  appDetails: z.object({ // Details for update_app
    displayName: z.string().optional(),
    signInAudience: z.string().optional(), // e.g., AzureADMyOrg, AzureADMultipleOrgs, AzureADandPersonalMicrosoftAccount, PersonalMicrosoftAccount
    // Add other updatable properties as needed
  }).optional(),
  filter: z.string().optional(), // OData filter for list_apps
});

export const azureAdDeviceSchema = z.object({
  action: z.enum(['list_devices', 'get_device', 'enable_device', 'disable_device', 'delete_device']),
  deviceId: z.string().optional(), // Object ID of the device
  filter: z.string().optional(), // OData filter for list_devices
});

export const azureAdSpSchema = z.object({
  action: z.enum(['list_sps', 'get_sp', 'add_owner', 'remove_owner']),
  spId: z.string().optional(), // Object ID of the Service Principal
  ownerId: z.string().optional(), // Object ID of the user to add/remove as owner
  filter: z.string().optional(), // OData filter for list_sps
});

export const callMicrosoftApiSchema = z.object({
  apiType: z.enum(["graph", "azure"]).describe("Type of Microsoft API: 'graph' or 'azure'."),
  path: z.string().describe("API URL path (e.g., '/users', '/subscriptions/{subId}/resourceGroups')."),
  method: z.enum(["get", "post", "put", "patch", "delete"]).describe("HTTP method."),
  apiVersion: z.string().optional().describe("Azure API version (required for 'azure')."),
  subscriptionId: z.string().optional().describe("Azure Subscription ID (required for most 'azure' paths)."),
  queryParams: z.record(z.string()).optional().describe("Query parameters as key-value pairs."),
  body: z.any().optional().describe("Request body (for POST, PUT, PATCH)."),
});

// --- Security & Compliance Schemas ---
export const auditLogSchema = z.object({
  filter: z.string().optional().describe("OData filter string (e.g., 'activityDateTime ge 2024-01-01T00:00:00Z and initiatedBy/user/id eq \\'...'')"),
  top: z.number().int().positive().optional().describe("Maximum number of records to return."),
});

export const alertSchema = z.object({
  action: z.enum(['list_alerts', 'get_alert']).describe("Action to perform."),
  alertId: z.string().optional().describe("ID of the alert (required for get_alert)."),
  filter: z.string().optional().describe("OData filter string (e.g., 'status eq \\'new\\'')."),
  top: z.number().int().positive().optional().describe("Maximum number of alerts to return."),
});


// Define tools with Zod schemas
export const m365CoreTools = [
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
    name: "call_microsoft_api",
    description: "Call a specific Microsoft Graph or Azure Resource Management API endpoint.",
    inputSchema: {
      type: "object",
      properties: {
        apiType: { type: "string", enum: ["graph", "azure"], description: "API type: 'graph' or 'azure'." },
        path: { type: "string", description: "API URL path (e.g., '/users')." },
        method: { type: "string", enum: ["get", "post", "put", "patch", "delete"], description: "HTTP method." },
        apiVersion: { type: "string", description: "Azure API version (required for 'azure')." },
        subscriptionId: { type: "string", description: "Azure Subscription ID (for 'azure')." },
        queryParams: { type: "object", additionalProperties: { type: "string" }, description: "Query parameters." },
        body: { type: "object", description: "Request body (for POST, PUT, PATCH)." } // Representing 'any' as object for schema
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
          description: "Template to use for site creation"
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
              defaultValue: { type: "any" }
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
