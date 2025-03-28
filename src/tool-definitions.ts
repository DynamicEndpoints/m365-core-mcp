export const m365CoreTools = [
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
