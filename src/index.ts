#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import {
  CallToolRequestSchema,
  ErrorCode,
  ListResourcesRequestSchema,
  ListResourceTemplatesRequestSchema,
  ListToolsRequestSchema,
  McpError,
  ReadResourceRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';
import {
  UserManagementArgs,
  OffboardingArgs,
  DistributionListArgs,
  SecurityGroupArgs,
  M365GroupArgs,
  ExchangeSettingsArgs,
  SharePointSiteArgs,
  SharePointListArgs,
} from './types.js';

// Define Zod schemas for validation
const sharePointSiteSchema = z.object({
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

const sharePointListSchema = z.object({
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

const distributionListSchema = z.object({
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

const securityGroupSchema = z.object({
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

const m365GroupSchema = z.object({
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

const exchangeSettingsSchema = z.object({
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

const userManagementSchema = z.object({
  action: z.enum(['get', 'update']),
  userId: z.string(),
  settings: z.record(z.unknown()).optional(),
});

const offboardingSchema = z.object({
  action: z.enum(['start', 'check', 'complete']),
  userId: z.string(),
  options: z.object({
    revokeAccess: z.boolean().optional(),
    retainMailbox: z.boolean().optional(),
    convertToShared: z.boolean().optional(),
    backupData: z.boolean().optional(),
  }).optional(),
});

// Define tools with Zod schemas
const m365CoreTools = [
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

// Environment validation
const MS_TENANT_ID = process.env.MS_TENANT_ID ?? '';
const MS_CLIENT_ID = process.env.MS_CLIENT_ID ?? '';
const MS_CLIENT_SECRET = process.env.MS_CLIENT_SECRET ?? '';

if (!MS_TENANT_ID || !MS_CLIENT_ID || !MS_CLIENT_SECRET) {
  throw new Error('Required environment variables (MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET) are missing');
}

class M365CoreServer {
  private server: Server;
  private graphClient: Client;
  private token: string | null = null;

  constructor() {
    this.server = new Server(
      {
        name: 'm365-core-server',
        version: '1.0.0',
      },
      {
        capabilities: {
          resources: {},
          tools: {},
        },
      }
    );

    this.graphClient = Client.init({
      authProvider: async (callback) => {
        try {
          if (!this.token) {
            this.token = await this.getAccessToken();
          }
          callback(null, this.token);
        } catch (error) {
          callback(error as Error, null);
        }
      }
    });

    this.setupHandlers();
    
    // Error handling
    this.server.onerror = (error) => console.error('[MCP Error]', error);
    process.on('SIGINT', async () => {
      await this.server.close();
      process.exit(0);
    });
  }

  private async getAccessToken(): Promise<string> {
    const tokenEndpoint = `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`;
    const scope = 'https://graph.microsoft.com/.default';

    const params = new URLSearchParams();
    params.append('client_id', MS_CLIENT_ID);
    params.append('client_secret', MS_CLIENT_SECRET);
    params.append('grant_type', 'client_credentials');
    params.append('scope', scope);

    const response = await fetch(tokenEndpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: params,
    });

    if (!response.ok) {
      throw new Error('Failed to get access token');
    }

    const data = await response.json();
    return data.access_token;
  }

  private setupHandlers(): void {
    // List available tools
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: m365CoreTools
    }));

    // Resource handlers
    this.server.setRequestHandler(ListResourcesRequestSchema, async () => ({
      resources: [
        {
          uri: 'm365://users/current',
          name: 'Current user information',
          mimeType: 'application/json',
          description: 'Information about the currently authenticated user',
        },
        {
          uri: 'm365://tenant/info',
          name: 'Tenant information',
          mimeType: 'application/json',
          description: 'Information about the Microsoft 365 tenant',
        },
        {
          uri: 'm365://sharepoint/sites',
          name: 'SharePoint sites',
          mimeType: 'application/json',
          description: 'List of SharePoint sites in the tenant',
        },
        {
          uri: 'm365://sharepoint/admin/settings',
          name: 'SharePoint admin settings',
          mimeType: 'application/json',
          description: 'SharePoint admin settings for the tenant',
        },
      ],
    }));

    // Resource template handlers
    this.server.setRequestHandler(ListResourceTemplatesRequestSchema, async () => ({
      resourceTemplates: [
        {
          uriTemplate: 'm365://users/{userId}',
          name: 'User information',
          mimeType: 'application/json',
          description: 'Information about a specific user',
        },
        {
          uriTemplate: 'm365://groups/{groupId}',
          name: 'Group information',
          mimeType: 'application/json',
          description: 'Information about a specific group',
        },
        {
          uriTemplate: 'm365://sharepoint/sites/{siteId}',
          name: 'SharePoint site information',
          mimeType: 'application/json',
          description: 'Information about a specific SharePoint site',
        },
        {
          uriTemplate: 'm365://sharepoint/sites/{siteId}/lists',
          name: 'SharePoint lists',
          mimeType: 'application/json',
          description: 'Lists in a specific SharePoint site',
        },
        {
          uriTemplate: 'm365://sharepoint/sites/{siteId}/lists/{listId}',
          name: 'SharePoint list information',
          mimeType: 'application/json',
          description: 'Information about a specific SharePoint list',
        },
        {
          uriTemplate: 'm365://sharepoint/sites/{siteId}/lists/{listId}/items',
          name: 'SharePoint list items',
          mimeType: 'application/json',
          description: 'Items in a specific SharePoint list',
        },
      ],
    }));

    // Read resource handler
    this.server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
      try {
        const uri = request.params.uri;
        
        // Static resources
        if (uri === 'm365://users/current') {
          const currentUser = await this.graphClient
            .api('/me')
            .get();
          
          return {
            contents: [
              {
                uri,
                mimeType: 'application/json',
                text: JSON.stringify(currentUser, null, 2),
              },
            ],
          };
        }
        
        if (uri === 'm365://tenant/info') {
          const tenantInfo = await this.graphClient
            .api('/organization')
            .get();
          
          return {
            contents: [
              {
                uri,
                mimeType: 'application/json',
                text: JSON.stringify(tenantInfo, null, 2),
              },
            ],
          };
        }
        
        if (uri === 'm365://sharepoint/sites') {
          const sites = await this.graphClient
            .api('/sites?search=*')
            .get();
          
          return {
            contents: [
              {
                uri,
                mimeType: 'application/json',
                text: JSON.stringify(sites, null, 2),
              },
            ],
          };
        }
        
        if (uri === 'm365://sharepoint/admin/settings') {
          const settings = await this.graphClient
            .api('/admin/sharepoint/settings')
            .get();
          
          return {
            contents: [
              {
                uri,
                mimeType: 'application/json',
                text: JSON.stringify(settings, null, 2),
              },
            ],
          };
        }
        
        // Dynamic resources
        const userMatch = uri.match(/^m365:\/\/users\/([^/]+)$/);
        if (userMatch) {
          const userId = decodeURIComponent(userMatch[1]);
          const user = await this.graphClient
            .api(`/users/${userId}`)
            .get();
          
          return {
            contents: [
              {
                uri,
                mimeType: 'application/json',
                text: JSON.stringify(user, null, 2),
              },
            ],
          };
        }
        
        const groupMatch = uri.match(/^m365:\/\/groups\/([^/]+)$/);
        if (groupMatch) {
          const groupId = decodeURIComponent(groupMatch[1]);
          const group = await this.graphClient
            .api(`/groups/${groupId}`)
            .get();
          
          return {
            contents: [
              {
                uri,
                mimeType: 'application/json',
                text: JSON.stringify(group, null, 2),
              },
            ],
          };
        }
        
        // SharePoint site resources
        const siteMatch = uri.match(/^m365:\/\/sharepoint\/sites\/([^/]+)$/);
        if (siteMatch) {
          const siteId = decodeURIComponent(siteMatch[1]);
          const site = await this.graphClient
            .api(`/sites/${siteId}`)
            .get();
          
          return {
            contents: [
              {
                uri,
                mimeType: 'application/json',
                text: JSON.stringify(site, null, 2),
              },
            ],
          };
        }
        
        // SharePoint lists resources
        const listsMatch = uri.match(/^m365:\/\/sharepoint\/sites\/([^/]+)\/lists$/);
        if (listsMatch) {
          const siteId = decodeURIComponent(listsMatch[1]);
          const lists = await this.graphClient
            .api(`/sites/${siteId}/lists`)
            .get();
          
          return {
            contents: [
              {
                uri,
                mimeType: 'application/json',
                text: JSON.stringify(lists, null, 2),
              },
            ],
          };
        }
        
        // SharePoint list resource
        const listMatch = uri.match(/^m365:\/\/sharepoint\/sites\/([^/]+)\/lists\/([^/]+)$/);
        if (listMatch) {
          const siteId = decodeURIComponent(listMatch[1]);
          const listId = decodeURIComponent(listMatch[2]);
          const list = await this.graphClient
            .api(`/sites/${siteId}/lists/${listId}`)
            .get();
          
          return {
            contents: [
              {
                uri,
                mimeType: 'application/json',
                text: JSON.stringify(list, null, 2),
              },
            ],
          };
        }
        
        // SharePoint list items resource
        const listItemsMatch = uri.match(/^m365:\/\/sharepoint\/sites\/([^/]+)\/lists\/([^/]+)\/items$/);
        if (listItemsMatch) {
          const siteId = decodeURIComponent(listItemsMatch[1]);
          const listId = decodeURIComponent(listItemsMatch[2]);
          const items = await this.graphClient
            .api(`/sites/${siteId}/lists/${listId}/items?expand=fields`)
            .get();
          
          return {
            contents: [
              {
                uri,
                mimeType: 'application/json',
                text: JSON.stringify(items, null, 2),
              },
            ],
          };
        }
        
        throw new McpError(ErrorCode.InvalidRequest, `Resource not found: ${uri}`);
      } catch (error) {
        if (error instanceof McpError) {
          throw error;
        }
        throw new McpError(
          ErrorCode.InternalError,
          `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
        );
      }
    });

    // Handle tool calls
    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      try {
        const { name, arguments: args } = request.params;

        if (!args || typeof args !== 'object') {
          throw new McpError(ErrorCode.InvalidParams, 'Missing or invalid arguments');
        }

        switch (name) {
          case 'manage_distribution_lists': {
            try {
              const dlArgs = distributionListSchema.parse(args);
              return await this.handleDistributionList(dlArgs);
            } catch (error) {
              if (error instanceof z.ZodError) {
                throw new McpError(
                  ErrorCode.InvalidParams,
                  `Invalid distribution list parameters: ${error.errors.map(e => e.message).join(', ')}`
                );
              }
              throw error;
            }
          }

          case 'manage_security_groups': {
            try {
              const sgArgs = securityGroupSchema.parse(args);
              return await this.handleSecurityGroup(sgArgs);
            } catch (error) {
              if (error instanceof z.ZodError) {
                throw new McpError(
                  ErrorCode.InvalidParams,
                  `Invalid security group parameters: ${error.errors.map(e => e.message).join(', ')}`
                );
              }
              throw error;
            }
          }

          case 'manage_m365_groups': {
            try {
              const m365Args = m365GroupSchema.parse(args);
              return await this.handleM365Group(m365Args);
            } catch (error) {
              if (error instanceof z.ZodError) {
                throw new McpError(
                  ErrorCode.InvalidParams,
                  `Invalid M365 group parameters: ${error.errors.map(e => e.message).join(', ')}`
                );
              }
              throw error;
            }
          }

          case 'manage_exchange_settings': {
            try {
              const exchangeArgs = exchangeSettingsSchema.parse(args);
              return await this.handleExchangeSettings(exchangeArgs);
            } catch (error) {
              if (error instanceof z.ZodError) {
                throw new McpError(
                  ErrorCode.InvalidParams,
                  `Invalid Exchange settings parameters: ${error.errors.map(e => e.message).join(', ')}`
                );
              }
              throw error;
            }
          }

          case 'manage_user_settings': {
            try {
              const userArgs = userManagementSchema.parse(args);
              return await this.handleUserSettings({
                action: userArgs.action,
                userPrincipalName: userArgs.userId,
                settings: userArgs.settings
              });
            } catch (error) {
              if (error instanceof z.ZodError) {
                throw new McpError(
                  ErrorCode.InvalidParams,
                  `Invalid user management parameters: ${error.errors.map(e => e.message).join(', ')}`
                );
              }
              throw error;
            }
          }

          case 'manage_offboarding': {
            try {
              const offboardingArgs = offboardingSchema.parse(args);
              return await this.handleOffboarding(offboardingArgs);
            } catch (error) {
              if (error instanceof z.ZodError) {
                throw new McpError(
                  ErrorCode.InvalidParams,
                  `Invalid offboarding parameters: ${error.errors.map(e => e.message).join(', ')}`
                );
              }
              throw error;
            }
          }
          
          case 'manage_sharepoint_sites': {
            try {
              const siteArgs = sharePointSiteSchema.parse(args);
              return await this.handleSharePointSite(siteArgs);
            } catch (error) {
              if (error instanceof z.ZodError) {
                throw new McpError(
                  ErrorCode.InvalidParams,
                  `Invalid SharePoint site parameters: ${error.errors.map(e => e.message).join(', ')}`
                );
              }
              throw error;
            }
          }
          
          case 'manage_sharepoint_lists': {
            try {
              const listArgs = sharePointListSchema.parse(args);
              return await this.handleSharePointList(listArgs);
            } catch (error) {
              if (error instanceof z.ZodError) {
                throw new McpError(
                  ErrorCode.InvalidParams,
                  `Invalid SharePoint list parameters: ${error.errors.map(e => e.message).join(', ')}`
                );
              }
              throw error;
            }
          }
        }

        throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${name}`);
      } catch (error) {
        if (error instanceof McpError) {
          throw error;
        }
        throw new McpError(
          ErrorCode.InternalError,
          `Error executing tool: ${error instanceof Error ? error.message : 'Unknown error'}`
        );
      }
    });
  }

  private validateDistributionListArgs(args: Record<string, unknown>): DistributionListArgs {
    if (!args.action || typeof args.action !== 'string') {
      throw new McpError(ErrorCode.InvalidParams, 'Missing required distribution list parameters');
    }
    return {
      action: args.action as DistributionListArgs['action'],
      listId: typeof args.listId === 'string' ? args.listId : undefined,
      displayName: typeof args.displayName === 'string' ? args.displayName : undefined,
      emailAddress: typeof args.emailAddress === 'string' ? args.emailAddress : undefined,
      members: Array.isArray(args.members) ? args.members.map(String) : undefined,
      settings: typeof args.settings === 'object' ? args.settings as DistributionListArgs['settings'] : undefined,
    };
  }

  private validateSecurityGroupArgs(args: Record<string, unknown>): SecurityGroupArgs {
    if (!args.action || typeof args.action !== 'string') {
      throw new McpError(ErrorCode.InvalidParams, 'Missing required security group parameters');
    }
    return {
      action: args.action as SecurityGroupArgs['action'],
      groupId: typeof args.groupId === 'string' ? args.groupId : undefined,
      displayName: typeof args.displayName === 'string' ? args.displayName : undefined,
      description: typeof args.description === 'string' ? args.description : undefined,
      members: Array.isArray(args.members) ? args.members.map(String) : undefined,
      settings: typeof args.settings === 'object' ? args.settings as SecurityGroupArgs['settings'] : undefined,
    };
  }

  private validateM365GroupArgs(args: Record<string, unknown>): M365GroupArgs {
    if (!args.action || typeof args.action !== 'string') {
      throw new McpError(ErrorCode.InvalidParams, 'Missing required M365 group parameters');
    }
    return {
      action: args.action as M365GroupArgs['action'],
      groupId: typeof args.groupId === 'string' ? args.groupId : undefined,
      displayName: typeof args.displayName === 'string' ? args.displayName : undefined,
      description: typeof args.description === 'string' ? args.description : undefined,
      owners: Array.isArray(args.owners) ? args.owners.map(String) : undefined,
      members: Array.isArray(args.members) ? args.members.map(String) : undefined,
      settings: typeof args.settings === 'object' ? args.settings as M365GroupArgs['settings'] : undefined,
    };
  }

  private validateExchangeSettingsArgs(args: Record<string, unknown>): ExchangeSettingsArgs {
    if (!args.action || !args.settingType || typeof args.action !== 'string' || typeof args.settingType !== 'string') {
      throw new McpError(ErrorCode.InvalidParams, 'Missing required Exchange settings parameters');
    }
    return {
      action: args.action as ExchangeSettingsArgs['action'],
      settingType: args.settingType as ExchangeSettingsArgs['settingType'],
      target: typeof args.target === 'string' ? args.target : undefined,
      settings: typeof args.settings === 'object' ? args.settings as ExchangeSettingsArgs['settings'] : undefined,
    };
  }

  private validateUserManagementArgs(args: Record<string, unknown>): UserManagementArgs {
    if (!args.action || !args.userId || typeof args.action !== 'string' || typeof args.userId !== 'string') {
      throw new McpError(ErrorCode.InvalidParams, 'Missing required user management parameters');
    }
    return {
      action: args.action as UserManagementArgs['action'],
      userPrincipalName: args.userId,
      settings: typeof args.settings === 'object' ? args.settings as Record<string, unknown> : undefined,
    };
  }

  private validateOffboardingArgs(args: Record<string, unknown>): OffboardingArgs {
    if (!args.action || !args.userId || typeof args.action !== 'string' || typeof args.userId !== 'string') {
      throw new McpError(ErrorCode.InvalidParams, 'Missing required offboarding parameters');
    }
    return {
      action: args.action as OffboardingArgs['action'],
      userId: args.userId,
      options: typeof args.options === 'object' ? args.options as OffboardingArgs['options'] : undefined,
    };
  }

  private async handleDistributionList(args: DistributionListArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    switch (args.action) {
      case 'get': {
        const list = await this.graphClient
          .api(`/groups/${args.listId}`)
          .get();
        return { content: [{ type: 'text', text: JSON.stringify(list, null, 2) }] };
      }
      case 'create': {
        const list = await this.graphClient
          .api('/groups')
          .post({
            displayName: args.displayName,
            mailEnabled: true,
            securityEnabled: false,
            mailNickname: args.emailAddress?.split('@')[0],
            ...args.settings,
          });
        return { content: [{ type: 'text', text: JSON.stringify(list, null, 2) }] };
      }
      case 'update': {
        await this.graphClient
          .api(`/groups/${args.listId}`)
          .patch({
            displayName: args.displayName,
            ...args.settings,
          });
        return { content: [{ type: 'text', text: 'Distribution list updated successfully' }] };
      }
      case 'delete': {
        await this.graphClient
          .api(`/groups/${args.listId}`)
          .delete();
        return { content: [{ type: 'text', text: 'Distribution list deleted successfully' }] };
      }
      case 'add_members': {
        if (!args.members?.length) {
          throw new McpError(ErrorCode.InvalidParams, 'No members specified to add');
        }
        for (const member of args.members) {
          await this.graphClient
            .api(`/groups/${args.listId}/members/$ref`)
            .post({
              '@odata.id': `https://graph.microsoft.com/v1.0/users/${member}`,
            });
        }
        return { content: [{ type: 'text', text: 'Members added successfully' }] };
      }
      case 'remove_members': {
        if (!args.members?.length) {
          throw new McpError(ErrorCode.InvalidParams, 'No members specified to remove');
        }
        for (const member of args.members) {
          await this.graphClient
            .api(`/groups/${args.listId}/members/${member}/$ref`)
            .delete();
        }
        return { content: [{ type: 'text', text: 'Members removed successfully' }] };
      }
      default:
        throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
    }
  }

  private async handleSecurityGroup(args: SecurityGroupArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    switch (args.action) {
      case 'get': {
        const group = await this.graphClient
          .api(`/groups/${args.groupId}`)
          .get();
        return { content: [{ type: 'text', text: JSON.stringify(group, null, 2) }] };
      }
      case 'create': {
        const group = await this.graphClient
          .api('/groups')
          .post({
            displayName: args.displayName,
            description: args.description,
            securityEnabled: true,
            mailEnabled: args.settings?.mailEnabled ?? false,
            mailNickname: args.displayName?.replace(/\s+/g, '').toLowerCase(),
          });
        return { content: [{ type: 'text', text: JSON.stringify(group, null, 2) }] };
      }
      case 'update': {
        await this.graphClient
          .api(`/groups/${args.groupId}`)
          .patch({
            displayName: args.displayName,
            description: args.description,
            ...args.settings,
          });
        return { content: [{ type: 'text', text: 'Security group updated successfully' }] };
      }
      case 'delete': {
        await this.graphClient
          .api(`/groups/${args.groupId}`)
          .delete();
        return { content: [{ type: 'text', text: 'Security group deleted successfully' }] };
      }
      case 'add_members':
      case 'remove_members': {
        if (!args.members?.length) {
          throw new McpError(ErrorCode.InvalidParams, 'No members specified');
        }
        for (const member of args.members) {
          if (args.action === 'add_members') {
            await this.graphClient
              .api(`/groups/${args.groupId}/members/$ref`)
              .post({
                '@odata.id': `https://graph.microsoft.com/v1.0/users/${member}`,
              });
          } else {
            await this.graphClient
              .api(`/groups/${args.groupId}/members/${member}/$ref`)
              .delete();
          }
        }
        return { content: [{ type: 'text', text: `Members ${args.action === 'add_members' ? 'added' : 'removed'} successfully` }] };
      }
      default:
        throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
    }
  }

  private async handleM365Group(args: M365GroupArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    switch (args.action) {
      case 'get': {
        const group = await this.graphClient
          .api(`/groups/${args.groupId}`)
          .get();
        return { content: [{ type: 'text', text: JSON.stringify(group, null, 2) }] };
      }
      case 'create': {
        const group = await this.graphClient
          .api('/groups')
          .post({
            displayName: args.displayName,
            description: args.description,
            groupTypes: ['Unified'],
            mailEnabled: true,
            securityEnabled: false,
            mailNickname: args.displayName?.replace(/\s+/g, '').toLowerCase(),
            visibility: args.settings?.visibility?.toLowerCase(),
            ...args.settings,
          });
        return { content: [{ type: 'text', text: JSON.stringify(group, null, 2) }] };
      }
      case 'update': {
        await this.graphClient
          .api(`/groups/${args.groupId}`)
          .patch({
            displayName: args.displayName,
            description: args.description,
            ...args.settings,
          });
        return { content: [{ type: 'text', text: 'M365 group updated successfully' }] };
      }
      case 'delete': {
        await this.graphClient
          .api(`/groups/${args.groupId}`)
          .delete();
        return { content: [{ type: 'text', text: 'M365 group deleted successfully' }] };
      }
      case 'add_members':
      case 'remove_members': {
        if (!args.members?.length) {
          throw new McpError(ErrorCode.InvalidParams, 'No members specified');
        }
        for (const member of args.members) {
          if (args.action === 'add_members') {
            await this.graphClient
              .api(`/groups/${args.groupId}/members/$ref`)
              .post({
                '@odata.id': `https://graph.microsoft.com/v1.0/users/${member}`,
              });
          } else {
            await this.graphClient
              .api(`/groups/${args.groupId}/members/${member}/$ref`)
              .delete();
          }
        }
        return { content: [{ type: 'text', text: `Members ${args.action === 'add_members' ? 'added' : 'removed'} successfully` }] };
      }
      default:
        throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
    }
  }

  private async handleExchangeSettings(args: ExchangeSettingsArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    switch (args.settingType) {
      case 'mailbox': {
        if (args.action === 'get') {
          const settings = await this.graphClient
            .api(`/users/${args.target}/mailboxSettings`)
            .get();
          return { content: [{ type: 'text', text: JSON.stringify(settings, null, 2) }] };
        } else {
          await this.graphClient
            .api(`/users/${args.target}/mailboxSettings`)
            .patch(args.settings);
          return { content: [{ type: 'text', text: 'Mailbox settings updated successfully' }] };
        }
      }
      case 'transport': {
        if (args.action === 'get') {
          const rules = await this.graphClient
            .api('/admin/transportRules')
            .get();
          return { content: [{ type: 'text', text: JSON.stringify(rules, null, 2) }] };
        } else {
          await this.graphClient
            .api('/admin/transportRules')
            .post(args.settings?.rules);
          return { content: [{ type: 'text', text: 'Transport rules updated successfully' }] };
        }
      }
      case 'organization': {
        if (args.action === 'get') {
          const settings = await this.graphClient
            .api('/admin/organization/settings')
            .get();
          return { content: [{ type: 'text', text: JSON.stringify(settings, null, 2) }] };
        } else {
          await this.graphClient
            .api('/admin/organization/settings')
            .patch(args.settings);
          return { content: [{ type: 'text', text: 'Organization settings updated successfully' }] };
        }
      }
      case 'retention': {
        if (args.action === 'get') {
          const policies = await this.graphClient
            .api('/admin/retentionPolicies')
            .get();
          return { content: [{ type: 'text', text: JSON.stringify(policies, null, 2) }] };
        } else {
          await this.graphClient
            .api('/admin/retentionPolicies')
            .post(args.settings?.retentionTags);
          return { content: [{ type: 'text', text: 'Retention policies updated successfully' }] };
        }
      }
      default:
        throw new McpError(ErrorCode.InvalidParams, `Invalid setting type: ${args.settingType}`);
    }
  }

  private async handleUserSettings(args: UserManagementArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    if (args.action === 'get') {
      const settings = await this.graphClient
        .api(`/users/${args.userPrincipalName}`)
        .get();
      return { content: [{ type: 'text', text: JSON.stringify(settings, null, 2) }] };
    } else {
      await this.graphClient
        .api(`/users/${args.userPrincipalName}`)
        .patch(args.settings);
      return { content: [{ type: 'text', text: 'User settings updated successfully' }] };
    }
  }

  private async handleOffboarding(args: OffboardingArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    switch (args.action) {
      case 'start': {
        // Block sign-ins
        await this.graphClient
          .api(`/users/${args.userId}`)
          .patch({ accountEnabled: false });

        if (args.options?.revokeAccess) {
          // Revoke all refresh tokens
          await this.graphClient
            .api(`/users/${args.userId}/revokeSignInSessions`)
            .post({});
        }

        if (args.options?.backupData) {
          // Trigger backup
          await this.graphClient
            .api(`/users/${args.userId}/drive/content`)
            .get();
        }

        return { content: [{ type: 'text', text: 'Offboarding process started successfully' }] };
      }
      case 'check': {
        const status = await this.graphClient
          .api(`/users/${args.userId}`)
          .get();
        return { content: [{ type: 'text', text: JSON.stringify(status, null, 2) }] };
      }
      case 'complete': {
        if (args.options?.convertToShared) {
          // Convert to shared mailbox
          await this.graphClient
            .api(`/users/${args.userId}/mailbox/convert`)
            .post({});
        } else if (!args.options?.retainMailbox) {
          // Delete user if not retaining mailbox
          await this.graphClient
            .api(`/users/${args.userId}`)
            .delete();
        }
        return { content: [{ type: 'text', text: 'Offboarding process completed successfully' }] };
      }
      default:
        throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
    }
  }

  private async handleSharePointSite(args: SharePointSiteArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    switch (args.action) {
      case 'get': {
        const site = await this.graphClient
          .api(`/sites/${args.siteId}`)
          .get();
        return { content: [{ type: 'text', text: JSON.stringify(site, null, 2) }] };
      }
      case 'create': {
        // Create a new SharePoint site
        const site = await this.graphClient
          .api('/sites/add')
          .post({
            displayName: args.title,
            description: args.description,
            webTemplate: args.template || 'STS#0', // Team site template
            url: args.url,
          });
        
        // Apply settings if provided
        if (args.settings) {
          await this.graphClient
            .api(`/sites/${site.id}/settings`)
            .patch({
              isPublic: args.settings.isPublic,
              sharingCapability: args.settings.allowSharing ? 'ExternalUserSharingOnly' : 'Disabled',
              storageQuota: args.settings.storageQuota,
            });
        }
        
        // Add owners if provided
        if (args.owners?.length) {
          for (const owner of args.owners) {
            await this.graphClient
              .api(`/sites/${site.id}/owners/$ref`)
              .post({
                '@odata.id': `https://graph.microsoft.com/v1.0/users/${owner}`,
              });
          }
        }
        
        // Add members if provided
        if (args.members?.length) {
          for (const member of args.members) {
            await this.graphClient
              .api(`/sites/${site.id}/members/$ref`)
              .post({
                '@odata.id': `https://graph.microsoft.com/v1.0/users/${member}`,
              });
          }
        }
        
        return { content: [{ type: 'text', text: JSON.stringify(site, null, 2) }] };
      }
      case 'update': {
        // Update site properties
        await this.graphClient
          .api(`/sites/${args.siteId}`)
          .patch({
            displayName: args.title,
            description: args.description,
          });
        
        // Update settings if provided
        if (args.settings) {
          await this.graphClient
            .api(`/sites/${args.siteId}/settings`)
            .patch({
              isPublic: args.settings.isPublic,
              sharingCapability: args.settings.allowSharing ? 'ExternalUserSharingOnly' : 'Disabled',
              storageQuota: args.settings.storageQuota,
            });
        }
        
        return { content: [{ type: 'text', text: 'SharePoint site updated successfully' }] };
      }
      case 'delete': {
        await this.graphClient
          .api(`/sites/${args.siteId}`)
          .delete();
        return { content: [{ type: 'text', text: 'SharePoint site deleted successfully' }] };
      }
      case 'add_users': {
        if (!args.members?.length) {
          throw new McpError(ErrorCode.InvalidParams, 'No users specified to add');
        }
        
        for (const member of args.members) {
          await this.graphClient
            .api(`/sites/${args.siteId}/members/$ref`)
            .post({
              '@odata.id': `https://graph.microsoft.com/v1.0/users/${member}`,
            });
        }
        
        return { content: [{ type: 'text', text: 'Users added to SharePoint site successfully' }] };
      }
      case 'remove_users': {
        if (!args.members?.length) {
          throw new McpError(ErrorCode.InvalidParams, 'No users specified to remove');
        }
        
        for (const member of args.members) {
          await this.graphClient
            .api(`/sites/${args.siteId}/members/${member}/$ref`)
            .delete();
        }
        
        return { content: [{ type: 'text', text: 'Users removed from SharePoint site successfully' }] };
      }
      default:
        throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
    }
  }

  private async handleSharePointList(args: SharePointListArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    switch (args.action) {
      case 'get': {
        const list = await this.graphClient
          .api(`/sites/${args.siteId}/lists/${args.listId}`)
          .get();
        return { content: [{ type: 'text', text: JSON.stringify(list, null, 2) }] };
      }
      case 'create': {
        // Create a new list
        const list = await this.graphClient
          .api(`/sites/${args.siteId}/lists`)
          .post({
            displayName: args.title,
            description: args.description,
            template: args.template || 'genericList',
          });
        
        // Add columns if provided
        if (args.columns?.length) {
          for (const column of args.columns) {
            await this.graphClient
              .api(`/sites/${args.siteId}/lists/${list.id}/columns`)
              .post({
                name: column.name,
                columnType: column.type,
                required: column.required || false,
                defaultValue: column.defaultValue,
              });
          }
        }
        
        return { content: [{ type: 'text', text: JSON.stringify(list, null, 2) }] };
      }
      case 'update': {
        await this.graphClient
          .api(`/sites/${args.siteId}/lists/${args.listId}`)
          .patch({
            displayName: args.title,
            description: args.description,
          });
        
        return { content: [{ type: 'text', text: 'SharePoint list updated successfully' }] };
      }
      case 'delete': {
        await this.graphClient
          .api(`/sites/${args.siteId}/lists/${args.listId}`)
          .delete();
        
        return { content: [{ type: 'text', text: 'SharePoint list deleted successfully' }] };
      }
      case 'add_items': {
        if (!args.items?.length) {
          throw new McpError(ErrorCode.InvalidParams, 'No items specified to add');
        }
        
        const results = [];
        for (const item of args.items) {
          const result = await this.graphClient
            .api(`/sites/${args.siteId}/lists/${args.listId}/items`)
            .post({
              fields: item,
            });
          
          results.push(result);
        }
        
        return { content: [{ type: 'text', text: JSON.stringify(results, null, 2) }] };
      }
      case 'get_items': {
        const items = await this.graphClient
          .api(`/sites/${args.siteId}/lists/${args.listId}/items?expand=fields`)
          .get();
        
        return { content: [{ type: 'text', text: JSON.stringify(items, null, 2) }] };
      }
      default:
        throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
    }
  }

  async run(): Promise<void> {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('Microsoft 365 Core MCP server running on stdio');
  }
}

const server = new M365CoreServer();
server.run().catch(console.error);
