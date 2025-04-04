#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
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
  AzureAdRoleArgs,
  AzureAdAppArgs,
  AzureAdDeviceArgs,
  AzureAdSpArgs,
  CallMicrosoftApiArgs,
  AuditLogArgs, // Import new type
  AlertArgs, // Import new type
} from './types.js';
import {
  m365CoreTools,
  sharePointSiteSchema,
  sharePointListSchema,
  distributionListSchema,
  securityGroupSchema,
  m365GroupSchema,
  exchangeSettingsSchema,
  userManagementSchema,
  offboardingSchema,
  azureAdRoleSchema,
  azureAdAppSchema,
  azureAdDeviceSchema,
  azureAdSpSchema,
  callMicrosoftApiSchema,
  auditLogSchema, // Import new schema
  alertSchema, // Import new schema
} from './tool-definitions.js';
import { z } from 'zod'; // Keep Zod import for error handling

// Environment validation
const MS_TENANT_ID = process.env.MS_TENANT_ID ?? '';
const MS_CLIENT_ID = process.env.MS_CLIENT_ID ?? '';
const MS_CLIENT_SECRET = process.env.MS_CLIENT_SECRET ?? '';

if (!MS_TENANT_ID || !MS_CLIENT_ID || !MS_CLIENT_SECRET) {
  throw new Error('Required environment variables (MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET) are missing');
}

// Define API configurations
const apiConfigs = {
  graph: {
    scope: "https://graph.microsoft.com/.default",
    baseUrl: "https://graph.microsoft.com/v1.0",
  },
  azure: {
    scope: "https://management.azure.com/.default",
    baseUrl: "https://management.azure.com",
  }
};

class M365CoreServer {
  private server: Server;
  private graphClient: Client;
  // Cache for tokens based on scope
  private tokenCache: Map<string, { token: string; expiresOn: number }> = new Map();


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

    // Initialize Graph client with default scope
    this.graphClient = Client.init({
      authProvider: async (callback) => {
        try {
          const token = await this.getAccessToken(apiConfigs.graph.scope); // Use default Graph scope
          callback(null, token);
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

  // Modified to accept scope and use cache
  private async getAccessToken(scope: string = apiConfigs.graph.scope): Promise<string> {
    const cached = this.tokenCache.get(scope);
    const now = Date.now();

    // Return cached token if valid (expires in > 60 seconds)
    if (cached && cached.expiresOn > now + 60 * 1000) {
      return cached.token;
    }

    // Fetch new token
    const tokenEndpoint = `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`;
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
       const errorData = await response.text();
       console.error("Token acquisition error:", response.status, errorData);
      throw new Error(`Failed to get access token for scope ${scope}. Status: ${response.status} ${response.statusText}. Details: ${errorData}`);
    }

    const data = await response.json();
    if (!data.access_token || !data.expires_in) {
       console.error("Invalid token response:", data);
       throw new Error(`Invalid token response received for scope ${scope}`);
    }

    // Cache the new token with expiration time (expires_in is in seconds)
    const expiresOn = now + data.expires_in * 1000;
    this.tokenCache.set(scope, { token: data.access_token, expiresOn });

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

          case 'manage_azure_ad_roles': {
            try {
              const roleArgs = azureAdRoleSchema.parse(args);
              return await this.handleAzureAdRoles(roleArgs);
            } catch (error) {
              if (error instanceof z.ZodError) {
                throw new McpError(
                  ErrorCode.InvalidParams,
                  `Invalid Azure AD role parameters: ${error.errors.map(e => e.message).join(', ')}`
                );
              }
              throw error;
            }
          }

          case 'manage_azure_ad_apps': {
             try {
               const appArgs = azureAdAppSchema.parse(args);
               return await this.handleAzureAdApps(appArgs);
             } catch (error) {
               if (error instanceof z.ZodError) {
                 throw new McpError(
                   ErrorCode.InvalidParams,
                   `Invalid Azure AD app parameters: ${error.errors.map(e => e.message).join(', ')}`
                 );
               }
               throw error;
             }
           }

           case 'manage_azure_ad_devices': {
             try {
               const deviceArgs = azureAdDeviceSchema.parse(args);
               return await this.handleAzureAdDevices(deviceArgs);
             } catch (error) {
               if (error instanceof z.ZodError) {
                 throw new McpError(
                   ErrorCode.InvalidParams,
                   `Invalid Azure AD device parameters: ${error.errors.map(e => e.message).join(', ')}`
                 );
               }
               throw error;
             }
           }

           case 'manage_service_principals': {
             try {
               const spArgs = azureAdSpSchema.parse(args);
               return await this.handleServicePrincipals(spArgs);
             } catch (error) {
               if (error instanceof z.ZodError) {
                 throw new McpError(
                   ErrorCode.InvalidParams,
                   `Invalid Service Principal parameters: ${error.errors.map(e => e.message).join(', ')}`
                 );
               }
               throw error;
             }
           }

           case 'call_microsoft_api': {
             try {
               const apiArgs = callMicrosoftApiSchema.parse(args);
               return await this.handleCallMicrosoftApi(apiArgs);
             } catch (error) {
               if (error instanceof z.ZodError) {
                 throw new McpError(
                   ErrorCode.InvalidParams,
                   `Invalid API call parameters: ${error.errors.map(e => e.message).join(', ')}`
                 );
               }
               throw error;
             }
           }

           case 'search_audit_log': {
             try {
               const auditArgs = auditLogSchema.parse(args);
               return await this.handleSearchAuditLog(auditArgs);
             } catch (error) {
               if (error instanceof z.ZodError) {
                 throw new McpError(
                   ErrorCode.InvalidParams,
                   `Invalid Audit Log parameters: ${error.errors.map(e => e.message).join(', ')}`
                 );
               }
               throw error;
             }
           }

           case 'manage_alerts': {
             try {
               const alertArgs = alertSchema.parse(args);
               return await this.handleManageAlerts(alertArgs);
             } catch (error) {
               if (error instanceof z.ZodError) {
                 throw new McpError(
                   ErrorCode.InvalidParams,
                   `Invalid Alert parameters: ${error.errors.map(e => e.message).join(', ')}`
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

  // --- Tool Handlers ---

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

  // --- New Azure AD Role Handler ---
  private async handleAzureAdRoles(args: AzureAdRoleArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    let apiPath = '';
    let result: any;

    switch (args.action) {
      case 'list_roles':
        apiPath = '/directoryRoles';
        if (args.filter) {
          apiPath += `?$filter=${encodeURIComponent(args.filter)}`;
        }
        result = await this.graphClient.api(apiPath).get();
        break;

      case 'list_role_assignments':
        // Note: Listing all role assignments requires Directory.Read.All
        // Filtering by principal requires RoleManagement.Read.Directory
        apiPath = '/roleManagement/directory/roleAssignments';
        if (args.filter) {
          // Example filter: $filter=principalId eq '{principalId}'
          apiPath += `?$filter=${encodeURIComponent(args.filter)}`;
        }
        result = await this.graphClient.api(apiPath).get();
        break;

      case 'assign_role':
        if (!args.roleId || !args.principalId) {
          throw new McpError(ErrorCode.InvalidParams, 'roleId and principalId are required for assign_role');
        }
        apiPath = '/roleManagement/directory/roleAssignments';
        const assignmentPayload = {
          '@odata.type': '#microsoft.graph.unifiedRoleAssignment',
          roleDefinitionId: args.roleId,
          principalId: args.principalId,
          directoryScopeId: '/', // Assign at tenant scope
        };
        result = await this.graphClient.api(apiPath).post(assignmentPayload);
        break;

      case 'remove_role_assignment':
        if (!args.assignmentId) {
          throw new McpError(ErrorCode.InvalidParams, 'assignmentId is required for remove_role_assignment');
        }
        apiPath = `/roleManagement/directory/roleAssignments/${args.assignmentId}`;
        await this.graphClient.api(apiPath).delete();
        result = { message: 'Role assignment removed successfully' };
        break;

      default:
        throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
    }

    return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
  }

  // --- New Azure AD App Handler ---
  private async handleAzureAdApps(args: AzureAdAppArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    let apiPath = '';
    let result: any;

    switch (args.action) {
      case 'list_apps':
        apiPath = '/applications';
        if (args.filter) {
          apiPath += `?$filter=${encodeURIComponent(args.filter)}`;
        }
        result = await this.graphClient.api(apiPath).get();
        break;

      case 'get_app':
        if (!args.appId) {
          throw new McpError(ErrorCode.InvalidParams, 'appId is required for get_app');
        }
        apiPath = `/applications/${args.appId}`;
        result = await this.graphClient.api(apiPath).get();
        break;

      case 'update_app':
        if (!args.appId || !args.appDetails) {
          throw new McpError(ErrorCode.InvalidParams, 'appId and appDetails are required for update_app');
        }
        apiPath = `/applications/${args.appId}`;
        await this.graphClient.api(apiPath).patch(args.appDetails);
        result = { message: 'Application updated successfully' };
        break;

      case 'add_owner':
        if (!args.appId || !args.ownerId) {
          throw new McpError(ErrorCode.InvalidParams, 'appId and ownerId are required for add_owner');
        }
        apiPath = `/applications/${args.appId}/owners/$ref`;
        const ownerPayload = {
          '@odata.id': `https://graph.microsoft.com/v1.0/users/${args.ownerId}`
        };
        await this.graphClient.api(apiPath).post(ownerPayload);
        result = { message: 'Owner added successfully' };
        break;

      case 'remove_owner':
        if (!args.appId || !args.ownerId) {
          throw new McpError(ErrorCode.InvalidParams, 'appId and ownerId are required for remove_owner');
        }
        // Need to get the specific owner reference ID first, as Graph requires the owner's directoryObject ID from the owners collection
        // This is a simplification; a real implementation might need to list owners first to find the correct reference ID.
        // For now, we'll assume ownerId is the directoryObject ID of the owner within the app's owners collection.
        apiPath = `/applications/${args.appId}/owners/${args.ownerId}/$ref`;
        await this.graphClient.api(apiPath).delete();
        result = { message: 'Owner removed successfully' };
        break;

      default:
        throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
    }

    return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
  }

  // --- New Azure AD Device Handler ---
  private async handleAzureAdDevices(args: AzureAdDeviceArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    let apiPath = '';
    let result: any;

    switch (args.action) {
      case 'list_devices':
        apiPath = '/devices';
        if (args.filter) {
          apiPath += `?$filter=${encodeURIComponent(args.filter)}`;
        }
        result = await this.graphClient.api(apiPath).get();
        break;

      case 'get_device':
        if (!args.deviceId) {
          throw new McpError(ErrorCode.InvalidParams, 'deviceId is required for get_device');
        }
        apiPath = `/devices/${args.deviceId}`;
        result = await this.graphClient.api(apiPath).get();
        break;

      case 'enable_device':
      case 'disable_device':
        if (!args.deviceId) {
          throw new McpError(ErrorCode.InvalidParams, `deviceId is required for ${args.action}`);
        }
        // Note: Enabling/Disabling devices is done via update, setting accountEnabled
        // This requires Device.ReadWrite.All permission.
        apiPath = `/devices/${args.deviceId}`;
        await this.graphClient.api(apiPath).patch({
          accountEnabled: args.action === 'enable_device'
        });
        result = { message: `Device ${args.action === 'enable_device' ? 'enabled' : 'disabled'} successfully` };
        break;

      case 'delete_device':
        if (!args.deviceId) {
          throw new McpError(ErrorCode.InvalidParams, 'deviceId is required for delete_device');
        }
        // Requires Device.ReadWrite.All permission.
        apiPath = `/devices/${args.deviceId}`;
        await this.graphClient.api(apiPath).delete();
        result = { message: 'Device deleted successfully' };
        break;

      default:
        throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
    }

    return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
  }

  // --- New Service Principal Handler ---
  private async handleServicePrincipals(args: AzureAdSpArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    let apiPath = '';
    let result: any;

    switch (args.action) {
      case 'list_sps':
        apiPath = '/servicePrincipals';
        if (args.filter) {
          apiPath += `?$filter=${encodeURIComponent(args.filter)}`;
        }
        result = await this.graphClient.api(apiPath).get();
        break;

      case 'get_sp':
        if (!args.spId) {
          throw new McpError(ErrorCode.InvalidParams, 'spId is required for get_sp');
        }
        apiPath = `/servicePrincipals/${args.spId}`;
        result = await this.graphClient.api(apiPath).get();
        break;

      case 'add_owner':
        if (!args.spId || !args.ownerId) {
          throw new McpError(ErrorCode.InvalidParams, 'spId and ownerId are required for add_owner');
        }
        // Requires Application.ReadWrite.All or Directory.ReadWrite.All
        apiPath = `/servicePrincipals/${args.spId}/owners/$ref`;
        const ownerPayload = {
          '@odata.id': `https://graph.microsoft.com/v1.0/users/${args.ownerId}`
        };
        await this.graphClient.api(apiPath).post(ownerPayload);
        result = { message: 'Owner added successfully to Service Principal' };
        break;

      case 'remove_owner':
        if (!args.spId || !args.ownerId) {
          throw new McpError(ErrorCode.InvalidParams, 'spId and ownerId are required for remove_owner');
        }
        // Requires Application.ReadWrite.All or Directory.ReadWrite.All
        // Similar to app owners, requires the directoryObject ID of the owner relationship
        apiPath = `/servicePrincipals/${args.spId}/owners/${args.ownerId}/$ref`;
        await this.graphClient.api(apiPath).delete();
        result = { message: 'Owner removed successfully from Service Principal' };
        break;

      default:
        throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
    }

    return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
  }

  // --- Generic API Call Handler ---
  private async handleCallMicrosoftApi(args: CallMicrosoftApiArgs): Promise<{ content: { type: string; text: string; }[]; isError?: boolean }> {
    try {
      const { apiType, path, method, apiVersion, subscriptionId, queryParams, body } = args;

      if (apiType === 'azure' && !apiVersion) {
        throw new McpError(ErrorCode.InvalidParams, "apiVersion is required for apiType 'azure'");
      }

      const config = apiConfigs[apiType];
      const token = await this.getAccessToken(config.scope);

      let url = config.baseUrl;
      const urlParams = new URLSearchParams();

      // Construct URL based on API type
      if (apiType === 'azure') {
        let azurePath = path;
        // Prepend subscription if provided and path doesn't already include it
        if (subscriptionId && !path.toLowerCase().startsWith('/subscriptions/')) {
           azurePath = `/subscriptions/${subscriptionId}${path.startsWith('/') ? '' : '/'}${path}`;
        }
        url += azurePath;
        urlParams.append('api-version', apiVersion!); // Already validated
      } else { // graph
        url += path.startsWith('/') ? path : `/${path}`;
      }

      // Add query parameters
      if (queryParams) {
        for (const [key, value] of Object.entries(queryParams)) {
          urlParams.append(key, value);
        }
      }

      const queryString = urlParams.toString();
      if (queryString) {
        url += `?${queryString}`;
      }

      // Prepare request options
      const headers: Record<string, string> = {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      };
      if (apiType === 'graph') {
        headers['ConsistencyLevel'] = 'eventual'; // Good practice for Graph count/filter
      }

      const requestOptions: RequestInit = {
        method: method.toUpperCase(),
        headers: headers
      };

      if (["POST", "PUT", "PATCH"].includes(method.toUpperCase()) && body !== undefined) {
        requestOptions.body = typeof body === 'string' ? body : JSON.stringify(body);
      }

      // Make API request
      const apiResponse = await fetch(url, requestOptions);
      const responseText = await apiResponse.text();
      let responseData: any;

      try {
        responseData = responseText ? JSON.parse(responseText) : {};
      } catch (e) {
        // If not JSON, return raw text
        responseData = { rawResponse: responseText };
      }

      if (!apiResponse.ok) {
        console.error(`API Error (${apiResponse.status}) for ${method.toUpperCase()} ${url}:`, responseData);
        // Return error details in a structured way
         return {
           content: [{ type: 'text', text: JSON.stringify({
               error: `API Error (${apiResponse.status} ${apiResponse.statusText})`,
               url: url,
               details: responseData
             }, null, 2)
           }],
           isError: true
         };
      }

      // Successful response
      return {
        content: [{ type: 'text', text: JSON.stringify(responseData, null, 2) }]
      };

    } catch (error) {
       console.error("Error in handleCallMicrosoftApi:", error);
       const errorMessage = error instanceof Error ? error.message : String(error);
       // Ensure McpError is thrown correctly if it's already one
       if (error instanceof McpError) {
          throw error;
       }
       // Wrap other errors
       throw new McpError(ErrorCode.InternalError, `Failed to execute API call: ${errorMessage}`);
    }
  }

  // --- Security & Compliance Handlers ---

  private async handleSearchAuditLog(args: AuditLogArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    // Primarily targets /auditLogs/directoryAudits for now
    // Requires AuditLog.Read.All permission
    let apiPath = '/auditLogs/directoryAudits';
    const queryOptions: string[] = [];

    if (args.filter) {
      queryOptions.push(`$filter=${encodeURIComponent(args.filter)}`);
    }
    if (args.top) {
      queryOptions.push(`$top=${args.top}`);
    }

    if (queryOptions.length > 0) {
      apiPath += `?${queryOptions.join('&')}`;
    }

    const result = await this.graphClient.api(apiPath).get();
    return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
  }

  private async handleManageAlerts(args: AlertArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    // Uses the newer alerts_v2 endpoint
    // Requires SecurityAlert.Read.All permission
    let apiPath = '/security/alerts_v2';
    let result: any;

    switch (args.action) {
      case 'list_alerts': {
        const queryOptions: string[] = [];
        if (args.filter) {
          queryOptions.push(`$filter=${encodeURIComponent(args.filter)}`);
        }
        if (args.top) {
          queryOptions.push(`$top=${args.top}`);
        }
        if (queryOptions.length > 0) {
          apiPath += `?${queryOptions.join('&')}`;
        }
        result = await this.graphClient.api(apiPath).get();
        break;
      }
      case 'get_alert': {
        if (!args.alertId) {
          throw new McpError(ErrorCode.InvalidParams, 'alertId is required for get_alert');
        }
        apiPath += `/${args.alertId}`;
        result = await this.graphClient.api(apiPath).get();
        break;
      }
      default:
        throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
    }

    return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
  }


  async run(): Promise<void> {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('Microsoft 365 Core MCP server running on stdio');
  }
}

const server = new M365CoreServer();
server.run().catch(console.error);
