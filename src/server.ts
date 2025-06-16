#!/usr/bin/env node
import { McpServer, ResourceTemplate } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { wrapToolHandler, formatTextResponse } from './utils.js';
import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';
import express from 'express';
import { randomUUID } from 'crypto';
import { z } from 'zod';
import { zodToJsonSchema } from 'zod-to-json-schema';

// Import Azure Identity library (commented out due to package installation issues)
// import { ClientSecretCredential } from '@azure/identity';
// import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';

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
  AuditLogArgs,
  AlertArgs,
} from './types.js';

import {
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
  auditLogSchema,
  alertSchema,
  dlpPolicySchema,
  dlpIncidentSchema,
  sensitivityLabelSchema,
  intuneMacOSDeviceSchema,
  intuneMacOSPolicySchema,
  intuneMacOSAppSchema,
  intuneMacOSComplianceSchema,
  complianceFrameworkSchema,
  complianceAssessmentSchema,
  complianceMonitoringSchema,
  evidenceCollectionSchema,
  gapAnalysisSchema,
  auditReportSchema,
  cisComplianceSchema,
  m365CoreTools,
} from './tool-definitions.js';

import {
  handleUserSettings,
  handleOffboarding,
  handleSharePointSite,
  handleSharePointList,
  handleAzureAdRoles,
  handleAzureAdApps,
  handleAzureAdDevices,
  handleServicePrincipals,
  handleCallMicrosoftApi,
  handleSearchAuditLog,
  handleManageAlerts,
} from './handlers.js';

// Import DLP handlers and types
import {
  handleDLPPolicies,
  handleDLPIncidents,
  handleDLPSensitivityLabels
} from './handlers/dlp-handler.js';
import {
  DLPPolicyArgs,
  DLPIncidentArgs,
  DLPSensitivityLabelArgs
} from './types/dlp-types.js';

// Import Intune macOS handlers and types
import {
  handleIntuneMacOSDevices,
  handleIntuneMacOSPolicies,
  handleIntuneMacOSApps,
  handleIntuneMacOSCompliance
} from './handlers/intune-macos-handler.js';
import {
  IntuneMacOSDeviceArgs,
  IntuneMacOSPolicyArgs,
  IntuneMacOSAppArgs,
  IntuneMacOSComplianceArgs
} from './types/intune-types.js';

// Import compliance handlers and types
import {
  handleComplianceFrameworks,
  handleComplianceAssessments,
  handleComplianceMonitoring,
  handleEvidenceCollection,
  handleGapAnalysis,
  handleCISCompliance
} from './handlers/compliance-handler.js';
import {
  ComplianceFrameworkArgs,
  ComplianceAssessmentArgs,
  ComplianceMonitoringArgs,
  EvidenceCollectionArgs,
  GapAnalysisArgs,
  CISComplianceArgs
} from './types/compliance-types.js';

// Import audit reporting handler
import { handleAuditReports } from './handlers/audit-reporting-handler.js';
import { AuditReportArgs } from './types/compliance-types.js';

import { handleExchangeSettings } from './exchange-handler.js';
import { setupExtendedResources, setupExtendedPrompts } from './extended-resources.js';

// Environment validation
const MS_TENANT_ID = process.env.MS_TENANT_ID ?? '';
const MS_CLIENT_ID = process.env.MS_CLIENT_ID ?? '';
const MS_CLIENT_SECRET = process.env.MS_CLIENT_SECRET ?? '';
const PORT = process.env.PORT ? parseInt(process.env.PORT, 10) : 3000;
const LOG_LEVEL = process.env.LOG_LEVEL ?? 'info';

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

export class M365CoreServer {
  public server: McpServer;
  private graphClient: Client;
  // Cache for tokens based on scope
  private tokenCache: Map<string, { token: string; expiresOn: number }> = new Map();

  constructor() {
    this.server = new McpServer({
      name: 'm365-core-server',
      version: '1.0.0',
    });

    // Current authentication method
    this.graphClient = Client.init({
      authProvider: async (callback: (error: Error | null, token: string | null) => void) => {
        try {
          const token = await this.getAccessToken(apiConfigs.graph.scope);
          callback(null, token);
        } catch (error) {
          callback(error as Error, null);
        }
      }
    });

    // Azure Identity authentication method (commented out due to package installation issues)
    // Uncomment this code and comment out the above authentication method once you've installed the required packages
    /*
    // Initialize Azure Credential
    const azureCredential = new ClientSecretCredential(MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET);

    // Initialize Graph Authentication Provider
    const authProvider = new TokenCredentialAuthenticationProvider(azureCredential, {
      scopes: ["https://graph.microsoft.com/.default"],
    });

    // Initialize Graph Client
    this.graphClient = Client.initWithMiddleware({
      authProvider: authProvider,
    });
    */    this.setupTools();
    this.setupResources();
    
    // Setup extended resources and prompts
    setupExtendedResources(this.server, this.graphClient);
    setupExtendedPrompts(this.server, this.graphClient);
    
    // Error handling
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
    const expiresOn = now + data.expires_in * 1000;    this.tokenCache.set(scope, { token: data.access_token, expiresOn });

    return data.access_token;
  }

  private validateCredentials(): void {
    const requiredEnvVars = ['MS_TENANT_ID', 'MS_CLIENT_ID', 'MS_CLIENT_SECRET'];
    const missingVars = requiredEnvVars.filter(varName => !process.env[varName]);
    
    if (missingVars.length > 0) {
      throw new McpError(
        ErrorCode.InvalidRequest,
        `Missing required environment variables for Microsoft 365 authentication: ${missingVars.join(', ')}. ` +
        `Please configure these variables:\n` +
        `- MS_TENANT_ID: Your Azure AD tenant ID\n` +
        `- MS_CLIENT_ID: Your Azure AD application (client) ID\n` +
        `- MS_CLIENT_SECRET: Your Azure AD application client secret\n\n` +
        `For setup instructions, visit: https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app`
      );
    }
  }
  private setupTools(): void {
    // Distribution Lists - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_distribution_lists",
      distributionListSchema,
      wrapToolHandler(async (args: DistributionListArgs) => {
        // Validate credentials only when tool is executed (lazy loading)
        this.validateCredentials();
        try {
          return await this.handleDistributionList(args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing tool: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Security Groups - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_security_groups",
      securityGroupSchema,
      wrapToolHandler(async (args: SecurityGroupArgs) => {
        // Validate credentials only when tool is executed (lazy loading)
        this.validateCredentials();
        try {
          return await this.handleSecurityGroup(args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing tool: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // M365 Groups
    this.server.tool(
      "manage_m365_groups",
      m365GroupSchema,
      wrapToolHandler(async (args: M365GroupArgs) => {
        try {
          return await this.handleM365Group(args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing tool: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Exchange Settings
    this.server.tool(
      "manage_exchange_settings",
      exchangeSettingsSchema,
      wrapToolHandler(async (args: ExchangeSettingsArgs) => {
        try {
          return await handleExchangeSettings(this.graphClient, args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing tool: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // User Management
    this.server.tool(
      "manage_user_settings",
      userManagementSchema,
      wrapToolHandler(async (args: UserManagementArgs) => {
        try {
          return await handleUserSettings(this.graphClient, args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing tool: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Offboarding
    this.server.tool(
      "manage_offboarding",
      offboardingSchema,
      wrapToolHandler(async (args: OffboardingArgs) => {
        try {
          return await handleOffboarding(this.graphClient, args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing tool: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // SharePoint Sites
    this.server.tool(
      "manage_sharepoint_sites",
      sharePointSiteSchema,
      wrapToolHandler(async (args: SharePointSiteArgs) => {
        try {
          return await handleSharePointSite(this.graphClient, args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing tool: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // SharePoint Lists
    this.server.tool(
      "manage_sharepoint_lists",
      sharePointListSchema,
      wrapToolHandler(async (args: SharePointListArgs) => {
        try {
          return await handleSharePointList(this.graphClient, args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing tool: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Azure AD Roles
    this.server.tool(
      "manage_azure_ad_roles",
      azureAdRoleSchema,
      wrapToolHandler(async (args: AzureAdRoleArgs) => {
        try {
          return await handleAzureAdRoles(this.graphClient, args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing tool: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Azure AD Apps
    this.server.tool(
      "manage_azure_ad_apps",
      azureAdAppSchema,
      wrapToolHandler(async (args: AzureAdAppArgs) => {
        try {
          return await handleAzureAdApps(this.graphClient, args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing tool: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Azure AD Devices
    this.server.tool(
      "manage_azure_ad_devices",
      azureAdDeviceSchema,
      wrapToolHandler(async (args: AzureAdDeviceArgs) => {
        try {
          return await handleAzureAdDevices(this.graphClient, args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing tool: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Service Principals
    this.server.tool(
      "manage_service_principals",
      azureAdSpSchema,
      wrapToolHandler(async (args: AzureAdSpArgs) => {
        try {
          return await handleServicePrincipals(this.graphClient, args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing tool: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Dynamic API Endpoint
    this.server.tool(
      "call_microsoft_api",
      callMicrosoftApiSchema,
      wrapToolHandler(async (args: CallMicrosoftApiArgs) => {
        try {
          return await handleCallMicrosoftApi(this.graphClient, args, this.getAccessToken.bind(this), apiConfigs);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing tool: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Audit Log
    this.server.tool(
      "search_audit_log",
      auditLogSchema,
      wrapToolHandler(async (args: AuditLogArgs) => {
        try {
          return await handleSearchAuditLog(this.graphClient, args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing tool: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Alerts
    this.server.tool(
      "manage_alerts",
      alertSchema,
      wrapToolHandler(async (args: AlertArgs) => {
        try {
          return await handleManageAlerts(this.graphClient, args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing tool: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );
  }

  private setupResources(): void {
    // Static resources
    this.server.resource(
      "current_user",
      "m365://users/current",
      async (uri: URL) => {
        try {
          const currentUser = await this.graphClient
            .api('/me')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(currentUser, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "tenant_info",
      "m365://tenant/info",
      async (uri: URL) => {
        try {
          const tenantInfo = await this.graphClient
            .api('/organization')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(tenantInfo, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "sharepoint_sites",
      "m365://sharepoint/sites",
      async (uri: URL) => {
        try {
          const sites = await this.graphClient
            .api('/sites?search=*')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(sites, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "sharepoint_admin_settings",
      "m365://sharepoint/admin/settings",
      async (uri: URL) => {
        try {
          const settings = await this.graphClient
            .api('/admin/sharepoint/settings')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(settings, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );

    // Dynamic resources with templates
    this.server.resource(
      "user_info",
      new ResourceTemplate("m365://users/{userId}", { list: undefined }),
      async (uri: URL, variables) => {
        try {
          const user = await this.graphClient
            .api(`/users/${variables.userId}`)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(user, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "group_info",
      new ResourceTemplate("m365://groups/{groupId}", { list: undefined }),
      async (uri: URL, variables) => {
        try {
          const group = await this.graphClient
            .api(`/groups/${variables.groupId}`)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(group, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "sharepoint_site_info",
      new ResourceTemplate("m365://sharepoint/sites/{siteId}", { list: undefined }),
      async (uri: URL, variables) => {
        try {
          const site = await this.graphClient
            .api(`/sites/${variables.siteId}`)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(site, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "sharepoint_lists",
      new ResourceTemplate("m365://sharepoint/sites/{siteId}/lists", { list: undefined }),
      async (uri: URL, variables) => {
        try {
          const lists = await this.graphClient
            .api(`/sites/${variables.siteId}/lists`)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(lists, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "sharepoint_list_info",
      new ResourceTemplate("m365://sharepoint/sites/{siteId}/lists/{listId}", { list: undefined }),
      async (uri: URL, variables) => {
        try {
          const list = await this.graphClient
            .api(`/sites/${variables.siteId}/lists/${variables.listId}`)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(list, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "sharepoint_list_items",
      new ResourceTemplate("m365://sharepoint/sites/{siteId}/lists/{listId}/items", { list: undefined }),
      async (uri: URL, variables) => {
        try {
          const items = await this.graphClient
            .api(`/sites/${variables.siteId}/lists/${variables.listId}/items?expand=fields`)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(items, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    // Additional M365 Resources (Security, Compliance, Intune, etc.)
    
    this.server.resource(
      "security_alerts",
      "m365://security/alerts",
      async (uri: URL) => {
        try {
          const alerts = await this.graphClient
            .api('/security/alerts_v2')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(alerts, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "security_incidents",
      "m365://security/incidents",
      async (uri: URL) => {
        try {
          const incidents = await this.graphClient
            .api('/security/incidents')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(incidents, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "conditional_access_policies",
      "m365://identity/conditionalAccess/policies",
      async (uri: URL) => {
        try {
          const policies = await this.graphClient
            .api('/identity/conditionalAccess/policies')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(policies, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "applications",
      "m365://applications",
      async (uri: URL) => {
        try {
          const applications = await this.graphClient
            .api('/applications')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(applications, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "service_principals",
      "m365://servicePrincipals",
      async (uri: URL) => {
        try {
          const servicePrincipals = await this.graphClient
            .api('/servicePrincipals')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(servicePrincipals, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "directory_roles",
      "m365://directoryRoles",
      async (uri: URL) => {
        try {
          const roles = await this.graphClient
            .api('/directoryRoles')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(roles, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "privileged_access",
      "m365://privilegedAccess/azureAD/resources",
      async (uri: URL) => {
        try {
          const resources = await this.graphClient
            .api('/privilegedAccess/azureAD/resources')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(resources, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "audit_logs_signin",
      "m365://auditLogs/signIns",
      async (uri: URL) => {
        try {
          const signIns = await this.graphClient
            .api('/auditLogs/signIns')
            .top(50)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(signIns, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "audit_logs_directory",
      "m365://auditLogs/directoryAudits",
      async (uri: URL) => {
        try {
          const directoryAudits = await this.graphClient
            .api('/auditLogs/directoryAudits')
            .top(50)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(directoryAudits, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "intune_devices",
      "m365://deviceManagement/managedDevices",
      async (uri: URL) => {
        try {
          const devices = await this.graphClient
            .api('/deviceManagement/managedDevices')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(devices, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "intune_apps",
      "m365://deviceAppManagement/mobileApps",
      async (uri: URL) => {
        try {
          const apps = await this.graphClient
            .api('/deviceAppManagement/mobileApps')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(apps, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "intune_compliance_policies",
      "m365://deviceManagement/deviceCompliancePolicies",
      async (uri: URL) => {
        try {
          const policies = await this.graphClient
            .api('/deviceManagement/deviceCompliancePolicies')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(policies, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "intune_configuration_policies",
      "m365://deviceManagement/deviceConfigurations",
      async (uri: URL) => {
        try {
          const configurations = await this.graphClient
            .api('/deviceManagement/deviceConfigurations')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(configurations, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "teams_list",
      "m365://teams",
      async (uri: URL) => {
        try {
          const teams = await this.graphClient
            .api('/teams')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(teams, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "mail_folders",
      "m365://me/mailFolders",
      async (uri: URL) => {
        try {
          const mailFolders = await this.graphClient
            .api('/me/mailFolders')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(mailFolders, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "calendar_events",
      "m365://me/events",
      async (uri: URL) => {
        try {
          const events = await this.graphClient
            .api('/me/events')
            .top(25)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(events, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "onedrive",
      "m365://me/drive",
      async (uri: URL) => {
        try {
          const drive = await this.graphClient
            .api('/me/drive')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(drive, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "planner_plans",
      "m365://planner/plans",
      async (uri: URL) => {
        try {
          const plans = await this.graphClient
            .api('/planner/plans')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(plans, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "information_protection",
      "m365://informationProtection/policy/labels",
      async (uri: URL) => {
        try {
          const labels = await this.graphClient
            .api('/informationProtection/policy/labels')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(labels, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "risky_users",
      "m365://identityProtection/riskyUsers",
      async (uri: URL) => {
        try {
          const riskyUsers = await this.graphClient
            .api('/identityProtection/riskyUsers')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(riskyUsers, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "threat_assessment",
      "m365://informationProtection/threatAssessmentRequests",
      async (uri: URL) => {
        try {
          const requests = await this.graphClient
            .api('/informationProtection/threatAssessmentRequests')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(requests, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    // Dynamic resources with parameters
    
    this.server.resource(
      "user_messages",
      new ResourceTemplate("m365://users/{userId}/messages", { list: undefined }),
      async (uri: URL, variables) => {
        try {
          const messages = await this.graphClient
            .api(`/users/${variables.userId}/messages`)
            .top(25)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(messages, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "user_calendar",
      new ResourceTemplate("m365://users/{userId}/calendar", { list: undefined }),
      async (uri: URL, variables) => {
        try {
          const calendar = await this.graphClient
            .api(`/users/${variables.userId}/calendar`)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(calendar, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "user_drive",
      new ResourceTemplate("m365://users/{userId}/drive", { list: undefined }),
      async (uri: URL, variables) => {
        try {
          const drive = await this.graphClient
            .api(`/users/${variables.userId}/drive`)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(drive, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "team_channels",
      new ResourceTemplate("m365://teams/{teamId}/channels", { list: undefined }),
      async (uri: URL, variables) => {
        try {
          const channels = await this.graphClient
            .api(`/teams/${variables.teamId}/channels`)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(channels, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "team_members",
      new ResourceTemplate("m365://teams/{teamId}/members", { list: undefined }),
      async (uri: URL, variables) => {
        try {
          const members = await this.graphClient
            .api(`/teams/${variables.teamId}/members`)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(members, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "device_info",
      new ResourceTemplate("m365://deviceManagement/managedDevices/{deviceId}", { list: undefined }),
      async (uri: URL, variables) => {
        try {
          const device = await this.graphClient
            .api(`/deviceManagement/managedDevices/${variables.deviceId}`)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(device, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "app_assignments",
      new ResourceTemplate("m365://deviceAppManagement/mobileApps/{appId}/assignments", { list: undefined }),
      async (uri: URL, variables) => {
        try {
          const assignments = await this.graphClient
            .api(`/deviceAppManagement/mobileApps/${variables.appId}/assignments`)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(assignments, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "policy_assignments",
      new ResourceTemplate("m365://deviceManagement/deviceCompliancePolicies/{policyId}/assignments", { list: undefined }),
      async (uri: URL, variables) => {
        try {
          const assignments = await this.graphClient
            .api(`/deviceManagement/deviceCompliancePolicies/${variables.policyId}/assignments`)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(assignments, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "group_members",
      new ResourceTemplate("m365://groups/{groupId}/members", { list: undefined }),
      async (uri: URL, variables) => {
        try {
          const members = await this.graphClient
            .api(`/groups/${variables.groupId}/members`)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(members, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "group_owners",
      new ResourceTemplate("m365://groups/{groupId}/owners", { list: undefined }),
      async (uri: URL, variables) => {
        try {
          const owners = await this.graphClient
            .api(`/groups/${variables.groupId}/owners`)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(owners, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "user_licenses",
      new ResourceTemplate("m365://users/{userId}/licenseDetails", { list: undefined }),
      async (uri: URL, variables) => {
        try {
          const licenses = await this.graphClient
            .api(`/users/${variables.userId}/licenseDetails`)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(licenses, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "user_groups",
      new ResourceTemplate("m365://users/{userId}/memberOf", { list: undefined }),
      async (uri: URL, variables) => {
        try {
          const groups = await this.graphClient
            .api(`/users/${variables.userId}/memberOf`)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(groups, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "security_score",
      "m365://security/secureScores",
      async (uri: URL) => {
        try {
          const secureScores = await this.graphClient
            .api('/security/secureScores')
            .top(10)
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(secureScores, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "compliance_policies_dlp",
      "m365://security/informationProtection/dlpPolicies",
      async (uri: URL) => {
        try {
          const dlpPolicies = await this.graphClient
            .api('/security/informationProtection/dlpPolicies')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(dlpPolicies, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "retention_policies",
      "m365://security/labels/retentionLabels",
      async (uri: URL) => {
        try {
          const retentionLabels = await this.graphClient
            .api('/security/labels/retentionLabels')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(retentionLabels, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "sensitivity_labels",
      "m365://security/informationProtection/sensitivityLabels",
      async (uri: URL) => {
        try {
          const sensitivityLabels = await this.graphClient
            .api('/security/informationProtection/sensitivityLabels')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(sensitivityLabels, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "communication_compliance",
      "m365://compliance/communicationCompliance/policies",
      async (uri: URL) => {
        try {
          const policies = await this.graphClient
            .api('/compliance/communicationCompliance/policies')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(policies, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "ediscovery_cases",
      "m365://compliance/ediscovery/cases",
      async (uri: URL) => {
        try {
          const cases = await this.graphClient
            .api('/compliance/ediscovery/cases')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(cases, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );
    
    this.server.resource(
      "subscribed_skus",
      "m365://subscribedSkus",
      async (uri: URL) => {
        try {
          const skus = await this.graphClient
            .api('/subscribedSkus')
            .get();
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(skus, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading resource: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );  }

  // --- Tool Handlers ---

  private async handleDistributionList(args: DistributionListArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    // Handler logic for managing distribution lists
    // This is a placeholder implementation - replace with actual logic
    return {
      content: [
        {
          type: "text",
          text: `Handled Distribution List: ${JSON.stringify(args, null, 2)}`,
        },
      ],
    };
  }

  private async handleSecurityGroup(args: SecurityGroupArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    // Handler logic for managing security groups
    // This is a placeholder implementation - replace with actual logic
    return {
      content: [
        {
          type: "text",
          text: `Handled Security Group: ${JSON.stringify(args, null, 2)}`,
        },
      ],
    };
  }

  private async handleM365Group(args: M365GroupArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    // Handler logic for managing M365 groups
    // This is a placeholder implementation - replace with actual logic
    return {
      content: [
        {
          type: "text",
          text: `Handled M365 Group: ${JSON.stringify(args, null, 2)}`,
        },
      ],
    };
  }

  private async handleAzureAdRoles(args: AzureAdRoleArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    // Handler logic for managing Azure AD roles
    // This is a placeholder implementation - replace with actual logic
    return {
      content: [
        {
          type: "text",
          text: `Handled Azure AD Role: ${JSON.stringify(args, null, 2)}`,
        },
      ],
    };
  }

  private async handleAzureAdApps(args: AzureAdAppArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    // Handler logic for managing Azure AD apps
    // This is a placeholder implementation - replace with actual logic
    return {
      content: [
        {
          type: "text",
          text: `Handled Azure AD App: ${JSON.stringify(args, null, 2)}`,
        },
      ],
    };
  }

  private async handleAzureAdDevices(args: AzureAdDeviceArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    // Handler logic for managing Azure AD devices
    // This is a placeholder implementation - replace with actual logic
    return {
      content: [
        {
          type: "text",
          text: `Handled Azure AD Device: ${JSON.stringify(args, null, 2)}`,
        },
      ],
    };
  }
  private async handleServicePrincipals(args: AzureAdSpArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    // Handler logic for managing service principals
    // This is a placeholder implementation - replace with actual logic
    return {
      content: [
        {
          type: "text",
          text: `Handled Service Principal: ${JSON.stringify(args, null, 2)}`,
        },
      ],
    };
  }
}