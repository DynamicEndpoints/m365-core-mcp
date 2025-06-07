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
  handleGapAnalysis
} from './handlers/compliance-handler.js';
import {
  ComplianceFrameworkArgs,
  ComplianceAssessmentArgs,
  ComplianceMonitoringArgs,
  EvidenceCollectionArgs,
  GapAnalysisArgs
} from './types/compliance-types.js';

// Import audit reporting handler
import { handleAuditReports } from './handlers/audit-reporting-handler.js';
import { AuditReportArgs } from './types/compliance-types.js';

import { handleExchangeSettings } from './exchange-handler.js';

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
    */

    this.setupTools();
    this.setupResources();
    
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
    const expiresOn = now + data.expires_in * 1000;
    this.tokenCache.set(scope, { token: data.access_token, expiresOn });

    return data.access_token;
  }

  private setupTools(): void {
    // Distribution Lists
    this.server.tool(
      "manage_distribution_lists",
      distributionListSchema,
      wrapToolHandler(async (args: DistributionListArgs) => {
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

    // Security Groups
    this.server.tool(
      "manage_security_groups",
      securityGroupSchema,
      wrapToolHandler(async (args: SecurityGroupArgs) => {
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
      "Dynamicendpoint_automation_assistant",
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

    // DLP Management Tools
    this.server.tool(
      "manage_dlp_policies",
      dlpPolicySchema,
      wrapToolHandler(async (args: DLPPolicyArgs) => {
        try {
          return await handleDLPPolicies(this.graphClient, args);
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

    this.server.tool(
      "manage_dlp_incidents",
      dlpIncidentSchema,
      wrapToolHandler(async (args: DLPIncidentArgs) => {
        try {
          return await handleDLPIncidents(this.graphClient, args);
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

    this.server.tool(
      "manage_sensitivity_labels",
      sensitivityLabelSchema,
      wrapToolHandler(async (args: DLPSensitivityLabelArgs) => {
        try {
          return await handleDLPSensitivityLabels(this.graphClient, args);
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

    // Intune macOS Management Tools
    this.server.tool(
      "manage_intune_macos_devices",
      intuneMacOSDeviceSchema,
      wrapToolHandler(async (args: IntuneMacOSDeviceArgs) => {
        try {
          return await handleIntuneMacOSDevices(this.graphClient, args);
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

    this.server.tool(
      "manage_intune_macos_policies",
      intuneMacOSPolicySchema,
      wrapToolHandler(async (args: IntuneMacOSPolicyArgs) => {
        try {
          return await handleIntuneMacOSPolicies(this.graphClient, args);
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

    this.server.tool(
      "manage_intune_macos_apps",
      intuneMacOSAppSchema,
      wrapToolHandler(async (args: IntuneMacOSAppArgs) => {
        try {
          return await handleIntuneMacOSApps(this.graphClient, args);
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

    this.server.tool(
      "manage_intune_macos_compliance",
      intuneMacOSComplianceSchema,
      wrapToolHandler(async (args: IntuneMacOSComplianceArgs) => {
        try {
          return await handleIntuneMacOSCompliance(this.graphClient, args);
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

    // Compliance Framework Management Tools
    this.server.tool(
      "manage_compliance_frameworks",
      complianceFrameworkSchema,
      wrapToolHandler(async (args: ComplianceFrameworkArgs) => {
        try {
          return await handleComplianceFrameworks(this.graphClient, args);
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

    this.server.tool(
      "manage_compliance_assessments",
      complianceAssessmentSchema,
      wrapToolHandler(async (args: ComplianceAssessmentArgs) => {
        try {
          return await handleComplianceAssessments(this.graphClient, args);
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

    this.server.tool(
      "manage_compliance_monitoring",
      complianceMonitoringSchema,
      wrapToolHandler(async (args: ComplianceMonitoringArgs) => {
        try {
          return await handleComplianceMonitoring(this.graphClient, args);
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

    this.server.tool(
      "manage_evidence_collection",
      evidenceCollectionSchema,
      wrapToolHandler(async (args: EvidenceCollectionArgs) => {
        try {
          return await handleEvidenceCollection(this.graphClient, args);
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

    this.server.tool(
      "manage_gap_analysis",
      gapAnalysisSchema,
      wrapToolHandler(async (args: GapAnalysisArgs) => {
        try {
          return await handleGapAnalysis(this.graphClient, args);
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

    this.server.tool(
      "generate_audit_reports",
      auditReportSchema,
      wrapToolHandler(async (args: AuditReportArgs) => {
        try {
          return await handleAuditReports(this.graphClient, args);
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
}
