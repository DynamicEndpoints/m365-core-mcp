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

// Import new Graph API framework components
import { DynamicToolGenerator } from './utils/dynamic-tool-generator.js';
import { GraphAdvancedFeatures, batchRequestSchema, deltaQuerySchema, webhookSubscriptionSchema, searchQuerySchema } from './utils/graph-advanced-features.js';
import { GraphMetadataService, GraphScopeManager } from './utils/graph-metadata-service.js';

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
  intuneWindowsDeviceSchema,
  intuneWindowsPolicySchema,
  intuneWindowsAppSchema,
  intuneWindowsComplianceSchema,
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
  handleDistributionLists,
  handleSecurityGroups,
  handleM365Groups,
  handleSharePointSites,
  handleSharePointLists,
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

// Import Intune Windows handlers and types
import {
  handleIntuneWindowsDevices,
  handleIntuneWindowsPolicies,
  handleIntuneWindowsApps,
  handleIntuneWindowsCompliance
} from './handlers/intune-windows-handler.js';
import {
  IntuneWindowsDeviceArgs,
  IntuneWindowsPolicyArgs,
  IntuneWindowsAppArgs,
  IntuneWindowsComplianceArgs
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

// Import resources and prompts
import { m365Resources, getResourceByUri, listResources } from './resources.js';
import { allM365Prompts, getPromptByName, listPrompts } from './prompts.js';

// Environment validation - will be checked lazily when tools are executed
// These values will be set from HTTP request configuration in index.ts
const PORT = process.env.PORT ? parseInt(process.env.PORT, 10) : 3000;
const LOG_LEVEL = process.env.LOG_LEVEL ?? 'info';

// Helper function to get environment variables with fallback
function getEnvVar(name: string): string {
  return process.env[name] ?? '';
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
  private graphClient: Client | null = null; // Make this nullable and initialize lazily
  // Cache for tokens based on scope
  private tokenCache: Map<string, { token: string; expiresOn: number }> = new Map();  // SSE and real-time capabilities
  public sseClients: Set<any> = new Set();
  public progressTrackers: Map<string, any> = new Map();
  
  constructor() {
    this.server = new McpServer({
      name: 'm365-core-server',
      version: '1.0.0',
      capabilities: {
        tools: {},
        resources: {
          subscribe: true,
          listChanged: true
        },
        prompts: {},
        logging: {}
      }
    });

    // Register tools, resources, and prompts immediately (no network calls)
    this.setupTools();
    this.setupResources();
    this.setupPrompts();
    
    // Error handling
    process.on('SIGINT', async () => {
      await this.server.close();
      process.exit(0);
    });
  }
  // Lazy initialization of Graph client
  private getGraphClient(): Client {
    if (!this.graphClient) {
      this.validateCredentials(); // This only checks env vars, no network calls
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
    }
    return this.graphClient;
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
    const tokenEndpoint = `https://login.microsoftonline.com/${getEnvVar('MS_TENANT_ID')}/oauth2/v2.0/token`;
    const params = new URLSearchParams();
    params.append('client_id', getEnvVar('MS_CLIENT_ID'));
    params.append('client_secret', getEnvVar('MS_CLIENT_SECRET'));
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
    const missingVars = requiredEnvVars.filter(varName => !getEnvVar(varName));
    
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
  }  private setupTools(): void {    // Distribution Lists - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_distribution_lists",
      distributionListSchema.shape,
      wrapToolHandler(async (args: DistributionListArgs) => {
        // Validate credentials only when tool is executed (lazy loading)
        this.validateCredentials();        try {
          return await handleDistributionLists(this.getGraphClient(), args);
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
    );// Security Groups - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_security_groups",
      securityGroupSchema.shape,
      wrapToolHandler(async (args: SecurityGroupArgs) => {
        // Validate credentials only when tool is executed (lazy loading)
        this.validateCredentials();        try {
          return await handleSecurityGroups(this.getGraphClient(), args);
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
    );    // M365 Groups - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_m365_groups",
      m365GroupSchema.shape,
      wrapToolHandler(async (args: M365GroupArgs) => {
        // Validate credentials only when tool is executed (lazy loading)
        this.validateCredentials();        try {
          return await handleM365Groups(this.getGraphClient(), args);
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
    );    // Exchange Settings - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_exchange_settings",
      exchangeSettingsSchema.shape,
      wrapToolHandler(async (args: ExchangeSettingsArgs) => {
        // Validate credentials only when tool is executed (lazy loading)
        this.validateCredentials();
        try {
          return await handleExchangeSettings(this.getGraphClient(), args);
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
    );    // User Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_user_settings",
      userManagementSchema.shape,
      wrapToolHandler(async (args: UserManagementArgs) => {
        // Validate credentials only when tool is executed (lazy loading)
        this.validateCredentials();
        try {
          return await handleUserSettings(this.getGraphClient(), args);
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
    );    // Offboarding - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_offboarding",
      offboardingSchema.shape,
      wrapToolHandler(async (args: OffboardingArgs) => {
        // Validate credentials only when tool is executed (lazy loading)
        this.validateCredentials();
        try {
          return await handleOffboarding(this.getGraphClient(), args);
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
    );    // SharePoint Sites - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_sharepoint_sites",
      sharePointSiteSchema.shape,
      wrapToolHandler(async (args: SharePointSiteArgs) => {
        // Validate credentials only when tool is executed (lazy loading)
        this.validateCredentials();
        try {
          return await handleSharePointSite(this.getGraphClient(), args);
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
    );    // SharePoint Lists - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_sharepoint_lists",
      sharePointListSchema.shape,
      wrapToolHandler(async (args: SharePointListArgs) => {
        // Validate credentials only when tool is executed (lazy loading)
        this.validateCredentials();
        try {
          return await handleSharePointList(this.getGraphClient(), args);
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
    );    // Azure AD Roles - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_azure_ad_roles",
      azureAdRoleSchema.shape,
      wrapToolHandler(async (args: AzureAdRoleArgs) => {
        // Validate credentials only when tool is executed (lazy loading)
        this.validateCredentials();
        try {
          return await handleAzureAdRoles(this.getGraphClient(), args);
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
    );    // Azure AD Apps - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_azure_ad_apps",
      azureAdAppSchema.shape,
      wrapToolHandler(async (args: AzureAdAppArgs) => {
        // Validate credentials only when tool is executed (lazy loading)
        this.validateCredentials();
        try {
          return await handleAzureAdApps(this.getGraphClient(), args);
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

    // Azure AD Devices - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_azure_ad_devices",
      azureAdDeviceSchema.shape,
      wrapToolHandler(async (args: AzureAdDeviceArgs) => {
        // Validate credentials only when tool is executed (lazy loading)
        this.validateCredentials();
        try {
          return await handleAzureAdDevices(this.getGraphClient(), args);
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

    // Service Principals - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_service_principals",
      azureAdSpSchema.shape,
      wrapToolHandler(async (args: AzureAdSpArgs) => {
        // Validate credentials only when tool is executed (lazy loading)
        this.validateCredentials();
        try {
          return await handleServicePrincipals(this.getGraphClient(), args);
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

    // Dynamic API Endpoint - Lazy loading enabled for tool discovery
    this.server.tool(
      "call_microsoft_api",
      callMicrosoftApiSchema.shape,
      wrapToolHandler(async (args: CallMicrosoftApiArgs) => {
        // Validate credentials only when tool is executed (lazy loading)
        this.validateCredentials();
        try {
          return await handleCallMicrosoftApi(this.getGraphClient(), args, this.getAccessToken.bind(this), apiConfigs);
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

    // Audit Log - Lazy loading enabled for tool discovery
    this.server.tool(
      "search_audit_log",
      auditLogSchema.shape,
      wrapToolHandler(async (args: AuditLogArgs) => {
        // Validate credentials only when tool is executed (lazy loading)
        this.validateCredentials();
        try {
          return await handleSearchAuditLog(this.getGraphClient(), args);
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

    // Alerts - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_alerts",
      alertSchema.shape,
      wrapToolHandler(async (args: AlertArgs) => {
        // Validate credentials only when tool is executed (lazy loading)
        this.validateCredentials();
        try {
          return await handleManageAlerts(this.getGraphClient(), args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing tool: ${error instanceof Error ? error.message : 'Unknown error'}`
          );        }
      })
    );

    // DLP Policy Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_dlp_policies",
      dlpPolicySchema.shape,
      wrapToolHandler(async (args: DLPPolicyArgs) => {
        this.validateCredentials();
        try {
          return await handleDLPPolicies(this.getGraphClient(), args);
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

    // DLP Incident Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_dlp_incidents",
      dlpIncidentSchema.shape,
      wrapToolHandler(async (args: DLPIncidentArgs) => {
        this.validateCredentials();
        try {
          return await handleDLPIncidents(this.getGraphClient(), args);
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

    // Sensitivity Label Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_sensitivity_labels",
      sensitivityLabelSchema.shape,
      wrapToolHandler(async (args: DLPSensitivityLabelArgs) => {
        this.validateCredentials();
        try {
          return await handleDLPSensitivityLabels(this.getGraphClient(), args);
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

    // Intune macOS Device Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_intune_macos_devices",
      intuneMacOSDeviceSchema.shape,
      wrapToolHandler(async (args: IntuneMacOSDeviceArgs) => {
        this.validateCredentials();
        try {
          return await handleIntuneMacOSDevices(this.getGraphClient(), args);
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

    // Intune macOS Policy Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_intune_macos_policies",
      intuneMacOSPolicySchema.shape,
      wrapToolHandler(async (args: IntuneMacOSPolicyArgs) => {
        this.validateCredentials();
        try {
          return await handleIntuneMacOSPolicies(this.getGraphClient(), args);
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

    // Intune macOS App Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_intune_macos_apps",
      intuneMacOSAppSchema.shape,
      wrapToolHandler(async (args: IntuneMacOSAppArgs) => {
        this.validateCredentials();
        try {
          return await handleIntuneMacOSApps(this.getGraphClient(), args);
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

    // Intune macOS Compliance Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_intune_macos_compliance",
      intuneMacOSComplianceSchema.shape,
      wrapToolHandler(async (args: IntuneMacOSComplianceArgs) => {
        this.validateCredentials();
        try {
          return await handleIntuneMacOSCompliance(this.getGraphClient(), args);
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

    // Intune Windows Device Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_intune_windows_devices",
      intuneWindowsDeviceSchema.shape,
      wrapToolHandler(async (args: IntuneWindowsDeviceArgs) => {
        this.validateCredentials();
        try {
          return await handleIntuneWindowsDevices(this.getGraphClient(), args);
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

    // Intune Windows Policy Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_intune_windows_policies",
      intuneWindowsPolicySchema.shape,
      wrapToolHandler(async (args: IntuneWindowsPolicyArgs) => {
        this.validateCredentials();
        try {
          return await handleIntuneWindowsPolicies(this.getGraphClient(), args);
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

    // Intune Windows App Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_intune_windows_apps",
      intuneWindowsAppSchema.shape,
      wrapToolHandler(async (args: IntuneWindowsAppArgs) => {
        this.validateCredentials();
        try {
          return await handleIntuneWindowsApps(this.getGraphClient(), args);
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

    // Intune Windows Compliance Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_intune_windows_compliance",
      intuneWindowsComplianceSchema.shape,
      wrapToolHandler(async (args: IntuneWindowsComplianceArgs) => {
        this.validateCredentials();
        try {
          return await handleIntuneWindowsCompliance(this.getGraphClient(), args);
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

    // Compliance Framework Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_compliance_frameworks",
      complianceFrameworkSchema.shape,
      wrapToolHandler(async (args: ComplianceFrameworkArgs) => {
        this.validateCredentials();
        try {
          return await handleComplianceFrameworks(this.getGraphClient(), args);
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

    // Compliance Assessment Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_compliance_assessments",
      complianceAssessmentSchema.shape,
      wrapToolHandler(async (args: ComplianceAssessmentArgs) => {
        this.validateCredentials();
        try {
          return await handleComplianceAssessments(this.getGraphClient(), args);
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

    // Compliance Monitoring - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_compliance_monitoring",
      complianceMonitoringSchema.shape,
      wrapToolHandler(async (args: ComplianceMonitoringArgs) => {
        this.validateCredentials();
        try {
          return await handleComplianceMonitoring(this.getGraphClient(), args);
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

    // Evidence Collection - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_evidence_collection",
      evidenceCollectionSchema.shape,
      wrapToolHandler(async (args: EvidenceCollectionArgs) => {
        this.validateCredentials();
        try {
          return await handleEvidenceCollection(this.getGraphClient(), args);
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

    // Gap Analysis - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_gap_analysis",
      gapAnalysisSchema.shape,
      wrapToolHandler(async (args: GapAnalysisArgs) => {
        this.validateCredentials();
        try {
          return await handleGapAnalysis(this.getGraphClient(), args);
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

    // Audit Reports - Lazy loading enabled for tool discovery
    this.server.tool(
      "generate_audit_reports",
      auditReportSchema.shape,
      wrapToolHandler(async (args: AuditReportArgs) => {
        this.validateCredentials();
        try {
          return await handleAuditReports(this.getGraphClient(), args);
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

    // CIS Compliance - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_cis_compliance",
      cisComplianceSchema.shape,
      wrapToolHandler(async (args: CISComplianceArgs) => {
        this.validateCredentials();
        try {
          return await handleCISCompliance(this.getGraphClient(), args);
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

    // Advanced Graph API Features
    this.setupAdvancedGraphTools();
    
    // Dynamic tool generation (async - will generate tools in background)
    this.setupDynamicTools();
  }

  // Setup advanced Graph API tools
  private setupAdvancedGraphTools(): void {
    // Batch Operations
    this.server.tool(
      "execute_graph_batch",
      batchRequestSchema.shape,
      wrapToolHandler(async (args: any) => {
        this.validateCredentials();
        try {
          const advancedFeatures = new GraphAdvancedFeatures(this.getGraphClient(), this.getAccessToken.bind(this));
          const result = await advancedFeatures.executeBatch(args.requests);
          return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing batch operation: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Delta Queries
    this.server.tool(
      "execute_delta_query",
      deltaQuerySchema.shape,
      wrapToolHandler(async (args: any) => {
        this.validateCredentials();
        try {
          const advancedFeatures = new GraphAdvancedFeatures(this.getGraphClient(), this.getAccessToken.bind(this));
          const result = await advancedFeatures.executeDeltaQuery(args.resource, args.deltaToken);
          return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing delta query: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Webhook Subscriptions
    this.server.tool(
      "manage_graph_subscriptions",
      z.object({
        action: z.enum(['create', 'update', 'delete', 'list']).describe('Subscription management action'),
        subscriptionId: z.string().optional().describe('Subscription ID for update/delete operations'),
        subscription: webhookSubscriptionSchema.optional().describe('Subscription details for create/update'),
        updates: webhookSubscriptionSchema.partial().optional().describe('Updates for existing subscription')
      }).shape,
      wrapToolHandler(async (args: any) => {
        this.validateCredentials();
        try {
          const advancedFeatures = new GraphAdvancedFeatures(this.getGraphClient(), this.getAccessToken.bind(this));
          let result: any;

          switch (args.action) {
            case 'create':
              if (!args.subscription) {
                throw new McpError(ErrorCode.InvalidParams, 'Subscription details required for create action');
              }
              result = await advancedFeatures.createSubscription(args.subscription);
              break;
            case 'update':
              if (!args.subscriptionId || !args.updates) {
                throw new McpError(ErrorCode.InvalidParams, 'Subscription ID and updates required for update action');
              }
              result = await advancedFeatures.updateSubscription(args.subscriptionId, args.updates);
              break;
            case 'delete':
              if (!args.subscriptionId) {
                throw new McpError(ErrorCode.InvalidParams, 'Subscription ID required for delete action');
              }
              result = await advancedFeatures.deleteSubscription(args.subscriptionId);
              break;
            case 'list':
              result = await advancedFeatures.listSubscriptions();
              break;
            default:
              throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
          }

          return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error managing subscriptions: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Advanced Search
    this.server.tool(
      "execute_graph_search",
      searchQuerySchema.shape,
      wrapToolHandler(async (args: any) => {
        this.validateCredentials();
        try {
          const advancedFeatures = new GraphAdvancedFeatures(this.getGraphClient(), this.getAccessToken.bind(this));
          const result = await advancedFeatures.executeSearch(args);
          return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing search: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );
  }

  // Setup dynamic tools (async background process)
  private setupDynamicTools(): void {
    // Generate dynamic tools in the background after server initialization
    setTimeout(async () => {
      try {
        console.log('🚀 Initializing dynamic Graph API tool generation...');
        const dynamicGenerator = new DynamicToolGenerator(
          this.server,
          this.getGraphClient(),
          this.getAccessToken.bind(this),
          this.validateCredentials.bind(this)
        );
        
        await dynamicGenerator.generateAllTools();
        console.log('✅ Dynamic Graph API tools initialized successfully');
      } catch (error) {
        console.error('❌ Failed to initialize dynamic tools:', error);
        // Don't throw - this is a background process
      }
    }, 1000); // Delay to allow server to fully initialize
  }
  
  private setupResources(): void {
    // Register all M365 resources
    for (const resource of m365Resources) {
      this.server.resource(
        resource.name,
        resource.uri,
        async (uri: URL) => {
          try {
            this.validateCredentials();
            const params = new URLSearchParams(uri.search);
            const data = await resource.handler(this.getGraphClient(), params);
            
            return {
              contents: [
                {
                  uri: uri.href,
                  mimeType: resource.mimeType,
                  text: JSON.stringify(data, null, 2),
                },
              ],
            };
          } catch (error) {
            throw new McpError(
              ErrorCode.InternalError,
              `Error reading resource ${resource.name}: ${error instanceof Error ? error.message : 'Unknown error'}`
            );
          }
        }
      );
    }
    
    console.log(`✅ Registered ${m365Resources.length} MCP resources`);
  }

  private setupPrompts(): void {
    // Register all M365 prompts (general + Intune-specific)
    for (const prompt of allM365Prompts) {
      // Convert arguments array to Zod schema shape required by MCP SDK
      const argsShape: Record<string, any> = {};
      if (prompt.arguments) {
        for (const arg of prompt.arguments) {
          // Create Zod string schema with description
          const zodSchema = z.string().describe(arg.description);
          // Make it optional if not required
          argsShape[arg.name] = arg.required ? zodSchema : zodSchema.optional();
        }
      }
      
      this.server.prompt(
        prompt.name,
        prompt.description,
        argsShape,
        async (args: Record<string, string>) => {
          try {
            this.validateCredentials();
            const message = await prompt.handler(this.getGraphClient(), args);
            
            return {
              messages: [
                {
                  role: 'user',
                  content: {
                    type: 'text',
                    text: message
                  }
                }
              ]
            };
          } catch (error) {
            throw new McpError(
              ErrorCode.InternalError,
              `Error executing prompt ${prompt.name}: ${error instanceof Error ? error.message : 'Unknown error'}`
            );
          }
        }
      );
    }
    
    console.log(`✅ Registered ${allM365Prompts.length} MCP prompts`);
  }

  // SSE and real-time capabilities
  public addSSEClient(client: any): void {
    this.sseClients.add(client);
    console.log(`SSE client connected. Total clients: ${this.sseClients.size}`);
  }

  public removeSSEClient(client: any): void {
    this.sseClients.delete(client);
    console.log(`SSE client disconnected. Total clients: ${this.sseClients.size}`);
  }

  public broadcastUpdate(update: any): void {
    this.sseClients.forEach(client => {
      try {
        client.write(`data: ${JSON.stringify(update)}\n\n`);
      } catch (error) {
        this.sseClients.delete(client);
      }
    });
  }

  public reportProgress(operationId: string, progress: number, message?: string): void {
    const progressUpdate = {
      type: 'progress',
      operationId,
      progress,
      message,
      timestamp: new Date().toISOString()
    };
    
    this.progressTrackers.set(operationId, progressUpdate);
    this.broadcastUpdate(progressUpdate);
  }

  public completeOperation(operationId: string, result: any): void {
    const completion = {
      type: 'completion',
      operationId,
      result,
      timestamp: new Date().toISOString()
    };
    
    this.progressTrackers.delete(operationId);
    this.broadcastUpdate(completion);
  }

  public notifyResourceChange(resourceUri: string, changeType: 'created' | 'updated' | 'deleted'): void {
    const notification = {
      type: 'resourceChange',
      resourceUri,
      changeType,
      timestamp: new Date().toISOString()
    };
    
    this.broadcastUpdate(notification);
  }
}

// Start the server
async function main() {
  try {
    const server = new M365CoreServer();
    
    const transport = process.env.NODE_ENV === 'http'
      ? new StreamableHTTPServerTransport({
          sessionIdGenerator: () => randomUUID()
        })
      : new StdioServerTransport();

    await server.server.connect(transport);
    console.log(`M365 Core MCP Server running on ${process.env.NODE_ENV === 'http' ? `http://localhost:${PORT}` : 'stdio'}`);
  } catch (error) {
    console.error('Error starting M365 Core MCP Server:', error);
    process.exit(1);
  }
}

// ES Module check - only run main if this file is executed directly
if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch(error => {
    console.error('Unhandled error:', error);
    process.exit(1);
  });
}
