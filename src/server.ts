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

// Import authentication types
import type { AuthInfo } from './auth/oauth-provider.js';

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
  dlpPolicyArgsSchema,
  retentionPolicyArgsSchema,
  sensitivityLabelArgsSchema,
  informationProtectionPolicyArgsSchema,
  conditionalAccessPolicyArgsSchema,
  defenderPolicyArgsSchema,
  teamsPolicyArgsSchema,
  exchangePolicyArgsSchema,
  sharePointGovernancePolicyArgsSchema,
  securityAlertPolicyArgsSchema,
  powerPointPresentationArgsSchema,
  wordDocumentArgsSchema,
  htmlReportArgsSchema,
  professionalReportArgsSchema,
  oauthAuthorizationArgsSchema
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

// Import Microsoft Purview/Compliance policy handlers and types
import {
  handleRetentionPolicies,
  handleSensitivityLabels,
  handleInformationProtectionPolicies
} from './handlers/purview-compliance-handler.js';
import {
  RetentionPolicyArgs,
  SensitivityLabelArgs,
  InformationProtectionPolicyArgs,
  ConditionalAccessPolicyArgs,
  DefenderPolicyArgs,
  TeamsPolicyArgs,
  ExchangePolicyArgs,
  SharePointGovernancePolicyArgs,
  SecurityAlertPolicyArgs
} from './types/policy-types.js';
import { handleConditionalAccessPolicies } from './handlers/conditional-access-handler.js';
import {
  handleDefenderPolicies,
  handleTeamsPolicies,
  handleExchangePolicies,
  handleSharePointGovernancePolicies,
  handleSecurityAlertPolicies
} from './handlers/security-policy-handlers.js';

// Import document generation handlers and types
import { handlePowerPointPresentations } from './handlers/powerpoint-handler.js';
import { handleWordDocuments } from './handlers/word-document-handler.js';
import { handleHTMLReports } from './handlers/html-report-handler.js';
import { handleProfessionalReports } from './handlers/professional-report-handler.js';
import { handleOAuthAuthorization } from './handlers/oauth-handler.js';
import {
  PowerPointPresentationArgs,
  WordDocumentArgs,
  HTMLReportArgs,
  ProfessionalReportArgs,
  OAuthAuthorizationArgs
} from './types/document-generation-types.js';

import { handleExchangeSettings } from './exchange-handler.js';

// Import resources and prompts
import { m365Resources, getResourceByUri, listResources } from './resources.js';
import { allM365Prompts, getPromptByName, listPrompts } from './prompts.js';
import { toolMetadata, getToolMetadata } from './tool-metadata.js';

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
      version: '1.0.0'
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
    // Configuration is optional - server can run without credentials for discovery
    // Credentials are only required when actually executing tools
    const requiredEnvVars = ['MS_TENANT_ID', 'MS_CLIENT_ID', 'MS_CLIENT_SECRET'];
    const missingVars = requiredEnvVars.filter(varName => !getEnvVar(varName));
    
    if (missingVars.length > 0) {
      throw new McpError(
        ErrorCode.InvalidRequest,
        `Missing required environment variables for Microsoft 365 authentication: ${missingVars.join(', ')}. ` +
        `Please configure these variables to use this tool:\n\n` +
        `Environment Variables:\n` +
        `- MS_TENANT_ID: Your Azure AD tenant ID (e.g., 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx')\n` +
        `- MS_CLIENT_ID: Your Azure AD application (client) ID\n` +
        `- MS_CLIENT_SECRET: Your Azure AD application client secret\n\n` +
        `Quick Setup:\n` +
        `1. Register an app in Azure AD: https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade\n` +
        `2. Grant necessary Microsoft Graph API permissions\n` +
        `3. Create a client secret\n` +
        `4. Set environment variables with your credentials\n\n` +
        `Full setup guide: https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app`
      );
    }
  }  private setupTools(): void {
    // Distribution Lists
    const distributionListsMeta = getToolMetadata("manage_distribution_lists")!;
    this.server.tool(
      "manage_distribution_lists",
      distributionListsMeta.description,
      distributionListSchema.shape,
      distributionListsMeta.annotations || {},
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
      "Manage Azure AD security groups for access control, including group creation, membership, and security settings.",
      securityGroupSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
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
      "Manage Microsoft 365 groups for team collaboration with shared resources like mailbox, calendar, and files.",
      m365GroupSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
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
      "Manage Exchange Online settings including mailbox configuration, transport rules, and organization policies.",
      exchangeSettingsSchema.shape,
      {"readOnlyHint":false,"destructiveHint":false,"idempotentHint":true},
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
      "Manage user account settings including profile information, mailbox settings, licenses, and authentication methods.",
      userManagementSchema.shape,
      {"readOnlyHint":false,"destructiveHint":false,"idempotentHint":true},
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
      "Automate user offboarding processes including account disablement, license removal, data backup, and access revocation.",
      offboardingSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
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
      "Manage SharePoint sites including creation, configuration, permissions, and site collection administration.",
      sharePointSiteSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
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
      "Manage SharePoint lists and libraries including schema definition, items, views, and permissions.",
      sharePointListSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
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
      "Manage Azure AD administrative roles including role assignments, custom roles, and privilege escalation controls.",
      azureAdRoleSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
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
      "Manage Azure AD application registrations including app permissions, credentials, and OAuth configurations.",
      azureAdAppSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
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
      "Manage devices registered in Azure AD including device compliance, BitLocker keys, and device actions.",
      azureAdDeviceSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
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
      "Manage service principals for application access including permissions, credentials, and enterprise applications.",
      azureAdSpSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
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
      "Make direct calls to any Microsoft Graph or Azure Resource Management API endpoint with full control over HTTP methods and parameters.",
      callMicrosoftApiSchema.shape,
      {"readOnlyHint":false,"destructiveHint":false,"idempotentHint":false},
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
      "Search and analyze Azure AD unified audit logs for security events, user activities, and compliance monitoring.",
      auditLogSchema.shape,
      {"readOnlyHint":true,"destructiveHint":false,"idempotentHint":true},
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
      "Manage security alerts from Microsoft Defender and other security products including investigation and remediation.",
      alertSchema.shape,
      {"readOnlyHint":false,"destructiveHint":false,"idempotentHint":true},
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
      "Manage Data Loss Prevention policies to protect sensitive data across Exchange, SharePoint, OneDrive, and Teams.",
      dlpPolicySchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
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
      "Investigate and manage DLP policy violations and incidents including user notifications and remediation actions.",
      dlpIncidentSchema.shape,
      {"readOnlyHint":false,"destructiveHint":false,"idempotentHint":false},
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

    // Intune macOS Device Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_intune_macos_devices",
      "Manage macOS devices in Intune including enrollment, compliance policies, device actions, and inventory management.",
      intuneMacOSDeviceSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
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
      "Manage macOS configuration profiles and compliance policies for device security and management settings.",
      intuneMacOSPolicySchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
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
      "Manage macOS application deployment including app assignments, updates, and installation requirements.",
      intuneMacOSAppSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
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
      "Assess macOS device compliance status and generate reports on policy adherence and security posture.",
      intuneMacOSComplianceSchema.shape,
      {"readOnlyHint":true,"destructiveHint":false,"idempotentHint":true},
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
      "Manage Windows devices in Intune including enrollment, autopilot deployment, device actions, and health monitoring.",
      intuneWindowsDeviceSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
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
      "Manage Windows configuration profiles and compliance policies including security baselines and update rings.",
      intuneWindowsPolicySchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
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
      "Manage Windows application deployment including Win32 apps, Microsoft Store apps, and Office 365 assignments.",
      intuneWindowsAppSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
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
      "Assess Windows device compliance status including BitLocker encryption, antivirus status, and security configurations.",
      intuneWindowsComplianceSchema.shape,
      {"readOnlyHint":true,"destructiveHint":false,"idempotentHint":true},
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
      "Manage compliance frameworks and standards including HIPAA, GDPR, SOX, PCI-DSS, ISO 27001, and NIST configurations.",
      complianceFrameworkSchema.shape,
      {"readOnlyHint":false,"destructiveHint":false,"idempotentHint":true},
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
      "Conduct compliance assessments and generate detailed reports on regulatory adherence and security controls.",
      complianceAssessmentSchema.shape,
      {"readOnlyHint":true,"destructiveHint":false,"idempotentHint":true},
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
      "Monitor ongoing compliance status with real-time alerts for policy violations and regulatory changes.",
      complianceMonitoringSchema.shape,
      {"readOnlyHint":true,"destructiveHint":false,"idempotentHint":true},
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
      "Collect and preserve compliance evidence including audit logs, configuration snapshots, and attestation records.",
      evidenceCollectionSchema.shape,
      {"readOnlyHint":true,"destructiveHint":false,"idempotentHint":true},
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
      "Perform gap analysis to identify compliance deficiencies and generate remediation recommendations.",
      gapAnalysisSchema.shape,
      {"readOnlyHint":true,"destructiveHint":false,"idempotentHint":true},
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
      "Generate comprehensive audit reports for compliance frameworks with evidence documentation and findings.",
      auditReportSchema.shape,
      {"readOnlyHint":true,"destructiveHint":false,"idempotentHint":true},
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
      "Manage CIS (Center for Internet Security) benchmark compliance including assessment and remediation tracking.",
      cisComplianceSchema.shape,
      {"readOnlyHint":false,"destructiveHint":false,"idempotentHint":true},
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

    // ===== Microsoft 365 Policy Management Tools =====

    // Retention Policy Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_retention_policies",
      "Manage retention policies for content across Exchange, SharePoint, OneDrive, and Teams with lifecycle rules.",
      retentionPolicyArgsSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
      wrapToolHandler(async (args: RetentionPolicyArgs) => {
        this.validateCredentials();
        try {
          return await handleRetentionPolicies(this.getGraphClient(), args);
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

    // Sensitivity Label Management
    const sensitivityLabelsMeta = getToolMetadata("manage_sensitivity_labels")!;
    this.server.tool(
      "manage_sensitivity_labels",
      sensitivityLabelsMeta.description,
      sensitivityLabelArgsSchema.shape,
      sensitivityLabelsMeta.annotations || {},
      wrapToolHandler(async (args: SensitivityLabelArgs) => {
        this.validateCredentials();
        try {
          return await handleSensitivityLabels(this.getGraphClient(), args);
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

    // Information Protection Policy Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_information_protection_policies",
      "Manage Azure Information Protection policies for data classification, encryption, and rights management.",
      informationProtectionPolicyArgsSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
      wrapToolHandler(async (args: InformationProtectionPolicyArgs) => {
        this.validateCredentials();
        try {
          return await handleInformationProtectionPolicies(this.getGraphClient(), args);
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

    // Conditional Access Policy Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_conditional_access_policies",
      "Manage Azure AD conditional access policies for zero-trust security including MFA, device compliance, and location-based controls.",
      conditionalAccessPolicyArgsSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
      wrapToolHandler(async (args: ConditionalAccessPolicyArgs) => {
        this.validateCredentials();
        try {
          return await handleConditionalAccessPolicies(this.getGraphClient(), args);
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

    // Microsoft Defender for Office 365 Policy Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_defender_policies",
      "Manage Microsoft Defender for Office 365 policies including Safe Attachments, Safe Links, anti-phishing, and anti-malware.",
      defenderPolicyArgsSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
      wrapToolHandler(async (args: DefenderPolicyArgs) => {
        this.validateCredentials();
        try {
          return await handleDefenderPolicies(this.getGraphClient(), args);
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

    // Microsoft Teams Policy Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_teams_policies",
      "Manage Microsoft Teams policies for messaging, meetings, calling, apps, and live events across the organization.",
      teamsPolicyArgsSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
      wrapToolHandler(async (args: TeamsPolicyArgs) => {
        this.validateCredentials();
        try {
          return await handleTeamsPolicies(this.getGraphClient(), args);
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

    // Exchange Online Policy Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_exchange_policies",
      "Manage Exchange Online policies including mail flow rules, mobile device access, and organization-wide settings.",
      exchangePolicyArgsSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
      wrapToolHandler(async (args: ExchangePolicyArgs) => {
        this.validateCredentials();
        try {
          return await handleExchangePolicies(this.getGraphClient(), args);
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

    // SharePoint Governance Policy Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_sharepoint_governance_policies",
      "Manage SharePoint governance policies including sharing controls, access restrictions, and site lifecycle management.",
      sharePointGovernancePolicyArgsSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
      wrapToolHandler(async (args: SharePointGovernancePolicyArgs) => {
        this.validateCredentials();
        try {
          return await handleSharePointGovernancePolicies(this.getGraphClient(), args);
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

    // Security and Compliance Alert Policy Management - Lazy loading enabled for tool discovery
    this.server.tool(
      "manage_security_alert_policies",
      "Manage security alert policies for monitoring threats, suspicious activities, and compliance violations across Microsoft 365.",
      securityAlertPolicyArgsSchema.shape,
      {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false},
      wrapToolHandler(async (args: SecurityAlertPolicyArgs) => {
        this.validateCredentials();
        try {
          return await handleSecurityAlertPolicies(this.getGraphClient(), args);
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

    // Document Generation Tools
    this.setupDocumentGenerationTools();

    // Advanced Graph API Features
    this.setupAdvancedGraphTools();
    
    // Dynamic tool generation (async - will generate tools in background)
    this.setupDynamicTools();
  }

  // Setup document generation tools
  private setupDocumentGenerationTools(): void {
    // PowerPoint Presentation Generation
    this.server.tool(
      "generate_powerpoint_presentation",
      "Create professional PowerPoint presentations with custom slides, charts, tables, and themes from Microsoft 365 data.",
      powerPointPresentationArgsSchema.shape,
      {"readOnlyHint":false,"destructiveHint":false,"idempotentHint":false},
      wrapToolHandler(async (args: PowerPointPresentationArgs) => {
        this.validateCredentials();
        try {
          const result = await handlePowerPointPresentations(args, this.getGraphClient());
          return { content: [{ type: 'text', text: result }] };
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error generating PowerPoint presentation: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Word Document Generation
    this.server.tool(
      "generate_word_document",
      "Create professional Word documents with formatted sections, tables, charts, and table of contents from analysis data.",
      wordDocumentArgsSchema.shape,
      {"readOnlyHint":false,"destructiveHint":false,"idempotentHint":false},
      wrapToolHandler(async (args: WordDocumentArgs) => {
        this.validateCredentials();
        try {
          const result = await handleWordDocuments(args, this.getGraphClient());
          return { content: [{ type: 'text', text: result }] };
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error generating Word document: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // HTML Report Generation
    this.server.tool(
      "generate_html_report",
      "Create interactive HTML reports and dashboards with responsive design, charts, and filtering capabilities.",
      htmlReportArgsSchema.shape,
      {"readOnlyHint":false,"destructiveHint":false,"idempotentHint":false},
      wrapToolHandler(async (args: HTMLReportArgs) => {
        this.validateCredentials();
        try {
          const result = await handleHTMLReports(args, this.getGraphClient());
          return { content: [{ type: 'text', text: result }] };
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error generating HTML report: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Professional Report Generation (Multi-format)
    this.server.tool(
      "generate_professional_report",
      "Generate comprehensive professional reports in multiple formats (PowerPoint, Word, HTML, PDF) from Microsoft 365 data.",
      professionalReportArgsSchema.shape,
      {"readOnlyHint":false,"destructiveHint":false,"idempotentHint":false},
      wrapToolHandler(async (args: ProfessionalReportArgs) => {
        this.validateCredentials();
        try {
          const result = await handleProfessionalReports(args, this.getGraphClient());
          return { content: [{ type: 'text', text: result }] };
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error generating professional report: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // OAuth Authorization
    this.server.tool(
      "oauth_authorize",
      "Manage OAuth 2.0 authorization for user-delegated access to OneDrive and SharePoint files with secure token handling.",
      oauthAuthorizationArgsSchema.shape,
      {"readOnlyHint":false,"destructiveHint":false,"idempotentHint":true},
      wrapToolHandler(async (args: OAuthAuthorizationArgs) => {
        try {
          // OAuth doesn't require Graph client validation - it handles its own token management
          const clientId = getEnvVar('MS_CLIENT_ID');
          const clientSecret = getEnvVar('MS_CLIENT_SECRET');
          const tenantId = getEnvVar('MS_TENANT_ID');
          const redirectUri = getEnvVar('MS_REDIRECT_URI') || 'http://localhost:3000/auth/callback';
          
          const result = await handleOAuthAuthorization(args, clientId, clientSecret, tenantId, redirectUri);
          return { content: [{ type: 'text', text: result }] };
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error handling OAuth authorization: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );
  }

  // Setup advanced Graph API tools
  private setupAdvancedGraphTools(): void {
    // Batch Operations
    this.server.tool(
      "execute_graph_batch",
      "Execute multiple Microsoft Graph API requests in a single batch operation for improved performance and efficiency.",
      batchRequestSchema.shape,
      {"readOnlyHint":false,"destructiveHint":false,"idempotentHint":false},
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
      "Track incremental changes to Microsoft Graph resources using delta queries for efficient synchronization.",
      deltaQuerySchema.shape,
      {"readOnlyHint":true,"destructiveHint":false,"idempotentHint":true},
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
    const graphSubscriptionsMeta = getToolMetadata("manage_graph_subscriptions")!;
    this.server.tool(
      "manage_graph_subscriptions",
      graphSubscriptionsMeta.description,
      z.object({
        action: z.enum(['create', 'update', 'delete', 'list']).describe('Subscription management action'),
        subscriptionId: z.string().optional().describe('Subscription ID for update/delete operations'),
        subscription: webhookSubscriptionSchema.optional().describe('Subscription details for create/update'),
        updates: webhookSubscriptionSchema.partial().optional().describe('Updates for existing subscription')
      }).shape,
      graphSubscriptionsMeta.annotations || {},
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
      "Execute advanced search queries across Microsoft 365 content including emails, files, messages, and calendar events.",
      searchQuerySchema.shape,
      {"readOnlyHint":true,"destructiveHint":false,"idempotentHint":true},
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
