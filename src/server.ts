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

// Environment validation - will be checked lazily when tools are executed
const MS_TENANT_ID = process.env.MS_TENANT_ID ?? '';
const MS_CLIENT_ID = process.env.MS_CLIENT_ID ?? '';
const MS_CLIENT_SECRET = process.env.MS_CLIENT_SECRET ?? '';
const PORT = process.env.PORT ? parseInt(process.env.PORT, 10) : 3000;
const LOG_LEVEL = process.env.LOG_LEVEL ?? 'info';

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
  private graphClient: Client | null = null;
  private tokenCache: Map<string, { token: string; expiresOn: number }> = new Map();
  public sseClients: Set<any> = new Set();
  public progressTrackers: Map<string, any> = new Map();
    constructor() {
    this.server = new McpServer({
      name: 'm365-core-server',
      version: '1.0.0',
      capabilities: {
        tools: {
          listChanged: true
        },
        resources: {
          subscribe: true,
          listChanged: true
        },
        prompts: {
          listChanged: true
        },
        logging: {
          level: 'info'
        },
        experimental: {
          progressReporting: true,
          streamingResponses: true
        }
      }
    });

    this.setupTools();
    this.setupResources();
    
    process.on('SIGINT', async () => {
      await this.server.close();
      process.exit(0);
    });
  }

  private getGraphClient(): Client {
    if (!this.graphClient) {
      this.validateCredentials();
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

  private async getAccessToken(scope: string = apiConfigs.graph.scope): Promise<string> {
    const cached = this.tokenCache.get(scope);
    const now = Date.now();

    if (cached && cached.expiresOn > now + 60 * 1000) {
      return cached.token;
    }

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

    const expiresOn = now + data.expires_in * 1000;
    this.tokenCache.set(scope, { token: data.access_token, expiresOn });

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
  }  private setupTools(): void {
    // Distribution Lists - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_distribution_lists",
      distributionListSchema.shape,
      wrapToolHandler(async (args: DistributionListArgs) => {
        // Lazy loading: Validate credentials only when tool is executed
        this.validateCredentials();
        try {
          return await handleDistributionLists(this.getGraphClient(), args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing manage_distribution_lists: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Security Groups - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_security_groups",
      securityGroupSchema.shape,
      wrapToolHandler(async (args: SecurityGroupArgs) => {
        this.validateCredentials();
        try {
          return await handleSecurityGroups(this.getGraphClient(), args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing manage_security_groups: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // M365 Groups - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_m365_groups",
      m365GroupSchema.shape,
      wrapToolHandler(async (args: M365GroupArgs) => {
        this.validateCredentials();
        try {
          return await handleM365Groups(this.getGraphClient(), args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing manage_m365_groups: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Exchange Settings - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_exchange_settings",
      exchangeSettingsSchema.shape,
      wrapToolHandler(async (args: ExchangeSettingsArgs) => {
        this.validateCredentials();
        try {
          return await handleExchangeSettings(this.getGraphClient(), args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing manage_exchange_settings: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // User Management - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_user_settings",
      userManagementSchema.shape,
      wrapToolHandler(async (args: UserManagementArgs) => {
        this.validateCredentials();
        try {
          return await handleUserSettings(this.getGraphClient(), args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing manage_user_settings: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Offboarding - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_offboarding",
      offboardingSchema.shape,
      wrapToolHandler(async (args: OffboardingArgs) => {
        this.validateCredentials();
        try {
          return await handleOffboarding(this.getGraphClient(), args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing manage_offboarding: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // SharePoint Sites - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_sharepoint_sites",
      sharePointSiteSchema.shape,
      wrapToolHandler(async (args: SharePointSiteArgs) => {
        this.validateCredentials();
        try {
          return await handleSharePointSite(this.getGraphClient(), args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing manage_sharepoint_sites: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // SharePoint Lists - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_sharepoint_lists",
      sharePointListSchema.shape,
      wrapToolHandler(async (args: SharePointListArgs) => {
        this.validateCredentials();
        try {
          return await handleSharePointList(this.getGraphClient(), args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing manage_sharepoint_lists: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Azure AD Roles - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_azuread_roles",
      azureAdRoleSchema.shape,
      wrapToolHandler(async (args: AzureAdRoleArgs) => {
        this.validateCredentials();
        try {
          return await handleAzureAdRoles(this.getGraphClient(), args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing manage_azuread_roles: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Azure AD Apps - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_azuread_apps",
      azureAdAppSchema.shape,
      wrapToolHandler(async (args: AzureAdAppArgs) => {
        this.validateCredentials();
        try {
          return await handleAzureAdApps(this.getGraphClient(), args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing manage_azuread_apps: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Azure AD Devices - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_azuread_devices",
      azureAdDeviceSchema.shape,
      wrapToolHandler(async (args: AzureAdDeviceArgs) => {
        this.validateCredentials();
        try {
          return await handleAzureAdDevices(this.getGraphClient(), args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing manage_azuread_devices: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Service Principals - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_service_principals",
      azureAdSpSchema.shape,
      wrapToolHandler(async (args: AzureAdSpArgs) => {
        this.validateCredentials();
        try {
          return await handleServicePrincipals(this.getGraphClient(), args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing manage_service_principals: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );    // Microsoft API Call - Enhanced with lazy loading and better error handling
    this.server.tool(
      "dynamicendpoints m365 assistant",
      callMicrosoftApiSchema.shape,
      wrapToolHandler(async (args: CallMicrosoftApiArgs) => {
        this.validateCredentials();
        try {
          return await handleCallMicrosoftApi(this.getGraphClient(), args, this.getAccessToken.bind(this), apiConfigs);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing dynamicendpoints m365 assistant: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Audit Log Search - Enhanced with lazy loading and better error handling
    this.server.tool(
      "search_audit_log",
      auditLogSchema.shape,
      wrapToolHandler(async (args: AuditLogArgs) => {
        this.validateCredentials();
        try {
          return await handleSearchAuditLog(this.getGraphClient(), args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing search_audit_log: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Alert Management - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_alerts",
      alertSchema.shape,
      wrapToolHandler(async (args: AlertArgs) => {
        this.validateCredentials();
        try {
          return await handleManageAlerts(this.getGraphClient(), args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing manage_alerts: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // DLP Policies - Enhanced with lazy loading and better error handling
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
            `Error executing manage_dlp_policies: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // DLP Incidents - Enhanced with lazy loading and better error handling
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
            `Error executing manage_dlp_incidents: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Sensitivity Labels - Enhanced with lazy loading and better error handling
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
            `Error executing manage_sensitivity_labels: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Intune macOS Devices - Enhanced with lazy loading and better error handling
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
            `Error executing manage_intune_macos_devices: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Intune macOS Policies - Enhanced with lazy loading and better error handling
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
            `Error executing manage_intune_macos_policies: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Intune macOS Apps - Enhanced with lazy loading and better error handling
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
            `Error executing manage_intune_macos_apps: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Intune macOS Compliance - Enhanced with lazy loading and better error handling
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
            `Error executing manage_intune_macos_compliance: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Compliance Frameworks - Enhanced with lazy loading and better error handling
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
            `Error executing manage_compliance_frameworks: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Compliance Assessments - Enhanced with lazy loading and better error handling
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
            `Error executing manage_compliance_assessments: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Compliance Monitoring - Enhanced with lazy loading and better error handling
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
            `Error executing manage_compliance_monitoring: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Evidence Collection - Enhanced with lazy loading and better error handling
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
            `Error executing manage_evidence_collection: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Gap Analysis - Enhanced with lazy loading and better error handling
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
            `Error executing manage_gap_analysis: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Audit Reports - Enhanced with lazy loading and better error handling
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
            `Error executing generate_audit_reports: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // CIS Compliance - Enhanced with lazy loading and better error handling
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
            `Error executing manage_cis_compliance: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );
  }
  private setupResources(): void {
    // SharePoint Sites resource
    this.server.resource(
      'sharepoint_sites',
      new ResourceTemplate('sharepoint://sites/{siteId}', { list: undefined }),
      async (uri: URL, variables: any) => {
        try {
          const client = this.getGraphClient();
          const lists = await client.api(`/sites/${variables?.siteId}/lists`).get();
          
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

    // SharePoint Lists resource
    this.server.resource(
      'sharepoint_lists',
      new ResourceTemplate('sharepoint://sites/{siteId}/lists/{listId}', { list: undefined }),
      async (uri: URL, variables: any) => {
        try {
          const client = this.getGraphClient();
          const list = await client.api(`/sites/${variables?.siteId}/lists/${variables?.listId}`).get();
          
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

    // SharePoint List Items resource
    this.server.resource(
      'sharepoint_list_items',
      new ResourceTemplate('sharepoint://sites/{siteId}/lists/{listId}/items', { list: undefined }),
      async (uri: URL, variables: any) => {
        try {
          const client = this.getGraphClient();
          const items = await client.api(`/sites/${variables?.siteId}/lists/${variables?.listId}/items?expand=fields`).get();
          
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
    );    // Security alerts resource - moved to extended-resources.ts to avoid duplication

    // Security incidents resource
    this.server.resource(
      'security_incidents',
      'security://incidents',
      async (uri: URL) => {
        try {
          const client = this.getGraphClient();
          const incidents = await client.api('/security/incidents').get();
          
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
          sessionIdGenerator: () => randomUUID() // Added sessionIdGenerator
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