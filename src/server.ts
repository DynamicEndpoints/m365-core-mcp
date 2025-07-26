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

// Enhanced utility classes for improved performance and reliability
class TokenCache {
  private cache: { [key: string]: { token: string; expiresAt: number } } = {};

  get(scope: string): string | null {
    const now = Date.now();
    const cached = this.cache[scope];
    
    if (cached && cached.expiresAt > now + 60000) { // 1 minute buffer
      return cached.token;
    }
    
    return null;
  }

  set(scope: string, token: string, expiresAt?: number): void {
    const now = Date.now();
    this.cache[scope] = {
      token,
      expiresAt: expiresAt || (now + 3600000) // 1 hour default
    };
  }

  clear(scope?: string): void {
    if (scope) {
      delete this.cache[scope];
    } else {
      this.cache = {};
    }
  }
}

class RateLimiter {
  private requests: number[] = [];
  private readonly maxRequests: number;
  private readonly windowMs: number;

  constructor(maxRequests: number = 100, windowMs: number = 60000) {
    this.maxRequests = maxRequests;
    this.windowMs = windowMs;
  }

  async checkLimit(): Promise<void> {
    const now = Date.now();
    this.requests = this.requests.filter(time => now - time < this.windowMs);
    
    if (this.requests.length >= this.maxRequests) {
      const oldestRequest = Math.min(...this.requests);
      const waitTime = this.windowMs - (now - oldestRequest);
      console.log(`Rate limit reached, waiting ${waitTime}ms`);
      await new Promise(resolve => setTimeout(resolve, waitTime));
    }
    
    this.requests.push(now);
  }

  getStats(): { current: number; max: number; windowMs: number } {
    const now = Date.now();
    this.requests = this.requests.filter(time => now - time < this.windowMs);
    return {
      current: this.requests.length,
      max: this.maxRequests,
      windowMs: this.windowMs
    };
  }
}

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
  CreateIntunePolicyArgs,
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
  azureAdRoleSchema,  azureAdAppSchema,
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
  cisComplianceSchema,  m365CoreTools,
} from './tool-definitions.js';

import { intuneTools, createIntunePolicySchema } from './tool-definitions-intune.js';
import { enhancedIntuneTools } from './tool-definitions-intune-enhanced.js';

import { handleCreateIntunePolicy } from './handlers/intune-handler.js';
import { handleCreateIntunePolicy as handleCreateIntunePolicyEnhanced } from './handlers/intune-handler-enhanced.js';
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
import { setupExtendedResources } from './extended-resources.js';

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
  
  // Enhanced utility instances
  private enhancedTokenCache: TokenCache = new TokenCache();
  private rateLimiter: RateLimiter = new RateLimiter();
  
  // Lazy loading state
  private isAuthenticated: boolean = false;
  private authenticationPromise: Promise<void> | null = null;
  private toolsRegistered: boolean = false;
  private resourcesRegistered: boolean = false;
  constructor() {
    this.server = new McpServer({
      name: 'm365-core-server',
      version: '1.1.0', // Enhanced version with improved API capabilities and lazy loading
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
      }    });

    // Initialize lazy loading - register tools and resources immediately for discovery
    // but authentication will still be lazy (only when tools are executed)
    this.setupLazyLoading();
    this.setupTools(); // Register tools immediately for Smithery discovery
    this.setupResources(); // Register resources immediately for Smithery discovery
    this.toolsRegistered = true;
    this.resourcesRegistered = true;
    
    process.on('SIGINT', async () => {
      await this.server.close();
      process.exit(0);
    });
  }
  // Lazy loading setup - tools and resources are registered for discovery, auth is lazy
  private setupLazyLoading(): void {
    console.log('🚀 Setting up lazy authentication for M365 Core MCP Server');
    console.log('   Tools and resources registered immediately for discovery');
    console.log('   Authentication will occur only when tools are executed');
  }

  // Ensure authentication is performed before tool execution
  async ensureAuthenticated(): Promise<void> {
    if (this.isAuthenticated) {
      return;
    }

    if (this.authenticationPromise) {
      return this.authenticationPromise;
    }

    this.authenticationPromise = this.performAuthentication();
    return this.authenticationPromise;
  }

  // Perform the actual authentication
  private async performAuthentication(): Promise<void> {
    try {
      console.log('🔐 Performing authentication on demand...');
      this.validateCredentials();
      
      // Test authentication by getting a token
      await this.getAccessToken(apiConfigs.graph.scope);
      
      this.isAuthenticated = true;
      console.log('✅ Authentication successful');
    } catch (error) {
      this.authenticationPromise = null;
      console.error('❌ Authentication failed:', error);
      throw error;
    }
  }

  // Ensure tools are registered (lazy loading)
  async ensureToolsRegistered(): Promise<void> {
    if (this.toolsRegistered) {
      return;
    }

    console.log('🔧 Registering tools on first use...');
    this.setupTools();
    this.toolsRegistered = true;
    console.log('✅ Tools registered successfully');
  }

  // Ensure resources are registered (lazy loading)
  async ensureResourcesRegistered(): Promise<void> {
    if (this.resourcesRegistered) {
      return;
    }

    console.log('📁 Registering resources on first use...');
    this.setupResources();
    this.resourcesRegistered = true;
    console.log('✅ Resources registered successfully');
  }

  private getGraphClient(): Client {
    if (!this.graphClient) {
      // Note: Credentials will be validated when first API call is made
      this.graphClient = Client.init({
        authProvider: async (callback: (error: Error | null, token: string | null) => void) => {
          try {
            // Validate credentials here when token is actually needed
            this.validateCredentials();
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
  }

  // Check if credentials are available without throwing errors
  private hasValidCredentials(): boolean {
    const requiredEnvVars = ['MS_TENANT_ID', 'MS_CLIENT_ID', 'MS_CLIENT_SECRET'];
    return requiredEnvVars.every(varName => !!process.env[varName]);
  }  private setupTools(): void {
    // Health Check Tool - No authentication required
    this.server.tool(
      "health_check",
      "Check server status and authentication configuration without requiring credentials",
      z.object({}).shape,
      wrapToolHandler(async () => {
        const hasCredentials = this.hasValidCredentials();
        const status = {
          serverStatus: "running",
          version: "1.1.0",
          timestamp: new Date().toISOString(),
          authentication: {
            configured: hasCredentials,
            requiredVariables: ["MS_TENANT_ID", "MS_CLIENT_ID", "MS_CLIENT_SECRET"],
            status: hasCredentials ? "ready" : "requires_configuration"
          },
          capabilities: {
            tools: true,
            resources: true,
            prompts: true,
            progressReporting: true,
            streamingResponses: true
          }
        };
        
        return {
          content: [
            {
              type: "text",
              text: `M365 Core MCP Server Health Check\n\n${JSON.stringify(status, null, 2)}\n\n` +
                   `${hasCredentials ? 
                     '✅ Server is ready for Microsoft 365 operations' : 
                     '⚠️  Server is running but requires environment variable configuration for Microsoft 365 operations'}`
            }
          ]
        };
      })
    );    // Distribution Lists - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_distribution_lists",
      "Create, update, delete, and manage Exchange distribution lists with members and properties",
      distributionListSchema.shape,
      wrapToolHandler(async (args: DistributionListArgs) => {
        // Lazy authentication: Only authenticate when tool is executed
        await this.ensureAuthenticated();
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
      "Create, update, delete, and manage Azure AD security groups with members and properties",
      securityGroupSchema.shape,
      wrapToolHandler(async (args: SecurityGroupArgs) => {
        await this.ensureAuthenticated();
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
    );// M365 Groups - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_m365_groups",
      "Create, update, delete, and manage Microsoft 365 groups with Teams integration and member management",
      m365GroupSchema.shape,
      wrapToolHandler(async (args: M365GroupArgs) => {
        await this.ensureAuthenticated();
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
    );    // Exchange Settings - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_exchange_settings",
      "Configure and manage Exchange Online settings including mailbox configurations, transport rules, and mail flow",
      exchangeSettingsSchema.shape,
      wrapToolHandler(async (args: ExchangeSettingsArgs) => {
        await this.ensureAuthenticated();
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
    );    // User Management - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_user_settings",
      "Get, update, and manage Azure AD user settings including profiles, licenses, and account properties",
      userManagementSchema.shape,
      wrapToolHandler(async (args: UserManagementArgs) => {
        await this.ensureAuthenticated();
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
    );    // Offboarding - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_offboarding",
      "Securely offboard users by disabling accounts, revoking access, transferring data, and managing group memberships",
      offboardingSchema.shape,
      wrapToolHandler(async (args: OffboardingArgs) => {
        await this.ensureAuthenticated();
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
    );    // SharePoint Sites - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_sharepoint_sites",
      "Create, update, delete, and manage SharePoint sites including permissions, properties, and site collections",
      sharePointSiteSchema.shape,
      wrapToolHandler(async (args: SharePointSiteArgs) => {
        await this.ensureAuthenticated();
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
    );    // SharePoint Lists - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_sharepoint_lists",
      "Create, update, delete, and manage SharePoint lists including columns, items, views, and permissions",
      sharePointListSchema.shape,
      wrapToolHandler(async (args: SharePointListArgs) => {
        await this.ensureAuthenticated();
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
    );    // Azure AD Roles - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_azuread_roles",
      "Assign, remove, and manage Azure AD directory roles and role memberships for users and groups",
      azureAdRoleSchema.shape,
      wrapToolHandler(async (args: AzureAdRoleArgs) => {
        await this.ensureAuthenticated();
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
      "Create, update, delete, and manage Azure AD application registrations including permissions and certificates",
      azureAdAppSchema.shape,
      wrapToolHandler(async (args: AzureAdAppArgs) => {
        await this.ensureAuthenticated();
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
    );    // Azure AD Devices - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_azuread_devices",
      "Register, update, delete, and manage Azure AD joined devices including compliance and configuration",
      azureAdDeviceSchema.shape,
      wrapToolHandler(async (args: AzureAdDeviceArgs) => {
        await this.ensureAuthenticated();
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
    );    // Service Principals - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_service_principals",
      "Create, update, delete, and manage Azure AD service principals including credentials and permissions",
      azureAdSpSchema.shape,
      wrapToolHandler(async (args: AzureAdSpArgs) => {
        await this.ensureAuthenticated();
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
      })    );    // Microsoft API Call - Enhanced with performance and reliability features
    this.server.tool(
      "dynamicendpoints m365 assistant",
      "Enhanced Microsoft Graph and Azure Resource Management API client with retry logic, rate limiting, field selection, and multiple response formats",
      callMicrosoftApiSchema.shape,
      wrapToolHandler(async (args: CallMicrosoftApiArgs) => {
        await this.ensureAuthenticated();
        try {
          return await handleCallMicrosoftApi(
            this.getGraphClient(), 
            args, 
            this.getAccessToken.bind(this), 
            apiConfigs,
            this.rateLimiter,
            this.enhancedTokenCache
          );
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
    );    // Audit Log Search - Enhanced with lazy loading and better error handling
    this.server.tool(
      "search_audit_log",
      "Search and retrieve Microsoft 365 audit logs with filtering by date, user, activity, and workload",
      auditLogSchema.shape,
      wrapToolHandler(async (args: AuditLogArgs) => {
        await this.ensureAuthenticated();
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
    );    // Alert Management - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_alerts",
      "List, get, and manage Microsoft 365 security alerts with filtering and status updates",
      alertSchema.shape,
      wrapToolHandler(async (args: AlertArgs) => {
        await this.ensureAuthenticated();
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
    );    // DLP Policies - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_dlp_policies",
      "Create, update, delete, and manage Data Loss Prevention (DLP) policies including rules, conditions, and actions",
      dlpPolicySchema.shape,
      wrapToolHandler(async (args: DLPPolicyArgs) => {
        await this.ensureAuthenticated();
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
    );    // DLP Incidents - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_dlp_incidents",
      "Investigate, review, and manage Data Loss Prevention (DLP) incidents including status updates and remediation",
      dlpIncidentSchema.shape,
      wrapToolHandler(async (args: DLPIncidentArgs) => {
        await this.ensureAuthenticated();
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
    );    // Sensitivity Labels - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_sensitivity_labels",
      "Create, update, delete, and manage Microsoft Purview sensitivity labels including policies and auto-labeling",
      sensitivityLabelSchema.shape,
      wrapToolHandler(async (args: DLPSensitivityLabelArgs) => {
        await this.ensureAuthenticated();
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
    );    // Intune macOS Devices - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_intune_macos_devices",
      "Manage Intune macOS devices including enrollment, compliance, configuration, and remote actions",
      intuneMacOSDeviceSchema.shape,
      wrapToolHandler(async (args: IntuneMacOSDeviceArgs) => {
        await this.ensureAuthenticated();
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
      "Create, update, delete, and manage Intune macOS configuration policies including device restrictions and compliance",
      intuneMacOSPolicySchema.shape,
      wrapToolHandler(async (args: IntuneMacOSPolicyArgs) => {
        await this.ensureAuthenticated();
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
      "Deploy, update, remove, and manage macOS applications through Microsoft Intune including assignment and monitoring",
      intuneMacOSAppSchema.shape,
      wrapToolHandler(async (args: IntuneMacOSAppArgs) => {
        await this.ensureAuthenticated();
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
      "Configure and manage macOS device compliance policies in Intune including requirements and actions",
      intuneMacOSComplianceSchema.shape,
      wrapToolHandler(async (args: IntuneMacOSComplianceArgs) => {
        await this.ensureAuthenticated();
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
      })    );

    // Intune Windows Devices - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_intune_windows_devices",
      "Manage Intune Windows devices including enrollment, compliance, configuration, and remote actions",
      intuneWindowsDeviceSchema.shape,
      wrapToolHandler(async (args: IntuneWindowsDeviceArgs) => {
        await this.ensureAuthenticated();
        try {
          return await handleIntuneWindowsDevices(this.getGraphClient(), args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing manage_intune_windows_devices: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Intune Windows Policies - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_intune_windows_policies",
      "Create, update, delete, and manage Intune Windows configuration policies including device restrictions and security",
      intuneWindowsPolicySchema.shape,
      wrapToolHandler(async (args: IntuneWindowsPolicyArgs) => {
        await this.ensureAuthenticated();
        try {
          return await handleIntuneWindowsPolicies(this.getGraphClient(), args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing manage_intune_windows_policies: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Intune Windows Apps - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_intune_windows_apps",
      "Deploy, update, remove, and manage Windows applications through Microsoft Intune including assignment and monitoring",
      intuneWindowsAppSchema.shape,
      wrapToolHandler(async (args: IntuneWindowsAppArgs) => {
        await this.ensureAuthenticated();
        try {
          return await handleIntuneWindowsApps(this.getGraphClient(), args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing manage_intune_windows_apps: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Intune Windows Compliance - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_intune_windows_compliance",
      "Configure and manage Windows device compliance policies in Intune including requirements and actions",
      intuneWindowsComplianceSchema.shape,
      wrapToolHandler(async (args: IntuneWindowsComplianceArgs) => {
        await this.ensureAuthenticated();
        try {
          return await handleIntuneWindowsCompliance(this.getGraphClient(), args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing manage_intune_windows_compliance: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Compliance Frameworks - Enhanced with lazy loading and better error handling
    this.server.tool(
      "manage_compliance_frameworks",
      "Assess and manage compliance against various frameworks (SOC2, ISO27001, NIST, GDPR, HIPAA)",
      complianceFrameworkSchema.shape,
      wrapToolHandler(async (args: ComplianceFrameworkArgs) => {
        await this.ensureAuthenticated();
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
      "Create, run, and manage compliance assessments with automated scoring and gap analysis",
      complianceAssessmentSchema.shape,
      wrapToolHandler(async (args: ComplianceAssessmentArgs) => {
        await this.ensureAuthenticated();
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
      "Monitor compliance status in real-time with alerts, reporting, and automated remediation workflows",
      complianceMonitoringSchema.shape,
      wrapToolHandler(async (args: ComplianceMonitoringArgs) => {
        await this.ensureAuthenticated();
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
      "Collect, organize, and manage compliance evidence including automated evidence gathering and validation",
      evidenceCollectionSchema.shape,
      wrapToolHandler(async (args: EvidenceCollectionArgs) => {
        await this.ensureAuthenticated();
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
      "Perform gap analysis against compliance frameworks with prioritized remediation recommendations",
      gapAnalysisSchema.shape,
      wrapToolHandler(async (args: GapAnalysisArgs) => {
        await this.ensureAuthenticated();
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
      "Generate comprehensive audit reports for compliance frameworks including evidence mapping, findings, and executive summaries in multiple formats",
      auditReportSchema.shape,
      wrapToolHandler(async (args: AuditReportArgs) => {
        await this.ensureAuthenticated();
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
      "Assess and manage CIS (Center for Internet Security) compliance benchmarks and controls",
      cisComplianceSchema.shape,
      wrapToolHandler(async (args: CISComplianceArgs) => {
        await this.ensureAuthenticated();
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
    );    // Intune Policy Creation Tool - Enhanced for accurate policy creation
    this.server.tool(
      "create_intune_policy",
      "Create accurate and complete Intune policies for Windows or macOS with validated settings and proper structure",
      createIntunePolicySchema.shape,
      wrapToolHandler(async (args: CreateIntunePolicyArgs) => {
        await this.ensureAuthenticated();
        try {
          return await handleCreateIntunePolicy(this.getGraphClient(), args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing create_intune_policy: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );    // Intune Policy Creation Tool - Unified and schema-driven
    intuneTools.forEach(tool => {
      this.server.tool(
        tool.name,
        tool.description,
        (tool.inputSchema as any).shape,
        wrapToolHandler(async (args: any) => {
          await this.ensureAuthenticated();
          try {
            return await handleCreateIntunePolicy(this.getGraphClient(), args);
          } catch (error) {
            if (error instanceof McpError) {
              throw error;
            }
            throw new McpError(
              ErrorCode.InternalError,
              `Error executing ${tool.name}: ${error instanceof Error ? error.message : 'Unknown error'}`
            );
          }
        })
      );
    });

    // Enhanced Intune Policy Creation Tools - Advanced validation and templates
    enhancedIntuneTools.forEach(tool => {
      this.server.tool(
        `enhanced_${tool.name}`,
        tool.description,
        (tool.inputSchema as any).shape,
        wrapToolHandler(async (args: any) => {
          await this.ensureAuthenticated();
          try {
            return await handleCreateIntunePolicyEnhanced(this.getGraphClient(), args);
          } catch (error) {
            if (error instanceof McpError) {
              throw error;
            }
            throw new McpError(
              ErrorCode.InternalError,
              `Error executing enhanced_${tool.name}: ${error instanceof Error ? error.message : 'Unknown error'}`
            );
          }
        })
      );
    });
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
    );    // Security alerts resource - moved to extended-resources.ts to avoid duplication    // Security incidents resource - using modern API
    this.server.resource(
      'security_incidents',
      'security://incidents',
      async (uri: URL) => {
        try {
          const incidents = await this.makeGraphApiCall('/security/incidents', {
            select: ['id', 'displayName', 'status', 'severity', 'createdDateTime', 'lastUpdateDateTime'],
            top: 50
          });
          
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
            `Error reading security incidents: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );

    // URI template resources for dynamic access
    this.server.resource(
      'security_alert_details',
      new ResourceTemplate('m365://security/alerts/{alertId}', { list: undefined }),
      async (uri: URL, variables: any) => {
        try {
          const alert = await this.makeGraphApiCall(`/security/alerts_v2/${variables?.alertId}`, {
            select: ['id', 'displayName', 'severity', 'status', 'createdDateTime', 'evidence']
          });
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(alert, null, 2),
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
    );    this.server.resource(
      'device_details',
      new ResourceTemplate('m365://devices/{deviceId}', { list: undefined }),
      async (uri: URL, variables: any) => {
        try {
          const client = this.getGraphClient();
          const device = await client.api(`/devices/${variables?.deviceId}`).get();
          
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
    );    // User compliance resource - optimized
    this.server.resource(
      'user_compliance',
      new ResourceTemplate('m365://users/{userId}/compliance', { list: undefined }),
      async (uri: URL, variables: any) => {
        try {
          const compliance = await this.makeGraphApiCall('/deviceManagement/managedDevices', {
            filter: `userId eq '${variables?.userId}'`,
            select: ['id', 'deviceName', 'operatingSystem', 'complianceState', 'lastSyncDateTime', 'enrolledDateTime']
          });
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(compliance, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `Error reading user compliance: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      }
    );    // Team governance resource - optimized with parallel calls
    this.server.resource(
      'team_governance',
      new ResourceTemplate('m365://teams/{teamId}/governance', { list: undefined }),
      async (uri: URL, variables: any) => {
        try {
          // Make parallel API calls for better performance
          const [team, members, channels] = await Promise.all([
            this.makeGraphApiCall(`/teams/${variables?.teamId}`, {
              select: ['id', 'displayName', 'memberSettings', 'guestSettings', 'funSettings', 'messagingSettings']
            }),
            this.makeGraphApiCall(`/teams/${variables?.teamId}/members`, {
              select: ['id', 'displayName', 'email', 'roles'],
              top: 100
            }),
            this.makeGraphApiCall(`/teams/${variables?.teamId}/channels`, {
              select: ['id', 'displayName', 'description', 'membershipType'],
              top: 50
            })
          ]);
            const governance = {
            team,
            members,
            channels,
            governance: {
              membershipType: (team as any)?.memberSettings?.allowAddRemoveApps,
              guestSettings: (team as any)?.guestSettings,
              funSettings: (team as any)?.funSettings,
              messagingSettings: (team as any)?.messagingSettings
            }
          };
          
          return {
            contents: [
              {
                uri: uri.href,
                mimeType: 'application/json',
                text: JSON.stringify(governance, null, 2),
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

    // Setup extended resources (40+ additional resources covering security, compliance, device management, and collaboration)
    setupExtendedResources(this.server, this.getGraphClient());
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
      resourceUri,      changeType,
      timestamp: new Date().toISOString()
    };
    
    this.broadcastUpdate(notification);
  }

  /**
   * Modern Graph API wrapper with retry logic and authentication on demand
   */
  private async makeGraphApiCall<T>(
    endpoint: string, 
    options: {
      method?: 'GET' | 'POST' | 'PATCH' | 'DELETE';
      body?: any;
      select?: string[];
      filter?: string;
      expand?: string;
      top?: number;
      maxRetries?: number;
    } = {}
  ): Promise<T> {
    const {
      method = 'GET',
      body,
      select,
      filter,
      expand,
      top,
      maxRetries = 3
    } = options;

    // Ensure authentication before making API call
    await this.ensureAuthenticated();

    const client = this.getGraphClient();
    let apiCall = client.api(endpoint);

    // Add query parameters for optimization
    if (select && select.length > 0) {
      apiCall = apiCall.select(select.join(','));
    }
    if (filter) {
      apiCall = apiCall.filter(filter);
    }
    if (expand) {
      apiCall = apiCall.expand(expand);
    }
    if (top) {
      apiCall = apiCall.top(top);
    }

    // Implement retry logic with exponential backoff
    for (let attempt = 0; attempt < maxRetries; attempt++) {
      try {
        let result: T;
        
        switch (method) {
          case 'GET':
            result = await apiCall.get();
            break;
          case 'POST':
            result = await apiCall.post(body);
            break;
          case 'PATCH':
            result = await apiCall.patch(body);
            break;
          case 'DELETE':
            result = await apiCall.delete();
            break;
          default:
            throw new Error(`Unsupported HTTP method: ${method}`);
        }

        // Log successful API call
        console.debug(`✅ Graph API Success: ${method} ${endpoint} (attempt ${attempt + 1})`);
        return result;

      } catch (error: any) {
        const isLastAttempt = attempt === maxRetries - 1;
        
        // Handle throttling (429) with retry-after
        if (error.status === 429 && !isLastAttempt) {
          const retryAfter = parseInt(error.headers?.['retry-after'] || '1');
          const backoffDelay = Math.min(retryAfter * 1000, Math.pow(2, attempt) * 1000);
          
          console.warn(`⚠️ Graph API Throttled: ${method} ${endpoint}, retrying in ${backoffDelay}ms (attempt ${attempt + 1})`);
          await new Promise(resolve => setTimeout(resolve, backoffDelay));
          continue;
        }
        
        // Handle other retryable errors (5xx)
        if (error.status >= 500 && error.status < 600 && !isLastAttempt) {
          const backoffDelay = Math.pow(2, attempt) * 1000;
          console.warn(`⚠️ Graph API Server Error: ${method} ${endpoint}, retrying in ${backoffDelay}ms (attempt ${attempt + 1})`);
          await new Promise(resolve => setTimeout(resolve, backoffDelay));
          continue;
        }        // Log and throw on final attempt or non-retryable errors
        console.error(`❌ Graph API Error: ${method} ${endpoint}`, {
          status: error.status,
          message: error.message,
          attempt: attempt + 1
        });
        
        throw new McpError(
          error.status >= 400 && error.status < 500 ? ErrorCode.InvalidParams : ErrorCode.InternalError,
          `Graph API ${method} ${endpoint} failed: ${error.message} (Status: ${error.status})`
        );
      }
    }    // This should never be reached due to the retry logic, but TypeScript requires it
    throw new McpError(ErrorCode.InternalError, `Graph API ${method} ${endpoint} failed after ${maxRetries} attempts`);
  }

  // ...existing code...
}

// Start the server
async function main() {
  try {
    const server = new M365CoreServer();
    
    // Tools and resources are already registered in constructor for Smithery discovery
    // Authentication will happen on-demand when tools are executed
    console.log('✅ Server initialized with tools registered and lazy authentication');
    
    const transport = process.env.NODE_ENV === 'http'
      ? new StreamableHTTPServerTransport({
          sessionIdGenerator: () => randomUUID() // Added sessionIdGenerator
        })
      : new StdioServerTransport();

    await server.server.connect(transport);
    console.log(`M365 Core MCP Server running on ${process.env.NODE_ENV === 'http' ? `http://localhost:${PORT}` : 'stdio'}`);
    console.log('🚀 Ready to serve requests with on-demand authentication');
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
