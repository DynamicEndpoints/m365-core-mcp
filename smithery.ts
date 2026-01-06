/**
 * Smithery TypeScript configuration for M365 Core MCP Server
 * https://smithery.ai/docs/build/deployments#typescript-deploy
 * 
 * Implements latest MCP SDK and Smithery authentication patterns:
 * - OAuth provider export for automatic endpoint mounting
 * - AuthInfo injection into createServer
 * - Type-safe configuration with Zod
 */

import { z } from 'zod';
import type { AuthInfo } from '@modelcontextprotocol/sdk/server/auth/types.js';

// Configuration schema for Smithery
export const configSchema = z.object({
  msTenantId: z.string().describe('Microsoft Tenant ID for authentication'),
  msClientId: z.string().describe('Microsoft Client ID for authentication'),
  msClientSecret: z.string().describe('Microsoft Client Secret for authentication'),
  useHttp: z.boolean().optional().describe('Use HTTP transport instead of stdio (default true for Smithery)'),
  stateless: z.boolean().optional().describe('Use stateless HTTP mode (default false)'),
  port: z.number().optional().describe('Port for HTTP server (default 8080)'),
  logLevel: z.enum(['debug', 'info', 'warn', 'error']).optional().describe('Log level (default info)')
});

// Tool definitions for Smithery discovery
export const tools = [
  // Core M365 Management Tools
  {
    name: 'manage_distribution_lists',
    description: 'Manage Microsoft 365 distribution lists - create, update, delete, and manage membership',
    category: 'Microsoft 365',
    tags: ['email', 'groups', 'distribution-lists', 'm365'],
    inputSchema: {
      type: 'object',
      properties: {
        action: { type: 'string', enum: ['create', 'update', 'delete', 'list', 'add_members', 'remove_members'] },
        name: { type: 'string', description: 'Distribution list name' },
        description: { type: 'string', description: 'Distribution list description' },
        members: { type: 'array', items: { type: 'string' }, description: 'Member email addresses' }
      },
      required: ['action']
    }
  },
  {
    name: 'manage_security_groups',
    description: 'Manage Azure AD security groups - create, update, delete, and manage membership',
    category: 'Azure AD',
    tags: ['security', 'groups', 'azure-ad', 'access-control'],
    inputSchema: {
      type: 'object',
      properties: {
        action: { type: 'string', enum: ['create', 'update', 'delete', 'list', 'add_members', 'remove_members'] },
        name: { type: 'string', description: 'Security group name' },
        description: { type: 'string', description: 'Security group description' },
        members: { type: 'array', items: { type: 'string' }, description: 'Member IDs or email addresses' }
      },
      required: ['action']
    }
  },
  {
    name: 'manage_m365_groups',
    description: 'Manage Microsoft 365 Groups - create, update, delete, and manage membership',
    category: 'Microsoft 365',
    tags: ['groups', 'collaboration', 'teams', 'm365'],
    inputSchema: {
      type: 'object',
      properties: {
        action: { type: 'string', enum: ['create', 'update', 'delete', 'list', 'add_members', 'remove_members'] },
        name: { type: 'string', description: 'M365 group name' },
        description: { type: 'string', description: 'M365 group description' },
        members: { type: 'array', items: { type: 'string' }, description: 'Member email addresses' }
      },
      required: ['action']    }
  },
  {
    name: 'manage_exchange_settings',
    description: 'Manage Exchange Online settings - mailbox configuration, transport rules, and policies',
    category: 'Exchange Online',
    tags: ['exchange', 'email', 'mailbox', 'transport-rules'],
    inputSchema: {
      type: 'object',
      properties: {
        action: { type: 'string', enum: ['create_mailbox', 'update_mailbox', 'delete_mailbox', 'list_mailboxes', 'create_transport_rule', 'list_transport_rules'] },
        userPrincipalName: { type: 'string', description: 'User principal name for mailbox operations' },
        displayName: { type: 'string', description: 'Display name' },
        ruleName: { type: 'string', description: 'Transport rule name' }
      },
      required: ['action']
    }
  },
  {
    name: 'manage_user_settings',
    description: 'Manage user settings and configurations across Microsoft 365 services',
    category: 'User Management',
    tags: ['users', 'settings', 'configuration', 'm365'],
    inputSchema: {
      type: 'object',
      properties: {
        action: { type: 'string', enum: ['create', 'update', 'delete', 'list', 'get'] },
        userPrincipalName: { type: 'string', description: 'User principal name' },
        displayName: { type: 'string', description: 'User display name' },
        department: { type: 'string', description: 'User department' }
      },
      required: ['action']
    }
  },
  {
    name: 'dynamicendpoints m365 assistant',
    description: 'Make direct calls to Microsoft Graph API or Azure Management API with authentication',
    category: 'API Access',
    tags: ['api', 'graph', 'azure', 'direct-access'],
    inputSchema: {
      type: 'object',
      properties: {
        endpoint: { type: 'string', description: 'API endpoint URL' },
        method: { type: 'string', enum: ['GET', 'POST', 'PUT', 'PATCH', 'DELETE'], description: 'HTTP method' },        body: { type: 'object', description: 'Request body for POST/PUT/PATCH requests' },
        apiType: { type: 'string', enum: ['graph', 'azure'], description: 'API type (graph or azure)' }
      },
      required: ['endpoint', 'method', 'apiType']
    }
  },
  {
    name: 'manage_offboarding',
    description: 'Automate user offboarding processes - disable accounts, backup data, transfer ownership',
    category: 'User Management',
    tags: ['offboarding', 'security', 'data-transfer', 'lifecycle'],
    inputSchema: {
      type: 'object',
      properties: {
        userPrincipalName: { type: 'string', description: 'User principal name to offboard' },
        transferMailboxTo: { type: 'string', description: 'User to transfer mailbox ownership to' },
        transferOneDriveTo: { type: 'string', description: 'User to transfer OneDrive ownership to' },
        disableAccount: { type: 'boolean', description: 'Whether to disable the user account' }
      },
      required: ['userPrincipalName']
    }
  },
  {
    name: 'manage_sharepoint_sites',
    description: 'Manage SharePoint sites - create, configure, and manage site settings and permissions',
    category: 'SharePoint',
    tags: ['sharepoint', 'sites', 'collaboration', 'content-management'],
    inputSchema: {
      type: 'object',
      properties: {
        action: { type: 'string', enum: ['create', 'update', 'delete', 'list', 'get'] },
        title: { type: 'string', description: 'Site title' },
        url: { type: 'string', description: 'Site URL' },
        template: { type: 'string', description: 'Site template' }
      },
      required: ['action']
    }
  },
  {
    name: 'manage_sharepoint_lists',
    description: 'Manage SharePoint lists and libraries - create, configure, and manage list items',
    category: 'SharePoint',
    tags: ['sharepoint', 'lists', 'data-management', 'content'],
    inputSchema: {
      type: 'object',
      properties: {
        action: { type: 'string', enum: ['create', 'update', 'delete', 'list', 'get'] },
        siteId: { type: 'string', description: 'SharePoint site ID' },
        title: { type: 'string', description: 'List title' },
        description: { type: 'string', description: 'List description' }
      },
      required: ['action']
    }
  },  {
    name: 'manage_azuread_roles',
    description: 'Manage Azure AD directory roles - assign, remove, and list role assignments',
    category: 'Azure AD',
    tags: ['roles', 'permissions', 'rbac', 'azure-ad'],
    inputSchema: {
      type: 'object',
      properties: {
        action: { type: 'string', enum: ['assign', 'remove', 'list'] },
        roleName: { type: 'string', description: 'Directory role name' },
        userPrincipalName: { type: 'string', description: 'User to assign/remove role' }
      },
      required: ['action']
    }
  },
  {
    name: 'manage_azuread_apps',
    description: 'Manage Azure AD application registrations - create, update, and manage app permissions',
    category: 'Azure AD',
    tags: ['applications', 'app-registrations', 'azure-ad', 'oauth'],
    inputSchema: {
      type: 'object',
      properties: {
        action: { type: 'string', enum: ['create', 'update', 'delete', 'list', 'get'] },
        displayName: { type: 'string', description: 'Application display name' },
        appId: { type: 'string', description: 'Application ID' }
      },
      required: ['action']
    }
  },
  {
    name: 'manage_azuread_devices',
    description: 'Manage Azure AD registered devices - enable, disable, delete, and monitor devices',
    category: 'Azure AD',
    tags: ['devices', 'device-management', 'azure-ad', 'security'],
    inputSchema: {
      type: 'object',
      properties: {
        action: { type: 'string', enum: ['enable', 'disable', 'delete', 'list', 'get'] },
        deviceId: { type: 'string', description: 'Device ID' },
        displayName: { type: 'string', description: 'Device display name' }
      },
      required: ['action']
    }
  },
  {
    name: 'manage_service_principals',
    description: 'Manage Azure AD service principals - create, update, and manage service principal permissions',
    category: 'Azure AD',
    tags: ['service-principals', 'applications', 'azure-ad', 'automation'],
    inputSchema: {
      type: 'object',
      properties: {
        action: { type: 'string', enum: ['create', 'update', 'delete', 'list', 'get'] },
        appId: { type: 'string', description: 'Application ID' },
        displayName: { type: 'string', description: 'Service principal display name' }
      },
      required: ['action']
    }
  },
  {
    name: 'search_audit_log',
    description: 'Search and analyze Azure AD unified audit logs for compliance and security monitoring',
    category: 'Security & Compliance',
    tags: ['audit-logs', 'compliance', 'security', 'monitoring'],
    inputSchema: {
      type: 'object',
      properties: {
        startDate: { type: 'string', description: 'Start date for audit log search' },
        endDate: { type: 'string', description: 'End date for audit log search' },
        operations: { type: 'array', items: { type: 'string' }, description: 'Operations to search for' }
      },
      required: ['startDate', 'endDate']
    }
  },
  {
    name: 'manage_alerts',
    description: 'Manage and respond to Microsoft security alerts from various security products',
    category: 'Security & Compliance',
    tags: ['security-alerts', 'incident-response', 'monitoring', 'threat-detection'],
    inputSchema: {
      type: 'object',
      properties: {
        action: { type: 'string', enum: ['list', 'get', 'update', 'dismiss'] },
        alertId: { type: 'string', description: 'Alert ID' },
        status: { type: 'string', description: 'Alert status' }
      },
      required: ['action']
    }
  },

  // DLP and Information Protection Tools
  {
    name: 'manage_dlp_policies',
    description: 'Manage Data Loss Prevention policies across Microsoft 365 services',
    category: 'Security & Compliance',
    tags: ['dlp', 'data-protection', 'compliance', 'security']
  },
  {
    name: 'manage_dlp_incidents',
    description: 'Manage and investigate DLP policy incidents and violations',
    category: 'Security & Compliance',
    tags: ['dlp', 'incidents', 'investigation', 'compliance']
  },
  {
    name: 'manage_sensitivity_labels',
    description: 'Manage sensitivity labels for information protection and classification',
    category: 'Security & Compliance',
    tags: ['sensitivity-labels', 'information-protection', 'classification', 'compliance']
  },

  // Intune macOS Management Tools
  {
    name: 'manage_intune_macos_devices',
    description: 'Manage macOS devices in Microsoft Intune - enrollment, compliance, and device actions',
    category: 'Device Management',
    tags: ['intune', 'macos', 'device-management', 'mdm']
  },
  {
    name: 'manage_intune_macos_policies',
    description: 'Manage macOS configuration and compliance policies in Intune',
    category: 'Device Management',
    tags: ['intune', 'macos', 'policies', 'configuration', 'compliance']
  },
  {
    name: 'manage_intune_macos_apps',
    description: 'Manage macOS application deployment and management through Intune',
    category: 'Device Management',
    tags: ['intune', 'macos', 'applications', 'app-deployment']
  },
  {
    name: 'assess_intune_macos_compliance',
    description: 'Assess and report on macOS device compliance with organizational policies',
    category: 'Device Management',
    tags: ['intune', 'macos', 'compliance', 'assessment', 'reporting']
  },

  // Compliance Framework Tools
  {
    name: 'manage_compliance_frameworks',
    description: 'Manage compliance frameworks (HITRUST, ISO27001, SOC2, CIS) and their configurations',
    category: 'Security & Compliance',
    tags: ['compliance', 'frameworks', 'governance', 'risk-management']
  },
  {
    name: 'run_compliance_assessments',
    description: 'Execute compliance assessments against various frameworks and standards',
    category: 'Security & Compliance',
    tags: ['compliance', 'assessment', 'auditing', 'frameworks']
  },
  {
    name: 'monitor_compliance_status',
    description: 'Monitor ongoing compliance status and receive alerts for compliance drift',
    category: 'Security & Compliance',
    tags: ['compliance', 'monitoring', 'alerts', 'status']
  },
  {
    name: 'collect_evidence',
    description: 'Automated evidence collection for compliance and audit purposes',
    category: 'Security & Compliance',
    tags: ['evidence', 'compliance', 'auditing', 'documentation']
  },
  {
    name: 'analyze_compliance_gaps',
    description: 'Perform gap analysis against compliance frameworks and standards',
    category: 'Security & Compliance',
    tags: ['gap-analysis', 'compliance', 'assessment', 'frameworks']
  },
  {
    name: 'generate_audit_reports',
    description: 'Generate comprehensive audit reports for compliance and security assessments',
    category: 'Security & Compliance',
    tags: ['audit-reports', 'compliance', 'documentation', 'reporting']
  },
  {
    name: 'assess_cis_compliance',
    description: 'Assess CIS (Center for Internet Security) benchmark compliance across systems',
    category: 'Security & Compliance',
    tags: ['cis', 'benchmarks', 'security', 'compliance', 'assessment']
  }
];

// Resource definitions for Smithery discovery
export const resources = [
  {
    name: 'current_user',
    description: 'Information about the currently authenticated user',
    uri: 'm365://user/me'
  },
  {
    name: 'tenant_info',
    description: 'Microsoft 365 tenant information and configuration',
    uri: 'm365://tenant/info'
  },
  {
    name: 'sharepoint_sites',
    description: 'List of SharePoint sites in the organization',
    uri: 'm365://sharepoint/sites'
  },
  {
    name: 'sharepoint_admin_settings',
    description: 'SharePoint admin settings and configuration',
    uri: 'm365://sharepoint/admin/settings'
  },
  {
    name: 'user_info',
    description: 'Detailed information about a specific user',
    uri: 'm365://users/{userId}'
  },
  {
    name: 'group_info',
    description: 'Information about a specific group',
    uri: 'm365://groups/{groupId}'
  },
  {
    name: 'device_info',
    description: 'Information about a specific device',
    uri: 'm365://devices/{deviceId}'
  }
];

/**
 * Server factory function for Smithery - MUST be default export
 * 
 * Per MCP SDK and Smithery best practices:
 * - Receives config object with validated configuration
 * - Can optionally receive auth object with AuthInfo
 * - Sets up environment and returns server instance
 */
export default async function createServer({ 
  config, 
  auth 
}: { 
  config: z.infer<typeof configSchema>;
  auth?: AuthInfo;
}) {
  // Set environment variables from Smithery config
  if (config.msTenantId) process.env.MS_TENANT_ID = config.msTenantId;
  if (config.msClientId) process.env.MS_CLIENT_ID = config.msClientId;
  if (config.msClientSecret) process.env.MS_CLIENT_SECRET = config.msClientSecret;
  if (config.useHttp !== undefined) process.env.USE_HTTP = config.useHttp.toString();
  if (config.stateless !== undefined) process.env.STATELESS = config.stateless.toString();
  if (config.port) process.env.PORT = config.port.toString();
  if (config.logLevel) process.env.LOG_LEVEL = config.logLevel;

  // Log auth info if provided (useful for debugging)
  if (auth) {
    console.log(`Authenticated user: ${auth.clientId || 'unknown'}`);
  }

  // Dynamic import at runtime to avoid compilation issues
  const module = await import('./src/server.js');
  const { M365CoreServer } = module;
  const server = new M365CoreServer();
  
  // Return the underlying server for Smithery
  return server.server;
}

/**
 * OAuth Provider Export for Smithery CLI
 * 
 * When exported, Smithery CLI automatically mounts OAuth endpoints:
 * - GET /oauth/authorize - Redirect to authorization
 * - POST /oauth/token - Token exchange
 * - GET /.well-known/oauth-authorization-server - Metadata
 * 
 * @see https://smithery.ai/docs/build/deployments/typescript
 */
export { M365OAuthProvider as oauth } from './src/auth/oauth-provider.js';
