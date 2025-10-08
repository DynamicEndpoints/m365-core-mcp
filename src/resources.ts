import { Client } from '@microsoft/microsoft-graph-client';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';

/**
 * MCP Resources for Microsoft 365
 * These resources provide LLMs and users with quick access to commonly needed M365 data
 */

export interface ResourceHandler {
  uri: string;
  name: string;
  description: string;
  mimeType: string;
  handler: (graphClient: Client, params?: URLSearchParams) => Promise<any>;
}

export const m365Resources: ResourceHandler[] = [
  // ===== SECURITY RESOURCES =====
  {
    uri: 'm365://security/alerts',
    name: 'Security Alerts',
    description: 'Current security alerts from Microsoft Defender and security products',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const alerts = await graphClient
        .api('/security/alerts_v2')
        .top(50)
        .orderby('createdDateTime desc')
        .get();
      return alerts;
    }
  },
  {
    uri: 'm365://security/incidents',
    name: 'Security Incidents',
    description: 'Active security incidents requiring attention',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const incidents = await graphClient
        .api('/security/incidents')
        .top(50)
        .orderby('createdDateTime desc')
        .get();
      return incidents;
    }
  },
  {
    uri: 'm365://security/conditional-access',
    name: 'Conditional Access Policies',
    description: 'Conditional access policies and their configuration',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const policies = await graphClient
        .api('/identity/conditionalAccess/policies')
        .get();
      return policies;
    }
  },
  {
    uri: 'm365://security/identity-protection',
    name: 'Identity Protection',
    description: 'Identity protection risk detections and risky users',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const [riskDetections, riskyUsers] = await Promise.all([
        graphClient.api('/identityProtection/riskDetections').top(50).get(),
        graphClient.api('/identityProtection/riskyUsers').top(50).get()
      ]);
      return { riskDetections, riskyUsers };
    }
  },
  {
    uri: 'm365://security/threat-assessment',
    name: 'Threat Assessment',
    description: 'Threat assessment summary and recommendations',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const secureScore = await graphClient
        .api('/security/secureScores')
        .top(1)
        .get();
      return secureScore;
    }
  },

  // ===== COMPLIANCE RESOURCES =====
  {
    uri: 'm365://compliance/policies',
    name: 'Compliance Policies',
    description: 'DLP policies, retention policies, and compliance configurations',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      // Note: DLP policies require specific endpoints
      const retentionPolicies = await graphClient
        .api('/compliance/retentionPolicies')
        .get()
        .catch(() => ({ value: [] }));
      return { retentionPolicies };
    }
  },
  {
    uri: 'm365://compliance/audit-summary',
    name: 'Audit Log Summary',
    description: 'Recent audit log activity summary',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const auditLogs = await graphClient
        .api('/auditLogs/directoryAudits')
        .top(100)
        .orderby('activityDateTime desc')
        .get();
      return auditLogs;
    }
  },
  {
    uri: 'm365://compliance/sensitivity-labels',
    name: 'Sensitivity Labels',
    description: 'Information protection sensitivity labels and policies',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const labels = await graphClient
        .api('/informationProtection/policy/labels')
        .get()
        .catch(() => ({ value: [] }));
      return labels;
    }
  },

  // ===== DEVICE MANAGEMENT RESOURCES =====
  {
    uri: 'm365://devices/overview',
    name: 'Device Overview',
    description: 'Overview of all managed devices across platforms',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const devices = await graphClient
        .api('/deviceManagement/managedDevices')
        .top(100)
        .get();
      return devices;
    }
  },
  {
    uri: 'm365://devices/compliance',
    name: 'Device Compliance Status',
    description: 'Compliance status of managed devices',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const [compliance, policies] = await Promise.all([
        graphClient.api('/deviceManagement/managedDevices')
          .select('id,deviceName,complianceState,operatingSystem,lastSyncDateTime')
          .get(),
        graphClient.api('/deviceManagement/deviceCompliancePolicies').get()
      ]);
      return { devices: compliance, policies };
    }
  },
  {
    uri: 'm365://devices/intune-policies',
    name: 'Intune Policies',
    description: 'Device configuration and compliance policies',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const [configPolicies, compliancePolicies] = await Promise.all([
        graphClient.api('/deviceManagement/deviceConfigurations').get(),
        graphClient.api('/deviceManagement/deviceCompliancePolicies').get()
      ]);
      return { configPolicies, compliancePolicies };
    }
  },
  {
    uri: 'm365://devices/apps',
    name: 'Managed Applications',
    description: 'Applications managed through Intune',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const apps = await graphClient
        .api('/deviceManagement/mobileApps')
        .top(100)
        .get();
      return apps;
    }
  },

  // ===== USER & IDENTITY RESOURCES =====
  {
    uri: 'm365://users/directory',
    name: 'User Directory',
    description: 'All users in the organization with key attributes',
    mimeType: 'application/json',
    handler: async (graphClient, params) => {
      const top = params?.get('top') || '100';
      const users = await graphClient
        .api('/users')
        .select('id,displayName,userPrincipalName,mail,jobTitle,department,accountEnabled')
        .top(parseInt(top))
        .get();
      return users;
    }
  },
  {
    uri: 'm365://users/licenses',
    name: 'License Assignment',
    description: 'User license assignments and available licenses',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const [subscribedSkus, users] = await Promise.all([
        graphClient.api('/subscribedSkus').get(),
        graphClient.api('/users')
          .select('id,displayName,assignedLicenses')
          .filter('assignedLicenses/$count ne 0')
          .top(100)
          .get()
      ]);
      return { subscribedSkus, licensedUsers: users };
    }
  },
  {
    uri: 'm365://users/privileged',
    name: 'Privileged Users',
    description: 'Users with administrative roles',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const roleAssignments = await graphClient
        .api('/roleManagement/directory/roleAssignments')
        .expand('principal')
        .get();
      return roleAssignments;
    }
  },

  // ===== GROUP RESOURCES =====
  {
    uri: 'm365://groups/all',
    name: 'All Groups',
    description: 'Security groups, M365 groups, and distribution lists',
    mimeType: 'application/json',
    handler: async (graphClient, params) => {
      const top = params?.get('top') || '100';
      const groups = await graphClient
        .api('/groups')
        .select('id,displayName,groupTypes,mailEnabled,securityEnabled,mail,description')
        .top(parseInt(top))
        .get();
      return groups;
    }
  },
  {
    uri: 'm365://groups/teams',
    name: 'Microsoft Teams',
    description: 'All Microsoft Teams in the organization',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const teams = await graphClient
        .api('/groups')
        .filter('resourceProvisioningOptions/Any(x:x eq \'Team\')')
        .select('id,displayName,description,mail,visibility')
        .get();
      return teams;
    }
  },

  // ===== SHAREPOINT RESOURCES =====
  {
    uri: 'm365://sharepoint/sites',
    name: 'SharePoint Sites',
    description: 'SharePoint sites in the tenant',
    mimeType: 'application/json',
    handler: async (graphClient, params) => {
      const top = params?.get('top') || '50';
      const sites = await graphClient
        .api('/sites')
        .select('id,displayName,webUrl,description,createdDateTime')
        .top(parseInt(top))
        .get();
      return sites;
    }
  },
  {
    uri: 'm365://sharepoint/permissions',
    name: 'SharePoint Permissions',
    description: 'External sharing and permission summary',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const sites = await graphClient
        .api('/sites')
        .select('id,displayName')
        .top(20)
        .get();
      
      // Get sharing links for each site
      const sharingData = await Promise.all(
        sites.value.map(async (site: any) => {
          try {
            const permissions = await graphClient
              .api(`/sites/${site.id}/permissions`)
              .get();
            return { site: site.displayName, permissions: permissions.value };
          } catch {
            return { site: site.displayName, permissions: [] };
          }
        })
      );
      
      return sharingData;
    }
  },

  // ===== EXCHANGE RESOURCES =====
  {
    uri: 'm365://exchange/mailboxes',
    name: 'Exchange Mailboxes',
    description: 'Mailbox configurations and settings',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const users = await graphClient
        .api('/users')
        .select('id,displayName,mail,mailboxSettings')
        .filter('mail ne null')
        .top(100)
        .get();
      return users;
    }
  },
  {
    uri: 'm365://exchange/transport-rules',
    name: 'Exchange Transport Rules',
    description: 'Mail flow rules and configurations',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      // Note: Transport rules require Exchange Online PowerShell or specific Graph endpoints
      return { 
        message: 'Transport rules require specific Exchange Online access',
        recommendation: 'Use manage_exchange_settings tool for detailed configuration'
      };
    }
  },

  // ===== COLLABORATION RESOURCES =====
  {
    uri: 'm365://collaboration/teams-activity',
    name: 'Teams Activity',
    description: 'Microsoft Teams usage and activity statistics',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const reports = await graphClient
        .api('/reports/getTeamsUserActivityUserDetail(period=\'D7\')')
        .get()
        .catch(() => ({ value: 'Reports require additional permissions' }));
      return reports;
    }
  },
  {
    uri: 'm365://collaboration/onedrive-usage',
    name: 'OneDrive Usage',
    description: 'OneDrive storage and usage statistics',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const reports = await graphClient
        .api('/reports/getOneDriveUsageAccountDetail(period=\'D7\')')
        .get()
        .catch(() => ({ value: 'Reports require additional permissions' }));
      return reports;
    }
  },

  // ===== APPLICATION RESOURCES =====
  {
    uri: 'm365://applications/registrations',
    name: 'Application Registrations',
    description: 'Azure AD application registrations',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const apps = await graphClient
        .api('/applications')
        .select('id,appId,displayName,signInAudience,createdDateTime')
        .top(100)
        .get();
      return apps;
    }
  },
  {
    uri: 'm365://applications/service-principals',
    name: 'Service Principals',
    description: 'Enterprise applications and service principals',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const sps = await graphClient
        .api('/servicePrincipals')
        .select('id,appId,displayName,servicePrincipalType,accountEnabled')
        .top(100)
        .get();
      return sps;
    }
  },
  {
    uri: 'm365://applications/consent',
    name: 'Application Consent',
    description: 'OAuth2 permissions and consent grants',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const grants = await graphClient
        .api('/oauth2PermissionGrants')
        .top(100)
        .get();
      return grants;
    }
  },

  // ===== GOVERNANCE RESOURCES =====
  {
    uri: 'm365://governance/access-reviews',
    name: 'Access Reviews',
    description: 'Access review campaigns and status',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const reviews = await graphClient
        .api('/identityGovernance/accessReviews/definitions')
        .get()
        .catch(() => ({ value: [] }));
      return reviews;
    }
  },
  {
    uri: 'm365://governance/entitlement',
    name: 'Entitlement Management',
    description: 'Access packages and entitlement policies',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const packages = await graphClient
        .api('/identityGovernance/entitlementManagement/accessPackages')
        .get()
        .catch(() => ({ value: [] }));
      return packages;
    }
  },

  // ===== TENANT INFORMATION =====
  {
    uri: 'm365://tenant/organization',
    name: 'Organization Information',
    description: 'Tenant configuration and organization details',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const org = await graphClient
        .api('/organization')
        .get();
      return org;
    }
  },
  {
    uri: 'm365://tenant/domains',
    name: 'Verified Domains',
    description: 'Verified domains in the tenant',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const domains = await graphClient
        .api('/domains')
        .get();
      return domains;
    }
  },
  {
    uri: 'm365://tenant/subscriptions',
    name: 'Subscriptions',
    description: 'Active subscriptions and SKUs',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const skus = await graphClient
        .api('/subscribedSkus')
        .get();
      return skus;
    }
  }
];

/**
 * Get resource by URI or URI pattern
 */
export function getResourceByUri(uri: string): ResourceHandler | undefined {
  // Direct match
  const direct = m365Resources.find(r => r.uri === uri);
  if (direct) return direct;

  // Pattern match (e.g., m365://users/directory?top=50)
  const uriWithoutParams = uri.split('?')[0];
  return m365Resources.find(r => r.uri === uriWithoutParams);
}

/**
 * List all available resources with descriptions
 */
export function listResources(): Array<{ uri: string; name: string; description: string; mimeType: string }> {
  return m365Resources.map(r => ({
    uri: r.uri,
    name: r.name,
    description: r.description,
    mimeType: r.mimeType
  }));
}
