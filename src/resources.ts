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
  },

  // ===== DOCUMENT GENERATION RESOURCES =====
  {
    uri: 'm365://documents/presentations',
    name: 'PowerPoint Presentations',
    description: 'List of PowerPoint presentations in user\'s OneDrive',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const presentations = await graphClient
        .api('/me/drive/root/search(q=\'.pptx\')')
        .top(50)
        .select('id,name,createdDateTime,lastModifiedDateTime,size,webUrl')
        .get();
      return presentations;
    }
  },
  {
    uri: 'm365://documents/word-documents',
    name: 'Word Documents',
    description: 'List of Word documents in user\'s OneDrive',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const documents = await graphClient
        .api('/me/drive/root/search(q=\'.docx\')')
        .top(50)
        .select('id,name,createdDateTime,lastModifiedDateTime,size,webUrl')
        .get();
      return documents;
    }
  },
  {
    uri: 'm365://documents/reports',
    name: 'Generated Reports',
    description: 'List of generated reports and analysis documents',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const reports = await graphClient
        .api('/me/drive/root/children')
        .filter('startsWith(name,\'Report_\') or startsWith(name,\'Analysis_\')')
        .select('id,name,createdDateTime,lastModifiedDateTime,size,webUrl')
        .get()
        .catch(() => ({ value: [] }));
      return reports;
    }
  },
  {
    uri: 'm365://documents/templates',
    name: 'Document Templates',
    description: 'Available document templates for report generation',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      // Return metadata about available templates
      return {
        powerpoint: {
          professional: { id: 'professional', name: 'Professional Business', slides: ['title', 'agenda', 'content', 'conclusion'] },
          security: { id: 'security', name: 'Security Assessment', slides: ['cover', 'executive_summary', 'findings', 'recommendations', 'roadmap'] },
          compliance: { id: 'compliance', name: 'Compliance Report', slides: ['title', 'overview', 'status', 'gaps', 'remediation'] }
        },
        word: {
          professional: { id: 'professional', name: 'Professional Report', sections: ['cover', 'toc', 'executive_summary', 'analysis', 'recommendations', 'appendix'] },
          technical: { id: 'technical', name: 'Technical Documentation', sections: ['title', 'overview', 'architecture', 'configuration', 'troubleshooting'] },
          audit: { id: 'audit', name: 'Audit Report', sections: ['cover', 'scope', 'methodology', 'findings', 'conclusions', 'action_plan'] }
        },
        html: {
          dashboard: { id: 'dashboard', name: 'Interactive Dashboard', theme: 'modern', features: ['charts', 'filters', 'export'] },
          report: { id: 'report', name: 'Web Report', theme: 'professional', features: ['toc', 'sections', 'print_view'] }
        }
      };
    }
  },

  // ===== POLICY MANAGEMENT RESOURCES =====
  {
    uri: 'm365://policies/conditional-access',
    name: 'Conditional Access Policies',
    description: 'All conditional access policies and their current state',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const policies = await graphClient
        .api('/identity/conditionalAccess/policies')
        .select('id,displayName,state,conditions,grantControls,sessionControls')
        .get();
      return policies;
    }
  },
  {
    uri: 'm365://policies/retention',
    name: 'Retention Policies',
    description: 'Retention policies across Microsoft 365 workloads',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const policies = await graphClient
        .api('/security/informationProtection/retentionLabels')
        .get()
        .catch(() => ({ value: [] }));
      return policies;
    }
  },
  {
    uri: 'm365://policies/information-protection',
    name: 'Information Protection Policies',
    description: 'Azure Information Protection and sensitivity label policies',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const [labels, labelPolicies] = await Promise.all([
        graphClient.api('/security/informationProtection/sensitivityLabels').get().catch(() => ({ value: [] })),
        graphClient.api('/security/informationProtection/labelPolicies').get().catch(() => ({ value: [] }))
      ]);
      return { labels, labelPolicies };
    }
  },
  {
    uri: 'm365://policies/defender',
    name: 'Defender Policies',
    description: 'Microsoft Defender for Office 365 security policies',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      // Defender policies are managed through Security & Compliance Center
      const threatPolicies = await graphClient
        .api('/security/threatSubmission/emailThreats')
        .top(50)
        .get()
        .catch(() => ({ value: [] }));
      return threatPolicies;
    }
  },
  {
    uri: 'm365://policies/teams',
    name: 'Teams Policies',
    description: 'Microsoft Teams messaging, meeting, and calling policies',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const teams = await graphClient
        .api('/teams')
        .top(50)
        .select('id,displayName,description')
        .get()
        .catch(() => ({ value: [] }));
      return { teams, note: 'Teams policies are managed through Teams admin center APIs' };
    }
  },
  {
    uri: 'm365://policies/exchange',
    name: 'Exchange Policies',
    description: 'Exchange Online mail flow, retention, and mobile device policies',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      // Exchange policies information
      return {
        note: 'Exchange policies are managed through Exchange Online PowerShell',
        categories: [
          'Mail flow rules (Transport rules)',
          'Mobile device policies',
          'Retention policies',
          'Sharing policies',
          'Address book policies'
        ]
      };
    }
  },
  {
    uri: 'm365://policies/sharepoint-governance',
    name: 'SharePoint Governance Policies',
    description: 'SharePoint sharing, access control, and lifecycle policies',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const adminSettings = await graphClient
        .api('/admin/sharepoint/settings')
        .get()
        .catch(() => null);
      
      const sites = await graphClient
        .api('/sites?search=*')
        .top(10)
        .select('id,displayName,webUrl,sharingCapabilities')
        .get()
        .catch(() => ({ value: [] }));
      
      return { adminSettings, recentSites: sites };
    }
  },
  {
    uri: 'm365://policies/security-alerts',
    name: 'Security Alert Policies',
    description: 'Security alert policies and notification configurations',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      const alertPolicies = await graphClient
        .api('/security/alerts_v2')
        .filter('status eq \'new\' or status eq \'inProgress\'')
        .top(50)
        .select('id,title,category,severity,status,createdDateTime')
        .orderby('createdDateTime desc')
        .get();
      return alertPolicies;
    }
  },
  {
    uri: 'm365://policies/overview',
    name: 'Policy Overview',
    description: 'Summary of all policies across Microsoft 365',
    mimeType: 'application/json',
    handler: async (graphClient) => {
      // Gather overview of policies across different areas
      const [conditionalAccess, dlpPolicies, alerts] = await Promise.all([
        graphClient.api('/identity/conditionalAccess/policies').select('id,displayName,state').get().catch(() => ({ value: [] })),
        graphClient.api('/security/informationProtection/sensitivityLabels').get().catch(() => ({ value: [] })),
        graphClient.api('/security/alerts_v2').top(10).get().catch(() => ({ value: [] }))
      ]);
      
      return {
        summary: {
          conditionalAccessPolicies: conditionalAccess.value?.length || 0,
          sensitivityLabels: dlpPolicies.value?.length || 0,
          activeAlerts: alerts.value?.length || 0
        },
        lastUpdated: new Date().toISOString()
      };
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
