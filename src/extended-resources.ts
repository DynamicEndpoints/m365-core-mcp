import { McpServer, ResourceTemplate } from '@modelcontextprotocol/sdk/server/mcp.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { z } from 'zod';

/**
 * Extended resources and prompts for M365 Core MCP Server
 * This module provides additional 40 resources and comprehensive prompts
 */

export function setupExtendedResources(server: McpServer, graphClient: Client): void {
  // Additional M365 Resources (Security, Compliance, Intune, etc.)
  
  server.resource(
    "security_alerts",
    "m365://security/alerts",
    async (uri: URL) => {
      try {
        const alerts = await graphClient
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
  
  server.resource(
    "security_incidents",
    "m365://security/incidents",
    async (uri: URL) => {
      try {
        const incidents = await graphClient
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
  
  server.resource(
    "conditional_access_policies",
    "m365://identity/conditionalAccess/policies",
    async (uri: URL) => {
      try {
        const policies = await graphClient
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
  
  server.resource(
    "applications",
    "m365://applications",
    async (uri: URL) => {
      try {
        const applications = await graphClient
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
  
  server.resource(
    "service_principals",
    "m365://servicePrincipals",
    async (uri: URL) => {
      try {
        const servicePrincipals = await graphClient
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
  
  server.resource(
    "directory_roles",
    "m365://directoryRoles",
    async (uri: URL) => {
      try {
        const roles = await graphClient
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
  
  server.resource(
    "privileged_access",
    "m365://privilegedAccess/azureAD/resources",
    async (uri: URL) => {
      try {
        const resources = await graphClient
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
  
  server.resource(
    "audit_logs_signin_extended",
    "m365://auditLogs/signIns/extended",
    async (uri: URL) => {
      try {
        const signIns = await graphClient
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
  
  server.resource(
    "audit_logs_directory_extended",
    "m365://auditLogs/directoryAudits/extended",
    async (uri: URL) => {
      try {
        const directoryAudits = await graphClient
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
  
  server.resource(
    "intune_devices_extended",
    "m365://deviceManagement/managedDevices/extended",
    async (uri: URL) => {
      try {
        const devices = await graphClient
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
  
  server.resource(
    "intune_apps_extended",
    "m365://deviceAppManagement/mobileApps/extended",
    async (uri: URL) => {
      try {
        const apps = await graphClient
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
  
  server.resource(
    "intune_compliance_policies_extended",
    "m365://deviceManagement/deviceCompliancePolicies/extended",
    async (uri: URL) => {
      try {
        const policies = await graphClient
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
  
  server.resource(
    "intune_configuration_policies_extended",
    "m365://deviceManagement/deviceConfigurations/extended",
    async (uri: URL) => {
      try {
        const configurations = await graphClient
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
  
  server.resource(
    "teams_list_extended",
    "m365://teams/extended",
    async (uri: URL) => {
      try {
        const teams = await graphClient
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
  
  server.resource(
    "mail_folders_extended",
    "m365://me/mailFolders/extended",
    async (uri: URL) => {
      try {
        const mailFolders = await graphClient
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
  
  server.resource(
    "calendar_events_extended",
    "m365://me/events/extended",
    async (uri: URL) => {
      try {
        const events = await graphClient
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
  
  server.resource(
    "onedrive_extended",
    "m365://me/drive/extended",
    async (uri: URL) => {
      try {
        const drive = await graphClient
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
  
  server.resource(
    "planner_plans_extended",
    "m365://planner/plans/extended",
    async (uri: URL) => {
      try {
        const plans = await graphClient
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
  
  server.resource(
    "information_protection_extended",
    "m365://informationProtection/policy/labels/extended",
    async (uri: URL) => {
      try {
        const labels = await graphClient
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
  
  server.resource(
    "risky_users_extended",
    "m365://identityProtection/riskyUsers/extended",
    async (uri: URL) => {
      try {
        const riskyUsers = await graphClient
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
  
  server.resource(
    "threat_assessment_extended",
    "m365://informationProtection/threatAssessmentRequests/extended",
    async (uri: URL) => {
      try {
        const requests = await graphClient
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
  
  // Dynamic resources with parameters (21-40)
  
  server.resource(
    "user_messages_extended",
    new ResourceTemplate("m365://users/{userId}/messages/extended", { list: undefined }),
    async (uri: URL, variables: any) => {
      try {
        const messages = await graphClient
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
  
  server.resource(
    "user_calendar_extended",
    new ResourceTemplate("m365://users/{userId}/calendar/extended", { list: undefined }),
    async (uri: URL, variables: any) => {
      try {
        const calendar = await graphClient
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
  
  server.resource(
    "user_drive_extended",
    new ResourceTemplate("m365://users/{userId}/drive/extended", { list: undefined }),
    async (uri: URL, variables: any) => {
      try {
        const drive = await graphClient
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
  
  server.resource(
    "team_channels_extended",
    new ResourceTemplate("m365://teams/{teamId}/channels/extended", { list: undefined }),
    async (uri: URL, variables: any) => {
      try {
        const channels = await graphClient
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
  
  server.resource(
    "team_members_extended",
    new ResourceTemplate("m365://teams/{teamId}/members/extended", { list: undefined }),
    async (uri: URL, variables: any) => {
      try {
        const members = await graphClient
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
  
  server.resource(
    "device_info_extended",
    new ResourceTemplate("m365://deviceManagement/managedDevices/{deviceId}/extended", { list: undefined }),
    async (uri: URL, variables: any) => {
      try {
        const device = await graphClient
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
  
  server.resource(
    "app_assignments_extended",
    new ResourceTemplate("m365://deviceAppManagement/mobileApps/{appId}/assignments/extended", { list: undefined }),
    async (uri: URL, variables: any) => {
      try {
        const assignments = await graphClient
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
  
  server.resource(
    "policy_assignments_extended",
    new ResourceTemplate("m365://deviceManagement/deviceCompliancePolicies/{policyId}/assignments/extended", { list: undefined }),
    async (uri: URL, variables: any) => {
      try {
        const assignments = await graphClient
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
  
  server.resource(
    "group_members_extended",
    new ResourceTemplate("m365://groups/{groupId}/members/extended", { list: undefined }),
    async (uri: URL, variables: any) => {
      try {
        const members = await graphClient
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
  
  server.resource(
    "group_owners_extended",
    new ResourceTemplate("m365://groups/{groupId}/owners/extended", { list: undefined }),
    async (uri: URL, variables: any) => {
      try {
        const owners = await graphClient
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
  
  server.resource(
    "user_licenses_extended",
    new ResourceTemplate("m365://users/{userId}/licenseDetails/extended", { list: undefined }),
    async (uri: URL, variables: any) => {
      try {
        const licenses = await graphClient
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
  
  server.resource(
    "user_groups_extended",
    new ResourceTemplate("m365://users/{userId}/memberOf/extended", { list: undefined }),
    async (uri: URL, variables: any) => {
      try {
        const groups = await graphClient
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
  
  server.resource(
    "security_score_extended",
    "m365://security/secureScores/extended",
    async (uri: URL) => {
      try {
        const secureScores = await graphClient
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
  
  server.resource(
    "compliance_policies_dlp_extended",
    "m365://security/informationProtection/dlpPolicies/extended",
    async (uri: URL) => {
      try {
        const dlpPolicies = await graphClient
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
  
  server.resource(
    "retention_policies_extended",
    "m365://security/labels/retentionLabels/extended",
    async (uri: URL) => {
      try {
        const retentionLabels = await graphClient
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
  
  server.resource(
    "sensitivity_labels_extended",
    "m365://security/informationProtection/sensitivityLabels/extended",
    async (uri: URL) => {
      try {
        const sensitivityLabels = await graphClient
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
  
  server.resource(
    "communication_compliance_extended",
    "m365://compliance/communicationCompliance/policies/extended",
    async (uri: URL) => {
      try {
        const policies = await graphClient
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
  
  server.resource(
    "ediscovery_cases_extended",
    "m365://compliance/ediscovery/cases/extended",
    async (uri: URL) => {
      try {
        const cases = await graphClient
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
  
  server.resource(
    "subscribed_skus_extended",
    "m365://subscribedSkus/extended",
    async (uri: URL) => {
      try {
        const skus = await graphClient
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
  );
  
  server.resource(
    "directory_role_assignments",
    new ResourceTemplate("m365://directoryRoles/{roleId}/members", { list: undefined }),
    async (uri: URL, variables: any) => {
      try {
        const members = await graphClient
          .api(`/directoryRoles/${variables.roleId}/members`)
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
}

export function setupExtendedPrompts(server: McpServer, graphClient: Client): void {  // Security Analysis Prompts
  server.prompt(
    "security_assessment",
    "Analyze M365 security posture and provide recommendations",
    {
      scope: z.string().optional().describe("Security assessment scope (identity, data, devices, applications)"),
      timeframe: z.string().optional().describe("Assessment timeframe (last 7 days, 30 days, 90 days)"),
    },
    async (args: any) => {
      const scope = args.scope || "all";
      const timeframe = args.timeframe || "30 days";
      
      try {
        // Gather security data
        const [alerts, riskyUsers, conditionalAccessPolicies, secureScores] = await Promise.all([
          graphClient.api('/security/alerts_v2').top(10).get(),
          graphClient.api('/identityProtection/riskyUsers').get(),
          graphClient.api('/identity/conditionalAccess/policies').get(),
          graphClient.api('/security/secureScores').top(5).get(),
        ]);
        
        const analysisPrompt = `# Microsoft 365 Security Assessment

## Scope: ${scope}
## Timeframe: ${timeframe}

### Current Security Data:

**Recent Security Alerts:**
${JSON.stringify(alerts, null, 2)}

**Risky Users:**
${JSON.stringify(riskyUsers, null, 2)}

**Conditional Access Policies:**
${JSON.stringify(conditionalAccessPolicies, null, 2)}

**Secure Score History:**
${JSON.stringify(secureScores, null, 2)}

### Analysis Request:
Please analyze the above security data and provide:
1. Overall security posture assessment
2. Critical risks and vulnerabilities identified
3. Prioritized recommendations for improvement
4. Compliance gaps (if any)
5. Implementation roadmap for recommended changes

Focus on practical, actionable insights that can improve the organization's security stance.`;

        return {
          description: `Security assessment for ${scope} scope over ${timeframe}`,
          messages: [
            {
              role: "user",
              content: {
                type: "text",
                text: analysisPrompt,
              },
            },
          ],
        };
      } catch (error) {
        throw new McpError(
          ErrorCode.InternalError,
          `Error generating security assessment: ${error instanceof Error ? error.message : 'Unknown error'}`
        );
      }
    }
  );
    // Compliance Monitoring Prompt
  server.prompt(
    "compliance_review",
    "Generate compliance review and gap analysis",
    {
      framework: z.string().optional().describe("Compliance framework (SOC2, ISO27001, NIST, GDPR, HIPAA)"),
      scope: z.string().optional().describe("Review scope (policies, controls, data protection)"),
    },
    async (args: any) => {
      const framework = args.framework || "NIST";
      const scope = args.scope || "all";
      
      try {
        // Gather compliance-related data
        const [dlpPolicies, retentionLabels, sensitivityLabels, auditLogs] = await Promise.all([
          graphClient.api('/security/informationProtection/dlpPolicies').get(),
          graphClient.api('/security/labels/retentionLabels').get(),
          graphClient.api('/security/informationProtection/sensitivityLabels').get(),
          graphClient.api('/auditLogs/directoryAudits').top(20).get(),
        ]);
        
        const compliancePrompt = `# Microsoft 365 Compliance Review

## Framework: ${framework}
## Scope: ${scope}

### Current Compliance Configuration:

**Data Loss Prevention Policies:**
${JSON.stringify(dlpPolicies, null, 2)}

**Retention Labels:**
${JSON.stringify(retentionLabels, null, 2)}

**Sensitivity Labels:**
${JSON.stringify(sensitivityLabels, null, 2)}

**Recent Audit Events:**
${JSON.stringify(auditLogs, null, 2)}

### Compliance Review Request:
Please review the above configuration against ${framework} requirements and provide:
1. Current compliance status assessment
2. Identified gaps and non-conformities
3. Risk assessment for each gap
4. Remediation recommendations with priority levels
5. Policy and procedure updates needed
6. Training and awareness requirements

Provide specific, actionable recommendations to achieve and maintain compliance.`;

        return {
          description: `Compliance review for ${framework} framework covering ${scope}`,
          messages: [
            {
              role: "user",
              content: {
                type: "text",
                text: compliancePrompt,
              },
            },
          ],
        };
      } catch (error) {
        throw new McpError(
          ErrorCode.InternalError,
          `Error generating compliance review: ${error instanceof Error ? error.message : 'Unknown error'}`
        );
      }
    }
  );
    // User Access Review Prompt
  server.prompt(
    "user_access_review",
    "Analyze user access rights and permissions",
    {
      userId: z.string().optional().describe("Specific user ID to review (optional - if not provided, reviews all users)"),
      focus: z.string().optional().describe("Review focus (permissions, licenses, group memberships, recent activity)"),
    },
    async (args: any) => {
      const userId = args.userId;
      const focus = args.focus || "all";
      
      try {
        let reviewData: any = {};
        
        if (userId) {
          // Single user review
          const [user, licenses, groups, signIns] = await Promise.all([
            graphClient.api(`/users/${userId}`).get(),
            graphClient.api(`/users/${userId}/licenseDetails`).get(),
            graphClient.api(`/users/${userId}/memberOf`).get(),
            graphClient.api('/auditLogs/signIns').filter(`userId eq '${userId}'`).top(10).get(),
          ]);
          
          reviewData = { user, licenses, groups, signIns };
        } else {
          // Organization-wide review
          const [users, groups, roles, signIns] = await Promise.all([
            graphClient.api('/users').select('id,displayName,userPrincipalName,accountEnabled,lastSignInDateTime').get(),
            graphClient.api('/groups').get(),
            graphClient.api('/directoryRoles').get(),
            graphClient.api('/auditLogs/signIns').top(50).get(),
          ]);
          
          reviewData = { users, groups, roles, signIns };
        }
        
        const accessReviewPrompt = `# Microsoft 365 User Access Review

## Focus: ${focus}
${userId ? `## Target User: ${userId}` : '## Scope: Organization-wide'}

### Current Access Configuration:
${JSON.stringify(reviewData, null, 2)}

### Access Review Request:
Please analyze the above access data and provide:
1. Access rights summary and analysis
2. Excessive or inappropriate permissions identified
3. Inactive or stale accounts
4. License optimization opportunities
5. Group membership rationalization
6. Privileged access review
7. Recommendations for access governance improvements

${userId ? 
  'Focus on this specific user\'s access patterns and provide targeted recommendations.' : 
  'Provide organization-wide insights and patterns that require attention.'}`;

        return {
          description: `User access review focusing on ${focus}${userId ? ` for user ${userId}` : ' organization-wide'}`,
          messages: [
            {
              role: "user",
              content: {
                type: "text",
                text: accessReviewPrompt,
              },
            },
          ],
        };
      } catch (error) {
        throw new McpError(
          ErrorCode.InternalError,
          `Error generating user access review: ${error instanceof Error ? error.message : 'Unknown error'}`
        );
      }
    }
  );
    // Device Management Prompt
  server.prompt(
    "device_compliance_analysis",
    "Analyze device compliance and management status",
    {
      platform: z.string().optional().describe("Device platform to focus on (Windows, iOS, Android, macOS)"),
      complianceStatus: z.string().optional().describe("Filter by compliance status (compliant, noncompliant, unknown)"),
    },
    async (args: any) => {
      const platform = args.platform || "all";
      const complianceStatus = args.complianceStatus || "all";
      
      try {
        // Gather device management data
        const [devices, apps, compliancePolicies, configurations] = await Promise.all([
          graphClient.api('/deviceManagement/managedDevices').get(),
          graphClient.api('/deviceAppManagement/mobileApps').get(),
          graphClient.api('/deviceManagement/deviceCompliancePolicies').get(),
          graphClient.api('/deviceManagement/deviceConfigurations').get(),
        ]);
        
        const deviceAnalysisPrompt = `# Microsoft Intune Device Management Analysis

## Platform Focus: ${platform}
## Compliance Status Filter: ${complianceStatus}

### Current Device Management Configuration:

**Managed Devices:**
${JSON.stringify(devices, null, 2)}

**Mobile Applications:**
${JSON.stringify(apps, null, 2)}

**Compliance Policies:**
${JSON.stringify(compliancePolicies, null, 2)}

**Configuration Profiles:**
${JSON.stringify(configurations, null, 2)}

### Device Management Analysis Request:
Please analyze the device management data and provide:
1. Overall device compliance status
2. Non-compliant devices and reasons
3. Policy coverage gaps
4. Application deployment status
5. Configuration drift analysis
6. Security posture assessment
7. Recommendations for improved device management

Focus on actionable insights to improve device security and compliance.`;

        return {
          description: `Device compliance analysis for ${platform} platform with ${complianceStatus} status filter`,
          messages: [
            {
              role: "user",
              content: {
                type: "text",
                text: deviceAnalysisPrompt,
              },
            },
          ],
        };
      } catch (error) {
        throw new McpError(
          ErrorCode.InternalError,
          `Error generating device compliance analysis: ${error instanceof Error ? error.message : 'Unknown error'}`
        );
      }
    }
  );
    // Teams and Collaboration Prompt
  server.prompt(
    "collaboration_governance",
    "Analyze Teams and collaboration governance",
    {
      governanceArea: z.string().optional().describe("Governance area to focus on (teams, sharing, guest access, data protection)"),
      timeframe: z.string().optional().describe("Analysis timeframe (last 30 days, 90 days, 6 months)"),
    },
    async (args: any) => {
      const governanceArea = args.governanceArea || "all";
      const timeframe = args.timeframe || "30 days";
      
      try {
        // Gather collaboration data
        const [teams, groups, sites, applications] = await Promise.all([
          graphClient.api('/teams').get(),
          graphClient.api('/groups').filter("groupTypes/any(c:c eq 'Unified')").get(),
          graphClient.api('/sites?search=*').get(),
          graphClient.api('/applications').get(),
        ]);
        
        const governancePrompt = `# Microsoft 365 Collaboration Governance Analysis

## Governance Area: ${governanceArea}
## Analysis Timeframe: ${timeframe}

### Current Collaboration Environment:

**Microsoft Teams:**
${JSON.stringify(teams, null, 2)}

**Microsoft 365 Groups:**
${JSON.stringify(groups, null, 2)}

**SharePoint Sites:**
${JSON.stringify(sites, null, 2)}

**Applications:**
${JSON.stringify(applications, null, 2)}

### Governance Analysis Request:
Please analyze the collaboration environment and provide:
1. Governance maturity assessment
2. Sprawl and proliferation issues
3. Guest access and external sharing risks
4. Data protection and classification status
5. Lifecycle management gaps
6. Compliance and retention concerns
7. Recommendations for improved governance

Focus on practical governance improvements that balance security with productivity.`;

        return {
          description: `Collaboration governance analysis focusing on ${governanceArea} over ${timeframe}`,
          messages: [
            {
              role: "user",
              content: {
                type: "text",
                text: governancePrompt,
              },
            },
          ],
        };
      } catch (error) {
        throw new McpError(
          ErrorCode.InternalError,
          `Error generating collaboration governance analysis: ${error instanceof Error ? error.message : 'Unknown error'}`
        );
      }
    }
  );
}
