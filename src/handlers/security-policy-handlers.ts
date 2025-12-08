import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { 
  DefenderPolicyArgs,
  TeamsPolicyArgs,
  ExchangePolicyArgs,
  SharePointGovernancePolicyArgs,
  SecurityAlertPolicyArgs 
} from '../types/policy-types.js';

// Microsoft Defender for Office 365 Policy Handler
export async function handleDefenderPolicies(
  graphClient: Client,
  args: DefenderPolicyArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  // Map policy types to API endpoints
  const policyEndpoints = {
    safeAttachments: '/security/attackSimulation/safeAttachmentPolicies',
    safeLinks: '/security/attackSimulation/safeLinksPolicies',
    antiPhishing: '/security/antiPhishingPolicies',
    antiMalware: '/security/antiMalwarePolicies',
    antiSpam: '/security/antiSpamPolicies'
  };

  const endpoint = policyEndpoints[args.policyType];
  if (!endpoint) {
    throw new McpError(ErrorCode.InvalidParams, `Unsupported policy type: ${args.policyType}`);
  }

  switch (args.action) {
    case 'list':
      apiPath = endpoint;
      result = await graphClient.api(apiPath).get();
      break;

    case 'get':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for get action');
      }
      apiPath = `${endpoint}/${args.policyId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'create':
      if (!args.displayName) {
        throw new McpError(ErrorCode.InvalidParams, 'displayName is required for create action');
      }
      
      const defenderPolicyPayload: any = {
        displayName: args.displayName,
        description: args.description || '',
        isEnabled: args.isEnabled !== undefined ? args.isEnabled : true,
        settings: args.settings || {},
        appliedTo: args.appliedTo || {}
      };

      apiPath = endpoint;
      result = await graphClient.api(apiPath).post(defenderPolicyPayload);
      break;

    case 'update':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for update action');
      }

      const updatePayload: any = {};
      if (args.displayName) updatePayload.displayName = args.displayName;
      if (args.description) updatePayload.description = args.description;
      if (args.isEnabled !== undefined) updatePayload.isEnabled = args.isEnabled;
      if (args.settings) updatePayload.settings = args.settings;
      if (args.appliedTo) updatePayload.appliedTo = args.appliedTo;

      apiPath = `${endpoint}/${args.policyId}`;
      result = await graphClient.api(apiPath).patch(updatePayload);
      break;

    case 'delete':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for delete action');
      }
      apiPath = `${endpoint}/${args.policyId}`;
      await graphClient.api(apiPath).delete();
      result = { message: `${args.policyType} policy ${args.policyId} deleted successfully` };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Unknown action: ${args.action}`);
  }

  return {
    content: [{
      type: 'text',
      text: `Defender ${args.policyType} Policy ${args.action} operation completed:\n\n${JSON.stringify(result, null, 2)}`
    }]
  };
}

// Microsoft Teams Policy Handler
export async function handleTeamsPolicies(
  graphClient: Client,
  args: TeamsPolicyArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  // Map policy types to API endpoints
  const policyEndpoints = {
    messaging: '/admin/serviceAnnouncement/policies/messaging',
    meeting: '/admin/serviceAnnouncement/policies/meeting',
    calling: '/admin/serviceAnnouncement/policies/calling',
    appSetup: '/admin/serviceAnnouncement/policies/appSetup',
    updateManagement: '/admin/serviceAnnouncement/policies/updateManagement'
  };

  const endpoint = policyEndpoints[args.policyType];
  if (!endpoint) {
    throw new McpError(ErrorCode.InvalidParams, `Unsupported policy type: ${args.policyType}`);
  }

  switch (args.action) {
    case 'list':
      apiPath = endpoint;
      result = await graphClient.api(apiPath).get();
      break;

    case 'get':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for get action');
      }
      apiPath = `${endpoint}/${args.policyId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'create':
      if (!args.displayName) {
        throw new McpError(ErrorCode.InvalidParams, 'displayName is required for create action');
      }
      
      const teamsPolicyPayload: any = {
        displayName: args.displayName,
        description: args.description || '',
        settings: args.settings || {}
      };

      apiPath = endpoint;
      result = await graphClient.api(apiPath).post(teamsPolicyPayload);
      break;

    case 'update':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for update action');
      }

      const updatePayload: any = {};
      if (args.displayName) updatePayload.displayName = args.displayName;
      if (args.description) updatePayload.description = args.description;
      if (args.settings) updatePayload.settings = args.settings;

      apiPath = `${endpoint}/${args.policyId}`;
      result = await graphClient.api(apiPath).patch(updatePayload);
      break;

    case 'delete':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for delete action');
      }
      apiPath = `${endpoint}/${args.policyId}`;
      await graphClient.api(apiPath).delete();
      result = { message: `Teams ${args.policyType} policy ${args.policyId} deleted successfully` };
      break;

    case 'assign':
      if (!args.policyId || !args.assignTo) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId and assignTo are required for assign action');
      }
      
      const assignPayload: any = {
        assignments: []
      };

      if (args.assignTo.users) {
        args.assignTo.users.forEach(userId => {
          assignPayload.assignments.push({
            target: {
              '@odata.type': '#microsoft.graph.userTarget',
              userId: userId
            }
          });
        });
      }

      if (args.assignTo.groups) {
        args.assignTo.groups.forEach(groupId => {
          assignPayload.assignments.push({
            target: {
              '@odata.type': '#microsoft.graph.groupTarget',
              groupId: groupId
            }
          });
        });
      }

      apiPath = `${endpoint}/${args.policyId}/assign`;
      result = await graphClient.api(apiPath).post(assignPayload);
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Unknown action: ${args.action}`);
  }

  return {
    content: [{
      type: 'text',
      text: `Teams ${args.policyType} Policy ${args.action} operation completed:\n\n${JSON.stringify(result, null, 2)}`
    }]
  };
}

// Exchange Online Policy Handler
export async function handleExchangePolicies(
  graphClient: Client,
  args: ExchangePolicyArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  // Map policy types to API endpoints
  const policyEndpoints = {
    addressBook: '/admin/exchange/addressBookPolicies',
    outlookWebApp: '/admin/exchange/owaMailboxPolicies',
    activeSyncMailbox: '/admin/exchange/activeSyncMailboxPolicies',
    retentionPolicy: '/admin/exchange/retentionPolicies',
    dlpPolicy: '/admin/exchange/dataLossPreventionPolicies'
  };

  const endpoint = policyEndpoints[args.policyType];
  if (!endpoint) {
    throw new McpError(ErrorCode.InvalidParams, `Unsupported policy type: ${args.policyType}`);
  }

  switch (args.action) {
    case 'list':
      apiPath = endpoint;
      result = await graphClient.api(apiPath).get();
      break;

    case 'get':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for get action');
      }
      apiPath = `${endpoint}/${args.policyId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'create':
      if (!args.displayName) {
        throw new McpError(ErrorCode.InvalidParams, 'displayName is required for create action');
      }
      
      const exchangePolicyPayload: any = {
        displayName: args.displayName,
        description: args.description || '',
        isDefault: args.isDefault || false,
        settings: args.settings || {},
        appliedTo: args.appliedTo || {}
      };

      apiPath = endpoint;
      result = await graphClient.api(apiPath).post(exchangePolicyPayload);
      break;

    case 'update':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for update action');
      }

      const updatePayload: any = {};
      if (args.displayName) updatePayload.displayName = args.displayName;
      if (args.description) updatePayload.description = args.description;
      if (args.isDefault !== undefined) updatePayload.isDefault = args.isDefault;
      if (args.settings) updatePayload.settings = args.settings;
      if (args.appliedTo) updatePayload.appliedTo = args.appliedTo;

      apiPath = `${endpoint}/${args.policyId}`;
      result = await graphClient.api(apiPath).patch(updatePayload);
      break;

    case 'delete':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for delete action');
      }
      apiPath = `${endpoint}/${args.policyId}`;
      await graphClient.api(apiPath).delete();
      result = { message: `Exchange ${args.policyType} policy ${args.policyId} deleted successfully` };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Unknown action: ${args.action}`);
  }

  return {
    content: [{
      type: 'text',
      text: `Exchange ${args.policyType} Policy ${args.action} operation completed:\n\n${JSON.stringify(result, null, 2)}`
    }]
  };
}

// SharePoint Governance Policy Handler
export async function handleSharePointGovernancePolicies(
  graphClient: Client,
  args: SharePointGovernancePolicyArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  // Map policy types to API endpoints
  const policyEndpoints = {
    sharingPolicy: '/admin/sharepoint/settings/sharing',
    accessPolicy: '/admin/sharepoint/settings/conditionalAccess',
    informationBarrier: '/admin/sharepoint/settings/informationBarriers',
    retentionLabel: '/admin/sharepoint/settings/retentionLabels'
  };

  const endpoint = policyEndpoints[args.policyType];
  if (!endpoint) {
    throw new McpError(ErrorCode.InvalidParams, `Unsupported policy type: ${args.policyType}`);
  }

  switch (args.action) {
    case 'list':
      apiPath = endpoint;
      result = await graphClient.api(apiPath).get();
      break;

    case 'get':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for get action');
      }
      apiPath = `${endpoint}/${args.policyId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'create':
      if (!args.displayName) {
        throw new McpError(ErrorCode.InvalidParams, 'displayName is required for create action');
      }
      
      const spPolicyPayload: any = {
        displayName: args.displayName,
        description: args.description || '',
        scope: args.scope || {},
        settings: args.settings || {}
      };

      apiPath = endpoint;
      result = await graphClient.api(apiPath).post(spPolicyPayload);
      break;

    case 'update':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for update action');
      }

      const updatePayload: any = {};
      if (args.displayName) updatePayload.displayName = args.displayName;
      if (args.description) updatePayload.description = args.description;
      if (args.scope) updatePayload.scope = args.scope;
      if (args.settings) updatePayload.settings = args.settings;

      apiPath = `${endpoint}/${args.policyId}`;
      result = await graphClient.api(apiPath).patch(updatePayload);
      break;

    case 'delete':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for delete action');
      }
      apiPath = `${endpoint}/${args.policyId}`;
      await graphClient.api(apiPath).delete();
      result = { message: `SharePoint ${args.policyType} policy ${args.policyId} deleted successfully` };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Unknown action: ${args.action}`);
  }

  return {
    content: [{
      type: 'text',
      text: `SharePoint ${args.policyType} Policy ${args.action} operation completed:\n\n${JSON.stringify(result, null, 2)}`
    }]
  };
}

// Security and Compliance Alert Policy Handler
export async function handleSecurityAlertPolicies(
  graphClient: Client,
  args: SecurityAlertPolicyArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'list':
      apiPath = '/security/alerts/policies';
      result = await graphClient.api(apiPath).get();
      break;

    case 'get':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for get action');
      }
      apiPath = `/security/alerts/policies/${args.policyId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'create':
      if (!args.displayName) {
        throw new McpError(ErrorCode.InvalidParams, 'displayName is required for create action');
      }
      
      const alertPolicyPayload: any = {
        displayName: args.displayName,
        description: args.description || '',
        category: args.category || 'Others',
        severity: args.severity || 'Medium',
        isEnabled: args.isEnabled !== undefined ? args.isEnabled : true,
        conditions: args.conditions || {},
        actions: args.actions || {}
      };

      apiPath = '/security/alerts/policies';
      result = await graphClient.api(apiPath).post(alertPolicyPayload);
      break;

    case 'update':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for update action');
      }

      const updatePayload: any = {};
      if (args.displayName) updatePayload.displayName = args.displayName;
      if (args.description) updatePayload.description = args.description;
      if (args.category) updatePayload.category = args.category;
      if (args.severity) updatePayload.severity = args.severity;
      if (args.isEnabled !== undefined) updatePayload.isEnabled = args.isEnabled;
      if (args.conditions) updatePayload.conditions = args.conditions;
      if (args.actions) updatePayload.actions = args.actions;

      apiPath = `/security/alerts/policies/${args.policyId}`;
      result = await graphClient.api(apiPath).patch(updatePayload);
      break;

    case 'delete':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for delete action');
      }
      apiPath = `/security/alerts/policies/${args.policyId}`;
      await graphClient.api(apiPath).delete();
      result = { message: `Security alert policy ${args.policyId} deleted successfully` };
      break;

    case 'enable':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for enable action');
      }
      apiPath = `/security/alerts/policies/${args.policyId}`;
      result = await graphClient.api(apiPath).patch({ isEnabled: true });
      break;

    case 'disable':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for disable action');
      }
      apiPath = `/security/alerts/policies/${args.policyId}`;
      result = await graphClient.api(apiPath).patch({ isEnabled: false });
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Unknown action: ${args.action}`);
  }

  return {
    content: [{
      type: 'text',
      text: `Security Alert Policy ${args.action} operation completed:\n\n${JSON.stringify(result, null, 2)}`
    }]
  };
}