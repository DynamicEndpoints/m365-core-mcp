import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { 
  DLPPolicyArgs,
  RetentionPolicyArgs, 
  SensitivityLabelArgs,
  InformationProtectionPolicyArgs 
} from '../types/policy-types.js';

// DLP Policy Management Handler
export async function handleDLPPolicies(
  graphClient: Client,
  args: DLPPolicyArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'list':
      // List all DLP policies
      apiPath = '/security/informationProtection/dataLossPreventionPolicies';
      result = await graphClient.api(apiPath).get();
      break;

    case 'get':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for get action');
      }
      apiPath = `/security/informationProtection/dataLossPreventionPolicies/${args.policyId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'create':
      if (!args.displayName) {
        throw new McpError(ErrorCode.InvalidParams, 'displayName is required for create action');
      }
      
      const dlpPolicyPayload = {
        displayName: args.displayName,
        description: args.description || '',
        mode: args.mode || 'Test',
        priority: args.priority || 0,
        locations: args.locations || {},
        rules: args.rules || []
      };

      apiPath = '/security/informationProtection/dataLossPreventionPolicies';
      result = await graphClient.api(apiPath).post(dlpPolicyPayload);
      break;

    case 'update':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for update action');
      }

      const updatePayload: any = {};
      if (args.displayName) updatePayload.displayName = args.displayName;
      if (args.description) updatePayload.description = args.description;
      if (args.mode) updatePayload.mode = args.mode;
      if (args.priority !== undefined) updatePayload.priority = args.priority;
      if (args.locations) updatePayload.locations = args.locations;
      if (args.rules) updatePayload.rules = args.rules;

      apiPath = `/security/informationProtection/dataLossPreventionPolicies/${args.policyId}`;
      result = await graphClient.api(apiPath).patch(updatePayload);
      break;

    case 'delete':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for delete action');
      }
      apiPath = `/security/informationProtection/dataLossPreventionPolicies/${args.policyId}`;
      await graphClient.api(apiPath).delete();
      result = { message: `DLP policy ${args.policyId} deleted successfully` };
      break;

    case 'enable':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for enable action');
      }
      apiPath = `/security/informationProtection/dataLossPreventionPolicies/${args.policyId}`;
      result = await graphClient.api(apiPath).patch({ mode: 'Enforce' });
      break;

    case 'disable':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for disable action');
      }
      apiPath = `/security/informationProtection/dataLossPreventionPolicies/${args.policyId}`;
      result = await graphClient.api(apiPath).patch({ mode: 'Test' });
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Unknown action: ${args.action}`);
  }

  return {
    content: [{
      type: 'text',
      text: `DLP Policy ${args.action} operation completed:\n\n${JSON.stringify(result, null, 2)}`
    }]
  };
}

// Retention Policy Management Handler
export async function handleRetentionPolicies(
  graphClient: Client,
  args: RetentionPolicyArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'list':
      // List all retention policies
      apiPath = '/security/informationProtection/retentionPolicies';
      result = await graphClient.api(apiPath).get();
      break;

    case 'get':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for get action');
      }
      apiPath = `/security/informationProtection/retentionPolicies/${args.policyId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'create':
      if (!args.displayName || !args.retentionSettings) {
        throw new McpError(ErrorCode.InvalidParams, 'displayName and retentionSettings are required for create action');
      }
      
      const retentionPolicyPayload = {
        displayName: args.displayName,
        description: args.description || '',
        isEnabled: args.isEnabled !== undefined ? args.isEnabled : true,
        retentionSettings: args.retentionSettings,
        locations: args.locations || {}
      };

      apiPath = '/security/informationProtection/retentionPolicies';
      result = await graphClient.api(apiPath).post(retentionPolicyPayload);
      break;

    case 'update':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for update action');
      }

      const updatePayload: any = {};
      if (args.displayName) updatePayload.displayName = args.displayName;
      if (args.description) updatePayload.description = args.description;
      if (args.isEnabled !== undefined) updatePayload.isEnabled = args.isEnabled;
      if (args.retentionSettings) updatePayload.retentionSettings = args.retentionSettings;
      if (args.locations) updatePayload.locations = args.locations;

      apiPath = `/security/informationProtection/retentionPolicies/${args.policyId}`;
      result = await graphClient.api(apiPath).patch(updatePayload);
      break;

    case 'delete':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for delete action');
      }
      apiPath = `/security/informationProtection/retentionPolicies/${args.policyId}`;
      await graphClient.api(apiPath).delete();
      result = { message: `Retention policy ${args.policyId} deleted successfully` };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Unknown action: ${args.action}`);
  }

  return {
    content: [{
      type: 'text',
      text: `Retention Policy ${args.action} operation completed:\n\n${JSON.stringify(result, null, 2)}`
    }]
  };
}

// Sensitivity Label Management Handler
export async function handleSensitivityLabels(
  graphClient: Client,
  args: SensitivityLabelArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'list':
      // List all sensitivity labels
      apiPath = '/security/informationProtection/sensitivityLabels';
      result = await graphClient.api(apiPath).get();
      break;

    case 'get':
      if (!args.labelId) {
        throw new McpError(ErrorCode.InvalidParams, 'labelId is required for get action');
      }
      apiPath = `/security/informationProtection/sensitivityLabels/${args.labelId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'create':
      if (!args.displayName) {
        throw new McpError(ErrorCode.InvalidParams, 'displayName is required for create action');
      }
      
      const sensitivityLabelPayload: any = {
        displayName: args.displayName,
        description: args.description || '',
        tooltip: args.tooltip || args.description || '',
        priority: args.priority || 0,
        isEnabled: args.isEnabled !== undefined ? args.isEnabled : true,
        labelActions: [],
        applicableTo: 'EmailMessage,File'
      };

      // Add settings if provided
      if (args.settings) {
        if (args.settings.contentMarking) {
          sensitivityLabelPayload.labelActions.push({
            '@odata.type': 'microsoft.graph.contentMarkingLabelAction',
            ...args.settings.contentMarking
          });
        }

        if (args.settings.encryption && args.settings.encryption.enabled) {
          sensitivityLabelPayload.labelActions.push({
            '@odata.type': 'microsoft.graph.encryptionLabelAction',
            ...args.settings.encryption
          });
        }

        if (args.settings.accessControl) {
          sensitivityLabelPayload.labelActions.push({
            '@odata.type': 'microsoft.graph.accessControlLabelAction',
            ...args.settings.accessControl
          });
        }

        if (args.settings.autoLabeling && args.settings.autoLabeling.enabled) {
          sensitivityLabelPayload.labelActions.push({
            '@odata.type': 'microsoft.graph.autoLabelingLabelAction',
            ...args.settings.autoLabeling
          });
        }
      }

      apiPath = '/security/informationProtection/sensitivityLabels';
      result = await graphClient.api(apiPath).post(sensitivityLabelPayload);
      break;

    case 'update':
      if (!args.labelId) {
        throw new McpError(ErrorCode.InvalidParams, 'labelId is required for update action');
      }

      const updatePayload: any = {};
      if (args.displayName) updatePayload.displayName = args.displayName;
      if (args.description) updatePayload.description = args.description;
      if (args.tooltip) updatePayload.tooltip = args.tooltip;
      if (args.priority !== undefined) updatePayload.priority = args.priority;
      if (args.isEnabled !== undefined) updatePayload.isEnabled = args.isEnabled;

      // Handle settings updates
      if (args.settings) {
        updatePayload.labelActions = [];
        
        if (args.settings.contentMarking) {
          updatePayload.labelActions.push({
            '@odata.type': 'microsoft.graph.contentMarkingLabelAction',
            ...args.settings.contentMarking
          });
        }

        if (args.settings.encryption && args.settings.encryption.enabled) {
          updatePayload.labelActions.push({
            '@odata.type': 'microsoft.graph.encryptionLabelAction',
            ...args.settings.encryption
          });
        }

        if (args.settings.accessControl) {
          updatePayload.labelActions.push({
            '@odata.type': 'microsoft.graph.accessControlLabelAction',
            ...args.settings.accessControl
          });
        }

        if (args.settings.autoLabeling && args.settings.autoLabeling.enabled) {
          updatePayload.labelActions.push({
            '@odata.type': 'microsoft.graph.autoLabelingLabelAction',
            ...args.settings.autoLabeling
          });
        }
      }

      apiPath = `/security/informationProtection/sensitivityLabels/${args.labelId}`;
      result = await graphClient.api(apiPath).patch(updatePayload);
      break;

    case 'delete':
      if (!args.labelId) {
        throw new McpError(ErrorCode.InvalidParams, 'labelId is required for delete action');
      }
      apiPath = `/security/informationProtection/sensitivityLabels/${args.labelId}`;
      await graphClient.api(apiPath).delete();
      result = { message: `Sensitivity label ${args.labelId} deleted successfully` };
      break;

    case 'publish':
      if (!args.labelId) {
        throw new McpError(ErrorCode.InvalidParams, 'labelId is required for publish action');
      }
      
      // Create a label policy to publish the label
      const publishPayload = {
        displayName: `${args.displayName || 'Label'} Policy`,
        description: `Policy for publishing sensitivity label`,
        labels: [args.labelId],
        settings: {
          mandatory: false,
          requireJustification: false
        }
      };

      apiPath = '/security/informationProtection/labelPolicies';
      result = await graphClient.api(apiPath).post(publishPayload);
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Unknown action: ${args.action}`);
  }

  return {
    content: [{
      type: 'text',
      text: `Sensitivity Label ${args.action} operation completed:\n\n${JSON.stringify(result, null, 2)}`
    }]
  };
}

// Information Protection Policy Management Handler
export async function handleInformationProtectionPolicies(
  graphClient: Client,
  args: InformationProtectionPolicyArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'list':
      // List all information protection policies
      apiPath = '/security/informationProtection/labelPolicies';
      result = await graphClient.api(apiPath).get();
      break;

    case 'get':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for get action');
      }
      apiPath = `/security/informationProtection/labelPolicies/${args.policyId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'create':
      if (!args.displayName) {
        throw new McpError(ErrorCode.InvalidParams, 'displayName is required for create action');
      }
      
      const infoPolicyPayload = {
        displayName: args.displayName,
        description: args.description || '',
        settings: args.settings || {}
      };

      apiPath = '/security/informationProtection/labelPolicies';
      result = await graphClient.api(apiPath).post(infoPolicyPayload);
      break;

    case 'update':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for update action');
      }

      const updatePayload: any = {};
      if (args.displayName) updatePayload.displayName = args.displayName;
      if (args.description) updatePayload.description = args.description;
      if (args.settings) updatePayload.settings = args.settings;

      apiPath = `/security/informationProtection/labelPolicies/${args.policyId}`;
      result = await graphClient.api(apiPath).patch(updatePayload);
      break;

    case 'delete':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for delete action');
      }
      apiPath = `/security/informationProtection/labelPolicies/${args.policyId}`;
      await graphClient.api(apiPath).delete();
      result = { message: `Information protection policy ${args.policyId} deleted successfully` };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Unknown action: ${args.action}`);
  }

  return {
    content: [{
      type: 'text',
      text: `Information Protection Policy ${args.action} operation completed:\n\n${JSON.stringify(result, null, 2)}`
    }]
  };
}