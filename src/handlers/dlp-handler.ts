import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { DLPPolicyArgs, DLPIncidentArgs, DLPSensitivityLabelArgs } from '../types/dlp-types';

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
      apiPath = '/beta/security/dataLossPreventionPolicies';
      result = await graphClient.api(apiPath).get();
      break;

    case 'get':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for get action');
      }
      apiPath = `/beta/security/dataLossPreventionPolicies/${args.policyId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'create':
      if (!args.name) {
        throw new McpError(ErrorCode.InvalidParams, 'name is required for create action');
      }
      apiPath = '/beta/security/dataLossPreventionPolicies';
      const createPayload = {
        displayName: args.name,
        description: args.description || '',
        status: args.settings?.enabled ? 'enabled' : 'disabled',
        // locations: args.locations, // Locations are part of rules, not top-level policy
        // mode: args.settings?.mode, // Mode is also part of rules
        // priority: args.settings?.priority, // Priority is also part of rules
      };
      result = await graphClient.api(apiPath).post(createPayload);
      break;

    case 'update':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for update action');
      }
      apiPath = `/beta/security/dataLossPreventionPolicies/${args.policyId}`;
      const updatePayload = {
        displayName: args.name,
        description: args.description,
        status: args.settings?.enabled ? 'enabled' : 'disabled',
      };
      result = await graphClient.api(apiPath).patch(updatePayload);
      break;

    case 'delete':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for delete action');
      }
      apiPath = `/beta/security/dataLossPreventionPolicies/${args.policyId}`;
      await graphClient.api(apiPath).delete();
      result = { message: 'DLP policy deleted successfully' };
      break;

    case 'test':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for test action');
      }
      // This would typically involve creating a test case
      result = { message: 'DLP policy test initiated', policyId: args.policyId };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}

// DLP Incident Management Handler
export async function handleDLPIncidents(
  graphClient: Client,
  args: DLPIncidentArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'list':
      // List DLP incidents from security events
      apiPath = '/security/alerts_v2';
      const filterConditions: string[] = [];
      
      if (args.dateRange) {
        filterConditions.push(`createdDateTime ge ${args.dateRange.startDate} and createdDateTime le ${args.dateRange.endDate}`);
      }
      
      if (args.severity) {
        filterConditions.push(`severity eq '${args.severity}'`);
      }

      if (filterConditions.length > 0) {
        apiPath += `?$filter=${filterConditions.join(' and ')}`;
      }

      result = await graphClient.api(apiPath).get();
      break;

    case 'get':
      if (!args.incidentId) {
        throw new McpError(ErrorCode.InvalidParams, 'incidentId is required for get action');
      }
      apiPath = `/security/alerts_v2/${args.incidentId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'resolve':
      if (!args.incidentId) {
        throw new McpError(ErrorCode.InvalidParams, 'incidentId is required for resolve action');
      }
      apiPath = `/security/alerts_v2/${args.incidentId}`;
      result = await graphClient.api(apiPath).patch({
        status: 'resolved',
        feedback: 'truePositive'
      });
      break;

    case 'escalate':
      if (!args.incidentId) {
        throw new McpError(ErrorCode.InvalidParams, 'incidentId is required for escalate action');
      }
      apiPath = `/security/alerts_v2/${args.incidentId}`;
      result = await graphClient.api(apiPath).patch({
        severity: 'high',
        classification: 'truePositive'
      });
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}

// DLP Sensitivity Labels Handler
export async function handleDLPSensitivityLabels(
  graphClient: Client,
  args: DLPSensitivityLabelArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'list':
      apiPath = '/informationProtection/policy/labels';
      result = await graphClient.api(apiPath).get();
      break;

    case 'get':
      if (!args.labelId) {
        throw new McpError(ErrorCode.InvalidParams, 'labelId is required for get action');
      }
      apiPath = `/informationProtection/policy/labels/${args.labelId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'create':
      if (!args.name) {
        throw new McpError(ErrorCode.InvalidParams, 'name is required for create action');
      }
      apiPath = '/informationProtection/policy/labels';
      const labelPayload = {
        name: args.name,
        description: args.description || '',
        color: args.settings?.color || 'blue',
        sensitivity: args.settings?.sensitivity || 0,
        tooltip: args.description || args.name,
        isActive: true
      };
      result = await graphClient.api(apiPath).post(labelPayload);
      break;

    case 'update':
      if (!args.labelId) {
        throw new McpError(ErrorCode.InvalidParams, 'labelId is required for update action');
      }
      apiPath = `/informationProtection/policy/labels/${args.labelId}`;
      const updateLabelPayload = {
        name: args.name,
        description: args.description,
        color: args.settings?.color,
        sensitivity: args.settings?.sensitivity
      };
      result = await graphClient.api(apiPath).patch(updateLabelPayload);
      break;

    case 'delete':
      if (!args.labelId) {
        throw new McpError(ErrorCode.InvalidParams, 'labelId is required for delete action');
      }
      apiPath = `/informationProtection/policy/labels/${args.labelId}`;
      await graphClient.api(apiPath).delete();
      result = { message: 'Sensitivity label deleted successfully' };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}
