import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { ConditionalAccessPolicyArgs } from '../types/policy-types.js';

// Conditional Access Policy Management Handler
export async function handleConditionalAccessPolicies(
  graphClient: Client,
  args: ConditionalAccessPolicyArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'list':
      // List all Conditional Access policies
      apiPath = '/identity/conditionalAccess/policies';
      result = await graphClient.api(apiPath).get();
      break;

    case 'get':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for get action');
      }
      apiPath = `/identity/conditionalAccess/policies/${args.policyId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'create':
      if (!args.displayName) {
        throw new McpError(ErrorCode.InvalidParams, 'displayName is required for create action');
      }
      
      const caPolicyPayload: any = {
        displayName: args.displayName,
        description: args.description || '',
        state: args.state || 'disabled',
        conditions: args.conditions || {
          users: {
            includeUsers: ['All']
          },
          applications: {
            includeApplications: ['All']
          }
        },
        grantControls: args.grantControls || {
          operator: 'OR',
          builtInControls: ['mfa']
        }
      };

      // Add session controls if provided
      if (args.sessionControls) {
        caPolicyPayload.sessionControls = args.sessionControls;
      }

      apiPath = '/identity/conditionalAccess/policies';
      result = await graphClient.api(apiPath).post(caPolicyPayload);
      break;

    case 'update':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for update action');
      }

      const updatePayload: any = {};
      if (args.displayName) updatePayload.displayName = args.displayName;
      if (args.description) updatePayload.description = args.description;
      if (args.state) updatePayload.state = args.state;
      if (args.conditions) updatePayload.conditions = args.conditions;
      if (args.grantControls) updatePayload.grantControls = args.grantControls;
      if (args.sessionControls) updatePayload.sessionControls = args.sessionControls;

      apiPath = `/identity/conditionalAccess/policies/${args.policyId}`;
      result = await graphClient.api(apiPath).patch(updatePayload);
      break;

    case 'delete':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for delete action');
      }
      apiPath = `/identity/conditionalAccess/policies/${args.policyId}`;
      await graphClient.api(apiPath).delete();
      result = { message: `Conditional Access policy ${args.policyId} deleted successfully` };
      break;

    case 'enable':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for enable action');
      }
      apiPath = `/identity/conditionalAccess/policies/${args.policyId}`;
      result = await graphClient.api(apiPath).patch({ state: 'enabled' });
      break;

    case 'disable':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for disable action');
      }
      apiPath = `/identity/conditionalAccess/policies/${args.policyId}`;
      result = await graphClient.api(apiPath).patch({ state: 'disabled' });
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Unknown action: ${args.action}`);
  }

  return {
    content: [{
      type: 'text',
      text: `Conditional Access Policy ${args.action} operation completed:\n\n${JSON.stringify(result, null, 2)}`
    }]
  };
}