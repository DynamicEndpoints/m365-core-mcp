
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { handleIntuneMacOSPolicies } from './intune-macos-handler.js';
import { handleIntuneWindowsPolicies } from './intune-windows-handler.js';

export async function handleCreateIntunePolicy(graphClient: Client, args: any): Promise<any> {
  const { platform, policyType, displayName, description, settings, assignments } = args;

  if (!platform || !policyType || !displayName || !settings) {
    throw new McpError(ErrorCode.InvalidParams, 'Missing required parameters.');
  }

  const policyArgs = {
    action: 'create',
    policyType: policyType,
    name: displayName,
    description: description,
    settings: settings,
    assignments: assignments
  };

  if (platform === 'macos') {
    return await handleIntuneMacOSPolicies(graphClient, policyArgs as any);
  } else if (platform === 'windows') {
    return await handleIntuneWindowsPolicies(graphClient, policyArgs as any);
  } else {
    throw new McpError(ErrorCode.InvalidParams, `Unsupported platform: ${platform}`);
  }
}
