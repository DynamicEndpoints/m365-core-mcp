import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';
import {
  UserManagementArgs,
  OffboardingArgs,
  SharePointSiteArgs,
  SharePointListArgs,
  AzureAdRoleArgs,
  AzureAdAppArgs,
  AzureAdDeviceArgs,
  AzureAdSpArgs,
  CallMicrosoftApiArgs,
  AuditLogArgs,
  AlertArgs,
} from './types.js';

// User Management Handler
export async function handleUserSettings(
  graphClient: Client,
  args: UserManagementArgs
): Promise<{ content: { type: string; text: string }[] }> {
  if (args.action === 'get') {
    const settings = await graphClient
      .api(`/users/${args.userId}`)
      .get();
    return { content: [{ type: 'text', text: JSON.stringify(settings, null, 2) }] };
  } else {
    await graphClient
      .api(`/users/${args.userId}`)
      .patch(args.settings);
    return { content: [{ type: 'text', text: 'User settings updated successfully' }] };
  }
}

// Offboarding Handler
export async function handleOffboarding(
  graphClient: Client,
  args: OffboardingArgs
): Promise<{ content: { type: string; text: string }[] }> {
  switch (args.action) {
    case 'start': {
      // Block sign-ins
      await graphClient
        .api(`/users/${args.userId}`)
        .patch({ accountEnabled: false });

      if (args.options?.revokeAccess) {
        // Revoke all refresh tokens
        await graphClient
          .api(`/users/${args.userId}/revokeSignInSessions`)
          .post({});
      }

      if (args.options?.backupData) {
        // Trigger backup
        await graphClient
          .api(`/users/${args.userId}/drive/content`)
          .get();
      }

      return { content: [{ type: 'text', text: 'Offboarding process started successfully' }] };
    }
    case 'check': {
      const status = await graphClient
        .api(`/users/${args.userId}`)
        .get();
      return { content: [{ type: 'text', text: JSON.stringify(status, null, 2) }] };
    }
    case 'complete': {
      if (args.options?.convertToShared) {
        // Convert to shared mailbox
        await graphClient
          .api(`/users/${args.userId}/mailbox/convert`)
          .post({});
      } else if (!args.options?.retainMailbox) {
        // Delete user if not retaining mailbox
        await graphClient
          .api(`/users/${args.userId}`)
          .delete();
      }
      return { content: [{ type: 'text', text: 'Offboarding process completed successfully' }] };
    }
    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }
}

// SharePoint Site Handler
export async function handleSharePointSite(
  graphClient: Client,
  args: SharePointSiteArgs
): Promise<{ content: { type: string; text: string }[] }> {
  switch (args.action) {
    case 'get': {
      const site = await graphClient
        .api(`/sites/${args.siteId}`)
        .get();
      return { content: [{ type: 'text', text: JSON.stringify(site, null, 2) }] };
    }
    case 'create': {
      // Create a new SharePoint site
      const site = await graphClient
        .api('/sites/add')
        .post({
          displayName: args.title,
          description: args.description,
          webTemplate: args.template || 'STS#0', // Team site template
          url: args.url,
        });
      
      // Apply settings if provided
      if (args.settings) {
        await graphClient
          .api(`/sites/${site.id}/settings`)
          .patch({
            isPublic: args.settings.isPublic,
            sharingCapability: args.settings.allowSharing ? 'ExternalUserSharingOnly' : 'Disabled',
            storageQuota: args.settings.storageQuota,
          });
      }
      
      // Add owners if provided
      if (args.owners?.length) {
        for (const owner of args.owners) {
          await graphClient
            .api(`/sites/${site.id}/owners/$ref`)
            .post({
              '@odata.id': `https://graph.microsoft.com/v1.0/users/${owner}`,
            });
        }
      }
      
      // Add members if provided
      if (args.members?.length) {
        for (const member of args.members) {
          await graphClient
            .api(`/sites/${site.id}/members/$ref`)
            .post({
              '@odata.id': `https://graph.microsoft.com/v1.0/users/${member}`,
            });
        }
      }
      
      return { content: [{ type: 'text', text: JSON.stringify(site, null, 2) }] };
    }
    case 'update': {
      // Update site properties
      await graphClient
        .api(`/sites/${args.siteId}`)
        .patch({
          displayName: args.title,
          description: args.description,
        });
      
      // Update settings if provided
      if (args.settings) {
        await graphClient
          .api(`/sites/${args.siteId}/settings`)
          .patch({
            isPublic: args.settings.isPublic,
            sharingCapability: args.settings.allowSharing ? 'ExternalUserSharingOnly' : 'Disabled',
            storageQuota: args.settings.storageQuota,
          });
      }
      
      return { content: [{ type: 'text', text: 'SharePoint site updated successfully' }] };
    }
    case 'delete': {
      await graphClient
        .api(`/sites/${args.siteId}`)
        .delete();
      return { content: [{ type: 'text', text: 'SharePoint site deleted successfully' }] };
    }
    case 'add_users': {
      if (!args.members?.length) {
        throw new McpError(ErrorCode.InvalidParams, 'No users specified to add');
      }
      
      for (const member of args.members) {
        await graphClient
          .api(`/sites/${args.siteId}/members/$ref`)
          .post({
            '@odata.id': `https://graph.microsoft.com/v1.0/users/${member}`,
          });
      }
      
      return { content: [{ type: 'text', text: 'Users added to SharePoint site successfully' }] };
    }
    case 'remove_users': {
      if (!args.members?.length) {
        throw new McpError(ErrorCode.InvalidParams, 'No users specified to remove');
      }
      
      for (const member of args.members) {
        await graphClient
          .api(`/sites/${args.siteId}/members/${member}/$ref`)
          .delete();
      }
      
      return { content: [{ type: 'text', text: 'Users removed from SharePoint site successfully' }] };
    }
    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }
}

// SharePoint List Handler
export async function handleSharePointList(
  graphClient: Client,
  args: SharePointListArgs
): Promise<{ content: { type: string; text: string }[] }> {
  switch (args.action) {
    case 'get': {
      const list = await graphClient
        .api(`/sites/${args.siteId}/lists/${args.listId}`)
        .get();
      return { content: [{ type: 'text', text: JSON.stringify(list, null, 2) }] };
    }
    case 'create': {
      // Create a new list
      const list = await graphClient
        .api(`/sites/${args.siteId}/lists`)
        .post({
          displayName: args.title,
          description: args.description,
          template: args.template || 'genericList',
        });
      
      // Add columns if provided
      if (args.columns?.length) {
        for (const column of args.columns) {
          await graphClient
            .api(`/sites/${args.siteId}/lists/${list.id}/columns`)
            .post({
              name: column.name,
              columnType: column.type,
              required: column.required || false,
              defaultValue: column.defaultValue,
            });
        }
      }
      
      return { content: [{ type: 'text', text: JSON.stringify(list, null, 2) }] };
    }
    case 'update': {
      await graphClient
        .api(`/sites/${args.siteId}/lists/${args.listId}`)
        .patch({
          displayName: args.title,
          description: args.description,
        });
      
      return { content: [{ type: 'text', text: 'SharePoint list updated successfully' }] };
    }
    case 'delete': {
      await graphClient
        .api(`/sites/${args.siteId}/lists/${args.listId}`)
        .delete();
      
      return { content: [{ type: 'text', text: 'SharePoint list deleted successfully' }] };
    }
    case 'add_items': {
      if (!args.items?.length) {
        throw new McpError(ErrorCode.InvalidParams, 'No items specified to add');
      }
      
      const results = [];
      for (const item of args.items) {
        const result = await graphClient
          .api(`/sites/${args.siteId}/lists/${args.listId}/items`)
          .post({
            fields: item,
          });
        
        results.push(result);
      }
      
      return { content: [{ type: 'text', text: JSON.stringify(results, null, 2) }] };
    }
    case 'get_items': {
      const items = await graphClient
        .api(`/sites/${args.siteId}/lists/${args.listId}/items?expand=fields`)
        .get();
      
      return { content: [{ type: 'text', text: JSON.stringify(items, null, 2) }] };
    }
    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }
}

// Azure AD Roles Handler
export async function handleAzureAdRoles(
  graphClient: Client,
  args: AzureAdRoleArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'list_roles':
      apiPath = '/directoryRoles';
      if (args.filter) {
        apiPath += `?$filter=${encodeURIComponent(args.filter)}`;
      }
      result = await graphClient.api(apiPath).get();
      break;

    case 'list_role_assignments':
      // Note: Listing all role assignments requires Directory.Read.All
      // Filtering by principal requires RoleManagement.Read.Directory
      apiPath = '/roleManagement/directory/roleAssignments';
      if (args.filter) {
        // Example filter: $filter=principalId eq '{principalId}'
        apiPath += `?$filter=${encodeURIComponent(args.filter)}`;
      }
      result = await graphClient.api(apiPath).get();
      break;

    case 'assign_role':
      if (!args.roleId || !args.principalId) {
        throw new McpError(ErrorCode.InvalidParams, 'roleId and principalId are required for assign_role');
      }
      apiPath = '/roleManagement/directory/roleAssignments';
      const assignmentPayload = {
        '@odata.type': '#microsoft.graph.unifiedRoleAssignment',
        roleDefinitionId: args.roleId,
        principalId: args.principalId,
        directoryScopeId: '/', // Assign at tenant scope
      };
      result = await graphClient.api(apiPath).post(assignmentPayload);
      break;

    case 'remove_role_assignment':
      if (!args.assignmentId) {
        throw new McpError(ErrorCode.InvalidParams, 'assignmentId is required for remove_role_assignment');
      }
      apiPath = `/roleManagement/directory/roleAssignments/${args.assignmentId}`;
      await graphClient.api(apiPath).delete();
      result = { message: 'Role assignment removed successfully' };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}

// Azure AD Apps Handler
export async function handleAzureAdApps(
  graphClient: Client,
  args: AzureAdAppArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'list_apps':
      apiPath = '/applications';
      if (args.filter) {
        apiPath += `?$filter=${encodeURIComponent(args.filter)}`;
      }
      result = await graphClient.api(apiPath).get();
      break;

    case 'get_app':
      if (!args.appId) {
        throw new McpError(ErrorCode.InvalidParams, 'appId is required for get_app');
      }
      apiPath = `/applications/${args.appId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'update_app':
      if (!args.appId || !args.appDetails) {
        throw new McpError(ErrorCode.InvalidParams, 'appId and appDetails are required for update_app');
      }
      apiPath = `/applications/${args.appId}`;
      await graphClient.api(apiPath).patch(args.appDetails);
      result = { message: 'Application updated successfully' };
      break;

    case 'add_owner':
      if (!args.appId || !args.ownerId) {
        throw new McpError(ErrorCode.InvalidParams, 'appId and ownerId are required for add_owner');
      }
      apiPath = `/applications/${args.appId}/owners/$ref`;
      const ownerPayload = {
        '@odata.id': `https://graph.microsoft.com/v1.0/users/${args.ownerId}`
      };
      await graphClient.api(apiPath).post(ownerPayload);
      result = { message: 'Owner added successfully' };
      break;

    case 'remove_owner':
      if (!args.appId || !args.ownerId) {
        throw new McpError(ErrorCode.InvalidParams, 'appId and ownerId are required for remove_owner');
      }
      // Need to get the specific owner reference ID first, as Graph requires the owner's directoryObject ID from the owners collection
      // This is a simplification; a real implementation might need to list owners first to find the correct reference ID.
      // For now, we'll assume ownerId is the directoryObject ID of the owner within the app's owners collection.
      apiPath = `/applications/${args.appId}/owners/${args.ownerId}/$ref`;
      await graphClient.api(apiPath).delete();
      result = { message: 'Owner removed successfully' };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}

// Azure AD Devices Handler
export async function handleAzureAdDevices(
  graphClient: Client,
  args: AzureAdDeviceArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'list_devices':
      apiPath = '/devices';
      if (args.filter) {
        apiPath += `?$filter=${encodeURIComponent(args.filter)}`;
      }
      result = await graphClient.api(apiPath).get();
      break;

    case 'get_device':
      if (!args.deviceId) {
        throw new McpError(ErrorCode.InvalidParams, 'deviceId is required for get_device');
      }
      apiPath = `/devices/${args.deviceId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'enable_device':
    case 'disable_device':
      if (!args.deviceId) {
        throw new McpError(ErrorCode.InvalidParams, `deviceId is required for ${args.action}`);
      }
      // Note: Enabling/Disabling devices is done via update, setting accountEnabled
      // This requires Device.ReadWrite.All permission.
      apiPath = `/devices/${args.deviceId}`;
      await graphClient.api(apiPath).patch({
        accountEnabled: args.action === 'enable_device'
      });
      result = { message: `Device ${args.action === 'enable_device' ? 'enabled' : 'disabled'} successfully` };
      break;

    case 'delete_device':
      if (!args.deviceId) {
        throw new McpError(ErrorCode.InvalidParams, 'deviceId is required for delete_device');
      }
      // Requires Device.ReadWrite.All permission.
      apiPath = `/devices/${args.deviceId}`;
      await graphClient.api(apiPath).delete();
      result = { message: 'Device deleted successfully' };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}

// Service Principals Handler
export async function handleServicePrincipals(
  graphClient: Client,
  args: AzureAdSpArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'list_sps':
      apiPath = '/servicePrincipals';
      if (args.filter) {
        apiPath += `?$filter=${encodeURIComponent(args.filter)}`;
      }
      result = await graphClient.api(apiPath).get();
      break;

    case 'get_sp':
      if (!args.spId) {
        throw new McpError(ErrorCode.InvalidParams, 'spId is required for get_sp');
      }
      apiPath = `/servicePrincipals/${args.spId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'add_owner':
      if (!args.spId || !args.ownerId) {
        throw new McpError(ErrorCode.InvalidParams, 'spId and ownerId are required for add_owner');
      }
      // Requires Application.ReadWrite.All or Directory.ReadWrite.All
      apiPath = `/servicePrincipals/${args.spId}/owners/$ref`;
      const ownerPayload = {
        '@odata.id': `https://graph.microsoft.com/v1.0/users/${args.ownerId}`
      };
      await graphClient.api(apiPath).post(ownerPayload);
      result = { message: 'Owner added successfully to Service Principal' };
      break;

    case 'remove_owner':
      if (!args.spId || !args.ownerId) {
        throw new McpError(ErrorCode.InvalidParams, 'spId and ownerId are required for remove_owner');
      }
      // Requires Application.ReadWrite.All or Directory.ReadWrite.All
      // Similar to app owners, requires the directoryObject ID of the owner relationship
      apiPath = `/servicePrincipals/${args.spId}/owners/${args.ownerId}/$ref`;
      await graphClient.api(apiPath).delete();
      result = { message: 'Owner removed successfully from Service Principal' };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}

// Generic API Call Handler
export async function handleCallMicrosoftApi(
  graphClient: Client,
  args: CallMicrosoftApiArgs,
  getAccessToken: (scope: string) => Promise<string>,
  apiConfigs: any
): Promise<{ content: { type: string; text: string }[]; isError?: boolean }> {
  try {
    const { apiType, path, method, apiVersion, subscriptionId, queryParams, body } = args;

    if (apiType === 'azure' && !apiVersion) {
      throw new McpError(ErrorCode.InvalidParams, "apiVersion is required for apiType 'azure'");
    }

    const config = apiConfigs[apiType];
    const token = await getAccessToken(config.scope);

    let url = config.baseUrl;
    const urlParams = new URLSearchParams();

    // Construct URL based on API type
    if (apiType === 'azure') {
      let azurePath = path;
      // Prepend subscription if provided and path doesn't already include it
      if (subscriptionId && !path.toLowerCase().startsWith('/subscriptions/')) {
         azurePath = `/subscriptions/${subscriptionId}${path.startsWith('/') ? '' : '/'}${path}`;
      }
      url += azurePath;
      urlParams.append('api-version', apiVersion!); // Already validated
    } else { // graph
      url += path.startsWith('/') ? path : `/${path}`;
    }

    // Add query parameters
    if (queryParams) {
      for (const [key, value] of Object.entries(queryParams)) {
        urlParams.append(key, value);
      }
    }

    const queryString = urlParams.toString();
    if (queryString) {
      url += `?${queryString}`;
    }

    // Prepare request options
    const headers: Record<string, string> = {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    };
    if (apiType === 'graph') {
      headers['ConsistencyLevel'] = 'eventual'; // Good practice for Graph count/filter
    }

    const requestOptions: RequestInit = {
      method: method.toUpperCase(),
      headers: headers
    };

    if (["POST", "PUT", "PATCH"].includes(method.toUpperCase()) && body !== undefined) {
      requestOptions.body = typeof body === 'string' ? body : JSON.stringify(body);
    }

    // Make API request
    const apiResponse = await fetch(url, requestOptions);
    const responseText = await apiResponse.text();
    let responseData: any;

    try {
      responseData = responseText ? JSON.parse(responseText) : {};
    } catch (e) {
      // If not JSON, return raw text
      responseData = { rawResponse: responseText };
    }

    if (!apiResponse.ok) {
      console.error(`API Error (${apiResponse.status}) for ${method.toUpperCase()} ${url}:`, responseData);
      // Return error details in a structured way
       return {
         content: [{ type: 'text', text: JSON.stringify({
             error: `API Error (${apiResponse.status} ${apiResponse.statusText})`,
             url: url,
             details: responseData
           }, null, 2)
         }],
         isError: true
       };
    }

    // Successful response
    return {
      content: [{ type: 'text', text: JSON.stringify(responseData, null, 2) }]
    };

  } catch (error) {
     console.error("Error in handleCallMicrosoftApi:", error);
     const errorMessage = error instanceof Error ? error.message : String(error);
     // Ensure McpError is thrown correctly if it's already one
     if (error instanceof McpError) {
        throw error;
     }
     // Wrap other errors
     throw new McpError(ErrorCode.InternalError, `Failed to execute API call: ${errorMessage}`);
  }
}

// Security & Compliance Handlers
export async function handleSearchAuditLog(
  graphClient: Client,
  args: AuditLogArgs
): Promise<{ content: { type: string; text: string }[] }> {
  // Primarily targets /auditLogs/directoryAudits for now
  // Requires AuditLog.Read.All permission
  let apiPath = '/auditLogs/directoryAudits';
  const queryOptions: string[] = [];

  if (args.filter) {
    queryOptions.push(`$filter=${encodeURIComponent(args.filter)}`);
  }
  if (args.top) {
    queryOptions.push(`$top=${args.top}`);
  }

  if (queryOptions.length > 0) {
    apiPath += `?${queryOptions.join('&')}`;
  }

  const result = await graphClient.api(apiPath).get();
  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}

export async function handleManageAlerts(
  graphClient: Client,
  args: AlertArgs
): Promise<{ content: { type: string; text: string }[] }> {
  // Uses the newer alerts_v2 endpoint
  // Requires SecurityAlert.Read.All permission
  let apiPath = '/security/alerts_v2';
  let result: any;

  switch (args.action) {
    case 'list_alerts': {
      const queryOptions: string[] = [];
      if (args.filter) {
        queryOptions.push(`$filter=${encodeURIComponent(args.filter)}`);
      }
      if (args.top) {
        queryOptions.push(`$top=${args.top}`);
      }
      if (queryOptions.length > 0) {
        apiPath += `?${queryOptions.join('&')}`;
      }
      result = await graphClient.api(apiPath).get();
      break;
    }
    case 'get_alert': {
      if (!args.alertId) {
        throw new McpError(ErrorCode.InvalidParams, 'alertId is required for get_alert');
      }
      apiPath += `/${args.alertId}`;
      result = await graphClient.api(apiPath).get();
      break;
    }
    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}
