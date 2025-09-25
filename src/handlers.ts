import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';
import {
  UserManagementArgs,
  OffboardingArgs,
  DistributionListArgs,
  SecurityGroupArgs,
  M365GroupArgs,
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

// Import DLP handlers and types
import {
  handleDLPPolicies,
  handleDLPIncidents,
  handleDLPSensitivityLabels
} from './handlers/dlp-handler.js';
import {
  DLPPolicyArgs,
  DLPIncidentArgs,
  DLPSensitivityLabelArgs
} from './types/dlp-types.js';

// Import Intune macOS handlers and types
import {
  handleIntuneMacOSDevices,
  handleIntuneMacOSPolicies,
  handleIntuneMacOSApps,
  handleIntuneMacOSCompliance
} from './handlers/intune-macos-handler.js';
import {
  IntuneMacOSDeviceArgs,
  IntuneMacOSPolicyArgs,
  IntuneMacOSAppArgs,
  IntuneMacOSComplianceArgs
} from './types/intune-types.js';

// Import compliance handlers and types
import {
  handleComplianceFrameworks,
  handleComplianceAssessments,
  handleComplianceMonitoring,
  handleEvidenceCollection,
  handleGapAnalysis
} from './handlers/compliance-handler.js';
import {
  ComplianceFrameworkArgs,
  ComplianceAssessmentArgs,
  ComplianceMonitoringArgs,
  EvidenceCollectionArgs,
  GapAnalysisArgs
} from './types/compliance-types.js';

// Import audit reporting handler
import { handleAuditReports } from './handlers/audit-reporting-handler.js';
import { AuditReportArgs } from './types/compliance-types.js';

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
      let apiPath = '';
      
      if (args.siteId) {
        // Get specific site by ID
        apiPath = `/sites/${args.siteId}`;
      } else if (args.url) {
        // Get site by URL (hostname:path format)
        const urlParts = args.url.replace('https://', '').split('/');
        const hostname = urlParts[0];
        const sitePath = urlParts.slice(1).join('/') || 'sites/root';
        apiPath = `/sites/${hostname}:/${sitePath}`;
      } else {
        throw new McpError(ErrorCode.InvalidParams, 'Either siteId or url is required for get action');
      }
      
      const site = await graphClient.api(apiPath).get();
      return { content: [{ type: 'text', text: JSON.stringify(site, null, 2) }] };
    }
    
    case 'list': {
      // List all sites in the organization
      const sites = await graphClient
        .api('/sites?search=*')
        .get();
      return { content: [{ type: 'text', text: JSON.stringify(sites, null, 2) }] };
    }
    
    case 'search': {
      if (!args.title) {
        throw new McpError(ErrorCode.InvalidParams, 'title (search query) is required for search action');
      }
      
      const sites = await graphClient
        .api(`/sites?search=${encodeURIComponent(args.title)}`)
        .get();
      return { content: [{ type: 'text', text: JSON.stringify(sites, null, 2) }] };
    }
    
    case 'create': {
      // Note: Direct site creation via Graph API is limited
      // This creates a communication site via SharePoint REST API
      throw new McpError(
        ErrorCode.InvalidParams, 
        'Direct site creation is not supported via Graph API. Use SharePoint admin center or PowerShell for site creation. You can use the "get" action to retrieve existing sites.'
      );
    }
    
    case 'update': {
      if (!args.siteId) {
        throw new McpError(ErrorCode.InvalidParams, 'siteId is required for update action');
      }
      
      const updatePayload: any = {};
      if (args.title) updatePayload.displayName = args.title;
      if (args.description) updatePayload.description = args.description;
      
      // Update site properties (limited to displayName and description)
      const result = await graphClient
        .api(`/sites/${args.siteId}`)
        .patch(updatePayload);
      
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    
    case 'delete': {
      throw new McpError(
        ErrorCode.InvalidParams, 
        'Site deletion is not supported via Graph API. Use SharePoint admin center or PowerShell for site deletion.'
      );
    }
    
    case 'get_permissions': {
      if (!args.siteId) {
        throw new McpError(ErrorCode.InvalidParams, 'siteId is required for get_permissions action');
      }
      
      const permissions = await graphClient
        .api(`/sites/${args.siteId}/permissions`)
        .get();
      return { content: [{ type: 'text', text: JSON.stringify(permissions, null, 2) }] };
    }
    
    case 'get_drives': {
      if (!args.siteId) {
        throw new McpError(ErrorCode.InvalidParams, 'siteId is required for get_drives action');
      }
      
      const drives = await graphClient
        .api(`/sites/${args.siteId}/drives`)
        .get();
      return { content: [{ type: 'text', text: JSON.stringify(drives, null, 2) }] };
    }
    
    case 'get_subsites': {
      if (!args.siteId) {
        throw new McpError(ErrorCode.InvalidParams, 'siteId is required for get_subsites action');
      }
      
      const subsites = await graphClient
        .api(`/sites/${args.siteId}/sites`)
        .get();
      return { content: [{ type: 'text', text: JSON.stringify(subsites, null, 2) }] };
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

// Generic API Call Handler - Enhanced with performance and reliability features
export async function handleCallMicrosoftApi(
  graphClient: Client,
  args: CallMicrosoftApiArgs,
  getAccessToken: (scope: string) => Promise<string>,
  apiConfigs: any,
  rateLimiter?: any,
  tokenCache?: any
): Promise<{ content: { type: string; text: string }[]; isError?: boolean }> {
  const startTime = Date.now();
  
  // Extract parameters with defaults
  const { 
    apiType, 
    path, 
    method, 
    apiVersion, 
    subscriptionId, 
    queryParams = {}, 
    body, 
    graphApiVersion = 'v1.0', 
    fetchAll = false,
    consistencyLevel,
    maxRetries = 3,
    retryDelay = 1000,
    timeout = 30000,
    customHeaders = {},
    responseFormat = 'json',
    selectFields,
    expandFields,
    batchSize = 100
  } = args;

  
  // Apply rate limiting if available
  if (rateLimiter) {
    await rateLimiter.checkLimit();
  }

  let determinedUrl: string | undefined;

  // Enhanced token caching helper
  const getTokenWithCache = async (scope: string): Promise<string> => {
    if (tokenCache) {
      const cached = tokenCache.get(scope);
      if (cached) {
        return cached;
      }
    }
    
    const token = await getAccessToken(scope);
    
    if (tokenCache) {
      tokenCache.set(scope, token);
    }
    
    return token;
  };

  // Auto-apply selectFields and expandFields for Graph API
  if (apiType === 'graph') {
    if (selectFields && selectFields.length > 0) {
      queryParams['$select'] = selectFields.join(',');
    }
    if (expandFields && expandFields.length > 0) {
      queryParams['$expand'] = expandFields.join(',');
    }
    if (fetchAll && batchSize !== 100) {
      queryParams['$top'] = batchSize.toString();
    }
  }

  // Retry logic wrapper with exponential backoff
  const executeWithRetry = async (operation: () => Promise<any>): Promise<any> => {
    let lastError: any;
    
    for (let attempt = 0; attempt <= maxRetries; attempt++) {
      try {
        if (attempt > 0) {
          const delay = retryDelay * Math.pow(2, attempt - 1); // Exponential backoff
          console.debug(`Retry attempt ${attempt}/${maxRetries}, waiting ${delay}ms`);
          await new Promise(resolve => setTimeout(resolve, delay));
        }
        
        return await operation();
      } catch (error: any) {
        lastError = error;
        
        // Don't retry on authentication errors or 4xx client errors (except 429)
        if (error.statusCode && error.statusCode >= 400 && error.statusCode < 500 && error.statusCode !== 429) {
          throw error;
        }
        
        if (attempt === maxRetries) {
          console.error(`All retry attempts exhausted for ${apiType} ${method} ${path}`);
          throw error;
        }
        
        console.warn(`Attempt ${attempt + 1} failed, will retry:`, error.message);
      }
    }
    
    throw lastError;
  };

  try {
    if (apiType === 'azure' && !apiVersion) {
      throw new McpError(ErrorCode.InvalidParams, "apiVersion is required for apiType 'azure'");
    }

    let responseData: any;

    // --- Microsoft Graph Logic (Enhanced) ---
    if (apiType === 'graph') {
      determinedUrl = `https://graph.microsoft.com/${graphApiVersion}`;
      
      responseData = await executeWithRetry(async () => {
        let request = graphClient.api(path).version(graphApiVersion);
        
        // Add query parameters if provided
        if (Object.keys(queryParams).length > 0) {
          request = request.query(queryParams);
        }
        
        // Add ConsistencyLevel header if provided
        if (consistencyLevel) {
          request = request.header('ConsistencyLevel', consistencyLevel);
        }
          // Add custom headers
        Object.entries(customHeaders).forEach(([key, value]) => {
          request = request.header(key, value);
        });

        // Note: Graph SDK doesn't support timeout directly, but we can implement AbortController for timeout
        
        // Handle different methods
        switch (method.toLowerCase()) {
          case 'get':
            if (fetchAll) {
              // Initialize with empty array for collecting all items
              let allItems: any[] = [];
              let nextLink: string | null | undefined = null;
              
              // Get first page
              const firstPageResponse = await request.get();
              
              // Store context from first page
              const odataContext = firstPageResponse['@odata.context'];
              
              // Add items from first page
              if (firstPageResponse.value && Array.isArray(firstPageResponse.value)) {
                allItems = [...firstPageResponse.value];
              }
              
              // Get nextLink from first page
              nextLink = firstPageResponse['@odata.nextLink'];
              
              // Fetch subsequent pages
              while (nextLink) {
                  // Create a new request for the next page
                const nextPageResponse = await graphClient.api(nextLink).get();
                
                // Add items from next page
                if (nextPageResponse.value && Array.isArray(nextPageResponse.value)) {
                  allItems = [...allItems, ...nextPageResponse.value];
                }
                
                // Update nextLink
                nextLink = nextPageResponse['@odata.nextLink'];
              }
              
              // Construct final response
              return {
                '@odata.context': odataContext,
                value: allItems,
                totalCount: allItems.length,
                fetchedAt: new Date().toISOString()
              };
            } else {
              return await request.get();
            }
          case 'post':
            return await request.post(body ?? {});
          case 'put':
            return await request.put(body ?? {});
          case 'patch':
            return await request.patch(body ?? {});
          case 'delete':
            const deleteResult = await request.delete();
            // Handle potential 204 No Content response
            return deleteResult === undefined || deleteResult === null 
              ? { status: "Success (No Content)", deletedAt: new Date().toISOString() } 
              : deleteResult;
          default:
            throw new Error(`Unsupported method: ${method}`);
        }
      });
    }
    // --- Azure Resource Management Logic (Enhanced) ---
    else { // apiType === 'azure'
      determinedUrl = "https://management.azure.com";
      
      responseData = await executeWithRetry(async () => {
        const token = await getTokenWithCache("https://management.azure.com/.default");
        
        let url = determinedUrl!;
        if (subscriptionId) {
          url += `/subscriptions/${subscriptionId}`;
        }
        url += path.startsWith('/') ? path : `/${path}`;
        
        const urlParams = new URLSearchParams();
        urlParams.append('api-version', apiVersion!);
        
        Object.entries(queryParams).forEach(([key, value]) => {
          urlParams.append(key, value);
        });
        
        url += `?${urlParams.toString()}`;
        
        // Prepare request options
        const headers: Record<string, string> = {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
          ...customHeaders
        };
        
        const requestOptions: RequestInit = {
          method: method.toUpperCase(),
          headers: headers,
          signal: AbortSignal.timeout(timeout)
        };
        
        if (["POST", "PUT", "PATCH"].includes(method.toUpperCase()) && body !== undefined) {
          requestOptions.body = typeof body === 'string' ? body : JSON.stringify(body);
        }
        
        // --- Pagination Logic for Azure RM ---
        if (fetchAll && method.toLowerCase() === 'get') {
          console.debug(`Fetching all pages for Azure RM starting from: ${url}`);
          
          let allValues: any[] = [];
          let currentUrl: string | null = url;
          
          while (currentUrl) {
            console.debug(`Fetching Azure RM page: ${currentUrl}`);
            
            // Re-acquire token for each page (Azure tokens might expire)
            const currentPageToken = await getTokenWithCache("https://management.azure.com/.default");
            const currentPageHeaders = { ...headers, 'Authorization': `Bearer ${currentPageToken}` };
            const currentPageRequestOptions: RequestInit = { 
              method: 'GET', 
              headers: currentPageHeaders,
              signal: AbortSignal.timeout(timeout)
            };
            
            const pageResponse = await fetch(currentUrl, currentPageRequestOptions);
            const pageText = await pageResponse.text();
            
            let pageData: any;
            try {
              pageData = pageText ? JSON.parse(pageText) : {};
            } catch (e) {
              console.error(`Failed to parse JSON from Azure RM page: ${currentUrl}`, pageText);
              pageData = { rawResponse: pageText };
            }
            
            if (!pageResponse.ok) {
              console.error(`API error on Azure RM page ${currentUrl}:`, pageData);
              throw new Error(`API error (${pageResponse.status}) during Azure RM pagination on ${currentUrl}: ${JSON.stringify(pageData)}`);
            }
            
            if (pageData.value && Array.isArray(pageData.value)) {
              allValues = allValues.concat(pageData.value);
            } else if (currentUrl === url && !pageData.nextLink) {
              // If this is the first page and there's no nextLink, it might be a single resource
              allValues.push(pageData);
            }
            
            currentUrl = pageData.nextLink || null; // Azure uses nextLink
          }
          
          return { 
            value: allValues, 
            totalCount: allValues.length,
            fetchedAt: new Date().toISOString()
          };
        } else {
          // Single page fetch for Azure RM
          console.debug(`Fetching single page for Azure RM: ${url}`);
          
          const apiResponse = await fetch(url, requestOptions);
          const responseText = await apiResponse.text();
          
          try {
            const data = responseText ? JSON.parse(responseText) : {};
            if (!apiResponse.ok) {
              throw new Error(`API error (${apiResponse.status}): ${JSON.stringify(data)}`);
            }
            return data;
          } catch (e) {
            if (!apiResponse.ok) {
              throw new Error(`API error (${apiResponse.status}): ${responseText}`);
            }
            return { rawResponse: responseText };
          }
        }
      });
    }

    // --- Enhanced Response Formatting ---
    const executionTime = Date.now() - startTime;
    let resultText = "";

    switch (responseFormat) {
      case "minimal":
        if (responseData && responseData.value && Array.isArray(responseData.value)) {
          resultText = JSON.stringify(responseData.value, null, 2);
        } else if (responseData && typeof responseData === 'object') {
          // Extract just the data, excluding metadata
          const { '@odata.context': _, '@odata.nextLink': __, ...cleanData } = responseData;
          resultText = JSON.stringify(cleanData, null, 2);
        } else {
          resultText = JSON.stringify(responseData, null, 2);
        }
        break;
      case "raw":
        resultText = JSON.stringify(responseData);
        break;
      default: // "json"
        resultText = `Result for ${apiType} API (${apiType === 'graph' ? graphApiVersion : apiVersion}) - ${method.toUpperCase()} ${path}:\n`;
        resultText += `Execution time: ${executionTime}ms\n`;
        if (fetchAll && responseData.totalCount !== undefined) {
          resultText += `Total items fetched: ${responseData.totalCount}\n`;
        }
        resultText += `\n${JSON.stringify(responseData, null, 2)}`;
        break;
    }

    // Add pagination note for single-page requests
    if (!fetchAll && method.toLowerCase() === 'get' && responseFormat === 'json') {
      const nextLinkKey = apiType === 'graph' ? '@odata.nextLink' : 'nextLink';
      if (responseData && responseData[nextLinkKey]) {
        resultText += `\n\nNote: More results are available. To retrieve all pages, add 'fetchAll: true' to your request.`;
      }
    }

    return {
      content: [{ type: "text", text: resultText }],
    };

  } catch (error) {
    const executionTime = Date.now() - startTime;
    console.error(`Error in enhanced Microsoft API call (apiType: ${apiType}, path: ${path}, method: ${method}, executionTime: ${executionTime}ms):`, error);
    
    // Try to determine the base URL even in case of error
    if (!determinedUrl) {
      determinedUrl = apiType === 'graph'
        ? `https://graph.microsoft.com/${graphApiVersion}`
        : "https://management.azure.com";
    }
    
    // Include error body if available
    let errorBody = 'N/A';
    let statusCode = 'N/A';
    
    // Type guard for error object with body property
    if (error && typeof error === 'object') {
      if ('body' in error) {
        const body = (error as any).body;
        errorBody = typeof body === 'string' ? body : JSON.stringify(body);
      }
      
      if ('statusCode' in error) {
        statusCode = String((error as any).statusCode);
      }
    }
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          error: error instanceof Error ? error.message : String(error),
          statusCode: statusCode,
          errorBody: errorBody,
          attemptedBaseUrl: determinedUrl,
          executionTime: executionTime,
          retryAttempts: maxRetries,
          timestamp: new Date().toISOString()
        }, null, 2),
      }],
      isError: true
    };
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

// Distribution Lists Handler
export async function handleDistributionLists(
  graphClient: Client,
  args: DistributionListArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'get':
      if (!args.listId) {
        throw new McpError(ErrorCode.InvalidParams, 'listId is required for get action');
      }
      apiPath = `/groups/${args.listId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'create':
      if (!args.displayName || !args.emailAddress) {
        throw new McpError(ErrorCode.InvalidParams, 'displayName and emailAddress are required for create action');
      }
      apiPath = '/groups';
      const createPayload = {
        displayName: args.displayName,
        mailNickname: args.emailAddress.split('@')[0],
        mailEnabled: true,
        securityEnabled: false,
        groupTypes: [], // Empty for distribution lists
        mail: args.emailAddress
      };
      result = await graphClient.api(apiPath).post(createPayload);
      break;

    case 'update':
      if (!args.listId) {
        throw new McpError(ErrorCode.InvalidParams, 'listId is required for update action');
      }
      apiPath = `/groups/${args.listId}`;
      const updatePayload: any = {};
      if (args.displayName) updatePayload.displayName = args.displayName;
      result = await graphClient.api(apiPath).patch(updatePayload);
      break;

    case 'delete':
      if (!args.listId) {
        throw new McpError(ErrorCode.InvalidParams, 'listId is required for delete action');
      }
      apiPath = `/groups/${args.listId}`;
      await graphClient.api(apiPath).delete();
      result = { message: 'Distribution list deleted successfully' };
      break;

    case 'add_members':
      if (!args.listId || !args.members?.length) {
        throw new McpError(ErrorCode.InvalidParams, 'listId and members are required for add_members action');
      }
      
      for (const member of args.members) {
        await graphClient
          .api(`/groups/${args.listId}/members/$ref`)
          .post({
            '@odata.id': `https://graph.microsoft.com/v1.0/users/${member}`
          });
      }
      result = { message: `Added ${args.members.length} members to distribution list` };
      break;

    case 'remove_members':
      if (!args.listId || !args.members?.length) {
        throw new McpError(ErrorCode.InvalidParams, 'listId and members are required for remove_members action');
      }
      
      for (const member of args.members) {
        await graphClient
          .api(`/groups/${args.listId}/members/${member}/$ref`)
          .delete();
      }
      result = { message: `Removed ${args.members.length} members from distribution list` };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}

// Security Groups Handler
export async function handleSecurityGroups(
  graphClient: Client,
  args: SecurityGroupArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'get':
      if (!args.groupId) {
        throw new McpError(ErrorCode.InvalidParams, 'groupId is required for get action');
      }
      apiPath = `/groups/${args.groupId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'create':
      if (!args.displayName) {
        throw new McpError(ErrorCode.InvalidParams, 'displayName is required for create action');
      }
      apiPath = '/groups';
      const createPayload = {
        displayName: args.displayName,
        description: args.description || '',
        mailNickname: args.displayName.replace(/\s+/g, '').toLowerCase(),
        mailEnabled: args.settings?.mailEnabled || false,
        securityEnabled: args.settings?.securityEnabled !== false, // Default to true
        groupTypes: []
      };
      result = await graphClient.api(apiPath).post(createPayload);
      break;

    case 'update':
      if (!args.groupId) {
        throw new McpError(ErrorCode.InvalidParams, 'groupId is required for update action');
      }
      apiPath = `/groups/${args.groupId}`;
      const updatePayload: any = {};
      if (args.displayName) updatePayload.displayName = args.displayName;
      if (args.description) updatePayload.description = args.description;
      result = await graphClient.api(apiPath).patch(updatePayload);
      break;

    case 'delete':
      if (!args.groupId) {
        throw new McpError(ErrorCode.InvalidParams, 'groupId is required for delete action');
      }
      apiPath = `/groups/${args.groupId}`;
      await graphClient.api(apiPath).delete();
      result = { message: 'Security group deleted successfully' };
      break;

    case 'add_members':
      if (!args.groupId || !args.members?.length) {
        throw new McpError(ErrorCode.InvalidParams, 'groupId and members are required for add_members action');
      }
      
      for (const member of args.members) {
        await graphClient
          .api(`/groups/${args.groupId}/members/$ref`)
          .post({
            '@odata.id': `https://graph.microsoft.com/v1.0/users/${member}`
          });
      }
      result = { message: `Added ${args.members.length} members to security group` };
      break;

    case 'remove_members':
      if (!args.groupId || !args.members?.length) {
        throw new McpError(ErrorCode.InvalidParams, 'groupId and members are required for remove_members action');
      }
      
      for (const member of args.members) {
        await graphClient
          .api(`/groups/${args.groupId}/members/${member}/$ref`)
          .delete();
      }
      result = { message: `Removed ${args.members.length} members from security group` };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}

// M365 Groups Handler
export async function handleM365Groups(
  graphClient: Client,
  args: M365GroupArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'get':
      if (!args.groupId) {
        throw new McpError(ErrorCode.InvalidParams, 'groupId is required for get action');
      }
      apiPath = `/groups/${args.groupId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'create':
      if (!args.displayName) {
        throw new McpError(ErrorCode.InvalidParams, 'displayName is required for create action');
      }
      apiPath = '/groups';
      const createPayload = {
        displayName: args.displayName,
        description: args.description || '',
        mailNickname: args.displayName.replace(/\s+/g, '').toLowerCase(),
        mailEnabled: true,
        securityEnabled: false,
        groupTypes: ['Unified'], // M365 groups are unified groups
        visibility: args.settings?.visibility || 'Private'
      };
      result = await graphClient.api(apiPath).post(createPayload);
      
      // Add owners if provided
      if (args.owners?.length && result.id) {
        for (const owner of args.owners) {
          await graphClient
            .api(`/groups/${result.id}/owners/$ref`)
            .post({
              '@odata.id': `https://graph.microsoft.com/v1.0/users/${owner}`
            });
        }
      }
      break;

    case 'update':
      if (!args.groupId) {
        throw new McpError(ErrorCode.InvalidParams, 'groupId is required for update action');
      }
      apiPath = `/groups/${args.groupId}`;
      const updatePayload: any = {};
      if (args.displayName) updatePayload.displayName = args.displayName;
      if (args.description) updatePayload.description = args.description;
      if (args.settings?.visibility) updatePayload.visibility = args.settings.visibility;
      result = await graphClient.api(apiPath).patch(updatePayload);
      break;

    case 'delete':
      if (!args.groupId) {
        throw new McpError(ErrorCode.InvalidParams, 'groupId is required for delete action');
      }
      apiPath = `/groups/${args.groupId}`;
      await graphClient.api(apiPath).delete();
      result = { message: 'M365 group deleted successfully' };
      break;

    case 'add_members':
      if (!args.groupId || !args.members?.length) {
        throw new McpError(ErrorCode.InvalidParams, 'groupId and members are required for add_members action');
      }
      
      for (const member of args.members) {
        await graphClient
          .api(`/groups/${args.groupId}/members/$ref`)
          .post({
            '@odata.id': `https://graph.microsoft.com/v1.0/users/${member}`
          });
      }
      result = { message: `Added ${args.members.length} members to M365 group` };
      break;

    case 'remove_members':
      if (!args.groupId || !args.members?.length) {
        throw new McpError(ErrorCode.InvalidParams, 'groupId and members are required for remove_members action');
      }
      
      for (const member of args.members) {
        await graphClient
          .api(`/groups/${args.groupId}/members/${member}/$ref`)
          .delete();
      }
      result = { message: `Removed ${args.members.length} members from M365 group` };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}

export async function handleSharePointSites(
  graphClient: Client,
  args: SharePointSiteArgs
): Promise<{ content: { type: string; text: string }[] }> {
  // Delegate to the main SharePoint site handler
  return await handleSharePointSite(graphClient, args);
}

export async function handleSharePointLists(
  graphClient: Client,
  args: SharePointListArgs
): Promise<{ content: { type: string; text: string }[] }> {
  // Delegate to the main SharePoint list handler
  return await handleSharePointList(graphClient, args);
}
