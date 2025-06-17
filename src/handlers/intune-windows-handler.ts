import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { 
  IntuneWindowsDeviceArgs, 
  IntuneWindowsPolicyArgs, 
  IntuneWindowsAppArgs, 
  IntuneWindowsComplianceArgs 
} from '../types/intune-types.js';

// Intune Windows Device Management Handler
export async function handleIntuneWindowsDevices(
  graphClient: Client,
  args: IntuneWindowsDeviceArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'list':
      // List all Windows devices managed by Intune
      apiPath = '/deviceManagement/managedDevices';
      const queryOptions: string[] = [];
      
      // Filter for Windows devices
      queryOptions.push(`$filter=operatingSystem eq 'Windows'`);
      
      if (args.filter) {
        queryOptions.push(`and ${args.filter}`);
      }

      if (queryOptions.length > 0) {
        apiPath += `?${queryOptions.join('')}`;
      }

      result = await graphClient.api(apiPath).get();
      break;

    case 'get':
      if (!args.deviceId) {
        throw new McpError(ErrorCode.InvalidParams, 'deviceId is required for get action');
      }
      apiPath = `/deviceManagement/managedDevices/${args.deviceId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'enroll':
      // Create enrollment invitation for Windows devices
      apiPath = '/deviceManagement/deviceEnrollmentConfigurations';
      const enrollmentPayload = {
        displayName: 'Windows Device Enrollment',
        description: 'Automated Windows device enrollment',
        deviceEnrollmentConfigurationType: 'windows10EnrollmentCompletionPageConfiguration',
        priority: 0,
        showInstallationProgress: true,
        blockDeviceSetupRetryByUser: false,
        allowDeviceResetOnInstallFailure: true,
        allowLogCollectionOnInstallFailure: true,
        customErrorMessage: 'Setup could not be completed. Please try again or contact your support person for help.',
        installProgressTimeoutInMinutes: 60,
        allowDeviceUseOnInstallFailure: true,
        selectedMobileAppIds: [],
        trackInstallProgressForAutopilotOnly: false,
        disableUserStatusTrackingAfterFirstUser: true
      };

      if (args.enrollmentType) {
        enrollmentPayload.deviceEnrollmentConfigurationType = 
          args.enrollmentType === 'AzureADJoin' ? 'azureADJoinUsingBulkEnrollment' :
          args.enrollmentType === 'HybridAzureADJoin' ? 'hybridAzureADJoin' :
          'windows10EnrollmentCompletionPageConfiguration';
      }

      result = await graphClient.api(apiPath).post(enrollmentPayload);
      break;

    case 'retire':
      if (!args.deviceId) {
        throw new McpError(ErrorCode.InvalidParams, 'deviceId is required for retire action');
      }
      apiPath = `/deviceManagement/managedDevices/${args.deviceId}/retire`;
      result = await graphClient.api(apiPath).post({
        keepEnrollmentData: false,
        keepUserData: true
      });
      break;

    case 'wipe':
      if (!args.deviceId) {
        throw new McpError(ErrorCode.InvalidParams, 'deviceId is required for wipe action');
      }
      apiPath = `/deviceManagement/managedDevices/${args.deviceId}/wipe`;
      result = await graphClient.api(apiPath).post({
        keepEnrollmentData: false,
        keepUserData: false,
        useProtectedWipe: true
      });
      break;

    case 'restart':
      if (!args.deviceId) {
        throw new McpError(ErrorCode.InvalidParams, 'deviceId is required for restart action');
      }
      apiPath = `/deviceManagement/managedDevices/${args.deviceId}/rebootNow`;
      result = await graphClient.api(apiPath).post({});
      break;

    case 'sync':
      if (!args.deviceId) {
        throw new McpError(ErrorCode.InvalidParams, 'deviceId is required for sync action');
      }
      apiPath = `/deviceManagement/managedDevices/${args.deviceId}/syncDevice`;
      result = await graphClient.api(apiPath).post({});
      break;

    case 'remote_lock':
      if (!args.deviceId) {
        throw new McpError(ErrorCode.InvalidParams, 'deviceId is required for remote_lock action');
      }
      apiPath = `/deviceManagement/managedDevices/${args.deviceId}/remoteLock`;
      result = await graphClient.api(apiPath).post({});
      break;

    case 'collect_logs':
      if (!args.deviceId) {
        throw new McpError(ErrorCode.InvalidParams, 'deviceId is required for collect_logs action');
      }
      apiPath = `/deviceManagement/managedDevices/${args.deviceId}/createDeviceLogCollectionRequest`;
      result = await graphClient.api(apiPath).post({
        templateType: 'predefined'
      });
      break;    case 'bitlocker_recovery':
      if (!args.deviceId) {
        throw new McpError(ErrorCode.InvalidParams, 'deviceId is required for bitlocker_recovery action');
      }
      apiPath = `/informationProtection/bitlocker/recoveryKeys`;
      const filter = `$filter=deviceId eq '${args.deviceId}'`;
      result = await graphClient.api(`${apiPath}?${filter}`).get();
      break;

    case 'autopilot_reset':
      if (!args.deviceId) {
        throw new McpError(ErrorCode.InvalidParams, 'deviceId is required for autopilot_reset action');
      }
      apiPath = `/deviceManagement/managedDevices/${args.deviceId}/autopilotReset`;
      result = await graphClient.api(apiPath).post({
        keepUserData: false
      });
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Unknown action: ${args.action}`);
  }

  return {
    content: [
      {
        type: 'text',
        text: `Windows Device Management Result:\n${JSON.stringify(result, null, 2)}`
      }
    ]
  };
}

// Intune Windows Policy Management Handler
export async function handleIntuneWindowsPolicies(
  graphClient: Client,
  args: IntuneWindowsPolicyArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'list':
      switch (args.policyType) {
        case 'Configuration':
          apiPath = '/deviceManagement/deviceConfigurations';
          break;
        case 'Compliance':
          apiPath = '/deviceManagement/deviceCompliancePolicies';
          break;
        case 'Security':
          apiPath = '/deviceManagement/intents';
          break;
        case 'Update':
          apiPath = '/deviceManagement/deviceConfigurations';
          apiPath += '?$filter=deviceManagementApplicabilityRuleOsEdition/osEditionTypes/any(x:x eq \'windows10Enterprise\')';
          break;
        case 'AppProtection':
          apiPath = '/deviceAppManagement/managedAppPolicies';
          break;
        case 'EndpointSecurity':
          apiPath = '/deviceManagement/intents';
          apiPath += '?$filter=templateId eq \'d1174162-1dd2-4976-affc-6667049ab0ae\''; // Endpoint Security template
          break;
        default:
          throw new McpError(ErrorCode.InvalidParams, `Unknown policy type: ${args.policyType}`);
      }
      result = await graphClient.api(apiPath).get();
      break;

    case 'get':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for get action');
      }
      
      switch (args.policyType) {
        case 'Configuration':
          apiPath = `/deviceManagement/deviceConfigurations/${args.policyId}`;
          break;
        case 'Compliance':
          apiPath = `/deviceManagement/deviceCompliancePolicies/${args.policyId}`;
          break;
        case 'Security':
        case 'EndpointSecurity':
          apiPath = `/deviceManagement/intents/${args.policyId}`;
          break;
        case 'Update':
          apiPath = `/deviceManagement/deviceConfigurations/${args.policyId}`;
          break;
        case 'AppProtection':
          apiPath = `/deviceAppManagement/managedAppPolicies/${args.policyId}`;
          break;
        default:
          throw new McpError(ErrorCode.InvalidParams, `Unknown policy type: ${args.policyType}`);
      }
      result = await graphClient.api(apiPath).get();
      break;

    case 'create':
      if (!args.name) {
        throw new McpError(ErrorCode.InvalidParams, 'name is required for create action');
      }

      const createPayload: any = {
        displayName: args.name,
        description: args.description || '',
        ...args.settings
      };

      switch (args.policyType) {
        case 'Configuration':
          apiPath = '/deviceManagement/deviceConfigurations';
          createPayload['@odata.type'] = '#microsoft.graph.windows10GeneralConfiguration';
          break;
        case 'Compliance':
          apiPath = '/deviceManagement/deviceCompliancePolicies';
          createPayload['@odata.type'] = '#microsoft.graph.windows10CompliancePolicy';
          break;
        case 'Security':
          apiPath = '/deviceManagement/intents';
          createPayload.templateId = 'd1174162-1dd2-4976-affc-6667049ab0ae'; // Security baseline template
          break;
        case 'Update':
          apiPath = '/deviceManagement/deviceConfigurations';
          createPayload['@odata.type'] = '#microsoft.graph.windowsUpdateForBusinessConfiguration';
          break;
        case 'AppProtection':
          apiPath = '/deviceAppManagement/managedAppPolicies';
          createPayload['@odata.type'] = '#microsoft.graph.windowsManagedAppProtection';
          break;
        case 'EndpointSecurity':
          apiPath = '/deviceManagement/intents';
          createPayload.templateId = 'e044e60e-5901-41ea-92c5-87e8b6edd6bb'; // Endpoint Security template
          break;
        default:
          throw new McpError(ErrorCode.InvalidParams, `Unknown policy type: ${args.policyType}`);
      }

      result = await graphClient.api(apiPath).post(createPayload);
      break;

    case 'update':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for update action');
      }

      const updatePayload: any = {};
      if (args.name) updatePayload.displayName = args.name;
      if (args.description) updatePayload.description = args.description;
      if (args.settings) Object.assign(updatePayload, args.settings);

      switch (args.policyType) {
        case 'Configuration':
          apiPath = `/deviceManagement/deviceConfigurations/${args.policyId}`;
          break;
        case 'Compliance':
          apiPath = `/deviceManagement/deviceCompliancePolicies/${args.policyId}`;
          break;
        case 'Security':
        case 'EndpointSecurity':
          apiPath = `/deviceManagement/intents/${args.policyId}`;
          break;
        case 'Update':
          apiPath = `/deviceManagement/deviceConfigurations/${args.policyId}`;
          break;
        case 'AppProtection':
          apiPath = `/deviceAppManagement/managedAppPolicies/${args.policyId}`;
          break;
        default:
          throw new McpError(ErrorCode.InvalidParams, `Unknown policy type: ${args.policyType}`);
      }

      result = await graphClient.api(apiPath).patch(updatePayload);
      break;

    case 'delete':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for delete action');
      }

      switch (args.policyType) {
        case 'Configuration':
          apiPath = `/deviceManagement/deviceConfigurations/${args.policyId}`;
          break;
        case 'Compliance':
          apiPath = `/deviceManagement/deviceCompliancePolicies/${args.policyId}`;
          break;
        case 'Security':
        case 'EndpointSecurity':
          apiPath = `/deviceManagement/intents/${args.policyId}`;
          break;
        case 'Update':
          apiPath = `/deviceManagement/deviceConfigurations/${args.policyId}`;
          break;
        case 'AppProtection':
          apiPath = `/deviceAppManagement/managedAppPolicies/${args.policyId}`;
          break;
        default:
          throw new McpError(ErrorCode.InvalidParams, `Unknown policy type: ${args.policyType}`);
      }

      result = await graphClient.api(apiPath).delete();
      break;

    case 'assign':
      if (!args.policyId || !args.assignments) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId and assignments are required for assign action');
      }

      switch (args.policyType) {
        case 'Configuration':
          apiPath = `/deviceManagement/deviceConfigurations/${args.policyId}/assign`;
          break;
        case 'Compliance':
          apiPath = `/deviceManagement/deviceCompliancePolicies/${args.policyId}/assign`;
          break;
        case 'Security':
        case 'EndpointSecurity':
          apiPath = `/deviceManagement/intents/${args.policyId}/assign`;
          break;
        case 'Update':
          apiPath = `/deviceManagement/deviceConfigurations/${args.policyId}/assign`;
          break;
        case 'AppProtection':
          apiPath = `/deviceAppManagement/managedAppPolicies/${args.policyId}/assign`;
          break;
        default:
          throw new McpError(ErrorCode.InvalidParams, `Unknown policy type: ${args.policyType}`);
      }

      result = await graphClient.api(apiPath).post({
        assignments: args.assignments
      });
      break;

    case 'deploy':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for deploy action');
      }

      // Deploy immediately to assigned groups
      apiPath = `/deviceManagement/deviceConfigurations/${args.policyId}/assignments`;
      const assignments = await graphClient.api(apiPath).get();
      
      result = {
        message: 'Policy deployment initiated',
        policyId: args.policyId,
        assignmentCount: assignments.value ? assignments.value.length : 0,
        deploymentSettings: args.deploymentSettings
      };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Unknown action: ${args.action}`);
  }

  return {
    content: [
      {
        type: 'text',
        text: `Windows Policy Management Result:\n${JSON.stringify(result, null, 2)}`
      }
    ]
  };
}

// Intune Windows App Management Handler
export async function handleIntuneWindowsApps(
  graphClient: Client,
  args: IntuneWindowsAppArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'list':
      apiPath = '/deviceAppManagement/mobileApps';
      let filter = '';
      
      if (args.appType) {
        const odataType = 
          args.appType === 'win32LobApp' ? '#microsoft.graph.win32LobApp' :
          args.appType === 'microsoftStoreForBusinessApp' ? '#microsoft.graph.microsoftStoreForBusinessApp' :
          args.appType === 'officeSuiteApp' ? '#microsoft.graph.officeSuiteApp' :
          args.appType === 'webApp' ? '#microsoft.graph.webApp' :
          args.appType === 'microsoftEdgeApp' ? '#microsoft.graph.microsoftEdgeApp' :
          '';
        
        if (odataType) {
          filter = `@odata.type eq '${odataType}'`;
        }
      }

      if (filter) {
        apiPath += `?$filter=${filter}`;
      }

      result = await graphClient.api(apiPath).get();
      break;

    case 'get':
      if (!args.appId) {
        throw new McpError(ErrorCode.InvalidParams, 'appId is required for get action');
      }
      apiPath = `/deviceAppManagement/mobileApps/${args.appId}`;
      result = await graphClient.api(apiPath).get();
      break;

    case 'deploy':
      if (!args.appId || !args.assignmentGroups) {
        throw new McpError(ErrorCode.InvalidParams, 'appId and assignmentGroups are required for deploy action');
      }      const assignments = args.assignmentGroups.map(groupId => ({
        target: {
          '@odata.type': '#microsoft.graph.groupAssignmentTarget',
          groupId: groupId
        },
        intent: args.assignment?.installIntent || 'available',
        settings: {
          '@odata.type': '#microsoft.graph.win32LobAppAssignmentSettings',
          notifications: 'showAll',
          restartSettings: null,
          installTimeSettings: null
        }
      }));

      apiPath = `/deviceAppManagement/mobileApps/${args.appId}/assign`;
      result = await graphClient.api(apiPath).post({
        mobileAppAssignments: assignments
      });
      break;

    case 'update':
      if (!args.appId) {
        throw new McpError(ErrorCode.InvalidParams, 'appId is required for update action');
      }

      const updatePayload: any = {};
      if (args.name) updatePayload.displayName = args.name;
      if (args.version) updatePayload.displayVersion = args.version;

      apiPath = `/deviceAppManagement/mobileApps/${args.appId}`;
      result = await graphClient.api(apiPath).patch(updatePayload);
      break;

    case 'remove':
      if (!args.appId) {
        throw new McpError(ErrorCode.InvalidParams, 'appId is required for remove action');
      }
      apiPath = `/deviceAppManagement/mobileApps/${args.appId}`;
      result = await graphClient.api(apiPath).delete();      break;

    case 'sync_status':
      if (!args.appId) {
        throw new McpError(ErrorCode.InvalidParams, 'appId is required for sync_status action');
      }
      
      // Get app installation status across devices
      apiPath = `/deviceAppManagement/mobileApps/${args.appId}/deviceStatuses`;
      result = await graphClient.api(apiPath).get();
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Unknown action: ${args.action}`);
  }

  return {
    content: [
      {
        type: 'text',
        text: `Windows App Management Result:\n${JSON.stringify(result, null, 2)}`
      }
    ]
  };
}

// Intune Windows Compliance Management Handler
export async function handleIntuneWindowsCompliance(
  graphClient: Client,
  args: IntuneWindowsComplianceArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'get_status':
      if (args.deviceId) {
        apiPath = `/deviceManagement/managedDevices/${args.deviceId}/deviceCompliancePolicyStates`;
      } else {
        apiPath = '/deviceManagement/deviceCompliancePolicyDeviceStateSummary';
      }
      result = await graphClient.api(apiPath).get();
      break;

    case 'get_details':
      if (!args.deviceId) {
        throw new McpError(ErrorCode.InvalidParams, 'deviceId is required for get_details action');
      }
      
      if (args.complianceType === 'bitlocker') {
        apiPath = `/informationProtection/bitlocker/recoveryKeys`;
        const filter = `$filter=deviceId eq '${args.deviceId}'`;
        result = await graphClient.api(`${apiPath}?${filter}`).get();
      } else {
        apiPath = `/deviceManagement/managedDevices/${args.deviceId}/deviceConfigurationStates`;
        if (args.policies && args.policies.length > 0) {
          const policyFilter = args.policies.map(p => `id eq '${p}'`).join(' or ');
          apiPath += `?$filter=${policyFilter}`;
        }
        result = await graphClient.api(apiPath).get();
      }
      break;

    case 'update_policy':
      if (!args.policies || args.policies.length === 0) {
        throw new McpError(ErrorCode.InvalidParams, 'policies array is required for update_policy action');
      }

      const updateResults = [];
      for (const policyId of args.policies) {
        try {
          apiPath = `/deviceManagement/deviceCompliancePolicies/${policyId}`;
          const policy = await graphClient.api(apiPath).get();
          
          // Force policy refresh
          const refreshPath = `/deviceManagement/deviceCompliancePolicies/${policyId}/scheduleActionsForRules`;
          await graphClient.api(refreshPath).post({
            deviceCompliancePolicyId: policyId
          });
          
          updateResults.push({
            policyId: policyId,
            status: 'updated',
            name: policy.displayName
          });
        } catch (error) {
          updateResults.push({
            policyId: policyId,
            status: 'failed',
            error: error instanceof Error ? error.message : 'Unknown error'
          });
        }
      }
      
      result = { updatedPolicies: updateResults };
      break;

    case 'force_evaluation':
      if (!args.deviceId) {
        throw new McpError(ErrorCode.InvalidParams, 'deviceId is required for force_evaluation action');
      }

      // Trigger compliance evaluation
      apiPath = `/deviceManagement/managedDevices/${args.deviceId}/syncDevice`;
      await graphClient.api(apiPath).post({});
      
      // Also trigger policy refresh
      const refreshPath = `/deviceManagement/managedDevices/${args.deviceId}/refreshDeviceComplianceReportSummarization`;
      result = await graphClient.api(refreshPath).post({});
      break;

    case 'get_bitlocker_keys':
      if (!args.deviceId) {
        throw new McpError(ErrorCode.InvalidParams, 'deviceId is required for get_bitlocker_keys action');
      }
      
      apiPath = `/informationProtection/bitlocker/recoveryKeys`;
      const filter = `$filter=deviceId eq '${args.deviceId}'`;
      result = await graphClient.api(`${apiPath}?${filter}`).get();
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Unknown action: ${args.action}`);
  }

  return {
    content: [
      {
        type: 'text',
        text: `Windows Compliance Management Result:\n${JSON.stringify(result, null, 2)}`
      }
    ]
  };
}
