import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { 
  IntuneMacOSDeviceArgs, 
  IntuneMacOSPolicyArgs, 
  IntuneMacOSAppArgs, 
  IntuneMacOSComplianceArgs 
} from '../types/intune-types.js';

// Intune macOS Device Management Handler
export async function handleIntuneMacOSDevices(
  graphClient: Client,
  args: IntuneMacOSDeviceArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'list':
      // List all macOS devices managed by Intune
      apiPath = '/deviceManagement/managedDevices';
      const queryOptions: string[] = [];
      
      // Filter for macOS devices
      queryOptions.push(`$filter=operatingSystem eq 'macOS'`);
      
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
      // Create enrollment invitation
      apiPath = '/deviceManagement/deviceEnrollmentConfigurations';
      const enrollmentPayload = {
        displayName: 'macOS Device Enrollment',
        description: 'Automated macOS device enrollment',
        deviceEnrollmentConfigurationType: 'appleDeviceEnrollmentProgram',
        enableAuthenticationViaCompanyPortal: true,
        requireUserAuthentication: true,
        assignmentTarget: args.assignmentTarget
      };
      result = await graphClient.api(apiPath).post(enrollmentPayload);
      break;

    case 'retire':
      if (!args.deviceId) {
        throw new McpError(ErrorCode.InvalidParams, 'deviceId is required for retire action');
      }
      apiPath = `/deviceManagement/managedDevices/${args.deviceId}/retire`;
      result = await graphClient.api(apiPath).post({});
      break;

    case 'wipe':
      if (!args.deviceId) {
        throw new McpError(ErrorCode.InvalidParams, 'deviceId is required for wipe action');
      }
      apiPath = `/deviceManagement/managedDevices/${args.deviceId}/wipe`;
      const wipePayload = {
        keepEnrollmentData: false,
        keepUserData: false,
        macOsUnlockCode: '', // Optional unlock code for macOS
        persistEsimDataPlan: false
      };
      result = await graphClient.api(apiPath).post(wipePayload);
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
      const logCollectionPayload = {
        templateType: 'predefined' // or 'custom'
      };
      result = await graphClient.api(apiPath).post(logCollectionPayload);
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}

// Intune macOS Policy Management Handler
export async function handleIntuneMacOSPolicies(
  graphClient: Client,
  args: IntuneMacOSPolicyArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'list':
      // List policies based on type
      switch (args.policyType) {
        case 'Configuration':
          apiPath = '/deviceManagement/deviceConfigurations';
          // Filter for macOS configuration policies by their OData type
          apiPath += `?$filter=odata.type eq '#microsoft.graph.macOSCustomConfiguration'`;
          break;
        case 'Compliance':
          apiPath = '/deviceManagement/deviceCompliancePolicies';
          // Filter for macOS compliance policies by their OData type
          apiPath += `?$filter=odata.type eq '#microsoft.graph.macOSCompliancePolicy'`;
          break;
        case 'Security':
          apiPath = '/deviceManagement/intents';
          break;
        case 'Update':
          apiPath = '/deviceManagement/softwareUpdateStatusSummary';
          break;
        case 'AppProtection':
          apiPath = '/deviceAppManagement/managedAppPolicies';
          break;
        default:
          throw new McpError(ErrorCode.InvalidParams, `Invalid policyType: ${args.policyType}`);
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
          apiPath = `/deviceManagement/intents/${args.policyId}`;
          break;
        default:
          throw new McpError(ErrorCode.InvalidParams, `Get operation not supported for policyType: ${args.policyType}`);
      }
      
      result = await graphClient.api(apiPath).get();
      break;

    case 'create':
      if (!args.name) {
        throw new McpError(ErrorCode.InvalidParams, 'name is required for create action');
      }

      switch (args.policyType) {
        case 'Configuration':
          apiPath = '/deviceManagement/deviceConfigurations';
          const configPayload = {
            '@odata.type': '#microsoft.graph.macOSCustomConfiguration',
            displayName: args.name,
            description: args.description || '',
            payloadFileName: args.settings?.customConfiguration?.payloadFileName || 'config.mobileconfig',
            payload: args.settings?.customConfiguration?.payload || '',
            platformType: 'macOS',
            version: 1
          };
          result = await graphClient.api(apiPath).post(configPayload);
          break;

        case 'Compliance':
          apiPath = '/deviceManagement/deviceCompliancePolicies';
          const compliancePayload = {
            '@odata.type': '#microsoft.graph.macOSCompliancePolicy',
            displayName: args.name,
            description: args.description || '',
            platformType: 'macOS',
            passwordRequired: args.settings?.compliance?.passwordRequired ?? false,
            passwordMinimumLength: args.settings?.compliance?.passwordMinimumLength ?? 4,
            passwordMinutesOfInactivityBeforeLock: args.settings?.compliance?.passwordMinutesOfInactivityBeforeLock ?? 15,
            storageRequireEncryption: args.settings?.compliance?.storageRequireEncryption ?? true,
            osMinimumVersion: args.settings?.compliance?.osMinimumVersion,
            osMaximumVersion: args.settings?.compliance?.osMaximumVersion,
            systemIntegrityProtectionEnabled: args.settings?.compliance?.systemIntegrityProtectionEnabled ?? true,
            firewallEnabled: args.settings?.compliance?.firewallEnabled ?? true,
            gatekeeperAllowedAppSource: args.settings?.compliance?.gatekeeperAllowedAppSource ?? 'macAppStoreAndIdentifiedDevelopers'
          };
          result = await graphClient.api(apiPath).post(compliancePayload);
          break;

        default:
          throw new McpError(ErrorCode.InvalidParams, `Create operation not supported for policyType: ${args.policyType}`);
      }
      break;

    case 'update':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for update action');
      }

      switch (args.policyType) {
        case 'Configuration':
          apiPath = `/deviceManagement/deviceConfigurations/${args.policyId}`;
          const updateConfigPayload = {
            displayName: args.name,
            description: args.description,
            payloadFileName: args.settings?.customConfiguration?.payloadFileName,
            payload: args.settings?.customConfiguration?.payload
          };
          result = await graphClient.api(apiPath).patch(updateConfigPayload);
          break;

        case 'Compliance':
          apiPath = `/deviceManagement/deviceCompliancePolicies/${args.policyId}`;
          const updateCompliancePayload = {
            displayName: args.name,
            description: args.description,
            passwordRequired: args.settings?.compliance?.passwordRequired,
            passwordMinimumLength: args.settings?.compliance?.passwordMinimumLength,
            storageRequireEncryption: args.settings?.compliance?.storageRequireEncryption,
            osMinimumVersion: args.settings?.compliance?.osMinimumVersion,
            osMaximumVersion: args.settings?.compliance?.osMaximumVersion
          };
          result = await graphClient.api(apiPath).patch(updateCompliancePayload);
          break;

        default:
          throw new McpError(ErrorCode.InvalidParams, `Update operation not supported for policyType: ${args.policyType}`);
      }
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
        default:
          throw new McpError(ErrorCode.InvalidParams, `Delete operation not supported for policyType: ${args.policyType}`);
      }

      await graphClient.api(apiPath).delete();
      result = { message: `${args.policyType} policy deleted successfully` };
      break;

    case 'assign':
      if (!args.policyId || !args.assignments) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId and assignments are required for assign action');
      }

      switch (args.policyType) {
        case 'Configuration':
          apiPath = `/deviceManagement/deviceConfigurations/${args.policyId}/assignments`;
          break;
        case 'Compliance':
          apiPath = `/deviceManagement/deviceCompliancePolicies/${args.policyId}/assignments`;
          break;
        default:
          throw new McpError(ErrorCode.InvalidParams, `Assign operation not supported for policyType: ${args.policyType}`);
      }

      const assignmentPayload = {
        assignments: args.assignments.map(assignment => ({
          target: {
            '@odata.type': assignment.target.groupId ? 
              '#microsoft.graph.groupAssignmentTarget' : 
              '#microsoft.graph.allDevicesAssignmentTarget',
            groupId: assignment.target.groupId
          },
          intent: assignment.intent || 'apply'
        }))
      };

      result = await graphClient.api(apiPath).post(assignmentPayload);
      break;

    case 'deploy':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for deploy action');
      }
      
      // For deploy action, we'll assign to all devices or specified groups
      const deployApiPath = args.policyType === 'Configuration' ? 
        `/deviceManagement/deviceConfigurations/${args.policyId}/assignments` :
        `/deviceManagement/deviceCompliancePolicies/${args.policyId}/assignments`;

      const deployPayload = {
        assignments: [{
          target: {
            '@odata.type': '#microsoft.graph.allDevicesAssignmentTarget'
          },
          intent: 'apply'
        }]
      };

      result = await graphClient.api(deployApiPath).post(deployPayload);
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}

// Intune macOS App Management Handler
export async function handleIntuneMacOSApps(
  graphClient: Client,
  args: IntuneMacOSAppArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'list':
      // List macOS applications
      apiPath = '/deviceAppManagement/mobileApps';
      apiPath += `?$filter=deviceType eq 'macOS'`;
      
      if (args.appType) {
        apiPath += ` and microsoft.graph.${args.appType}`;
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
      if (!args.appId || !args.assignment) {
        throw new McpError(ErrorCode.InvalidParams, 'appId and assignment are required for deploy action');
      }

      apiPath = `/deviceAppManagement/mobileApps/${args.appId}/assignments`;
      const deploymentPayload = {
        mobileAppAssignments: args.assignment.groupIds.map(groupId => ({
          target: {
            '@odata.type': '#microsoft.graph.groupAssignmentTarget',
            groupId: groupId
          },
          intent: args.assignment?.installIntent || 'available',
          settings: {
            '@odata.type': '#microsoft.graph.macOSLobAppAssignmentSettings',
            uninstallOnDeviceRemoval: true
          }
        }))
      };

      result = await graphClient.api(apiPath).post(deploymentPayload);
      break;

    case 'update':
      if (!args.appId) {
        throw new McpError(ErrorCode.InvalidParams, 'appId is required for update action');
      }

      apiPath = `/deviceAppManagement/mobileApps/${args.appId}`;
      const updatePayload = {
        displayName: args.appInfo?.displayName,
        description: args.appInfo?.description,
        publisher: args.appInfo?.publisher,
        bundleId: args.appInfo?.bundleId,
        buildNumber: args.appInfo?.buildNumber,
        versionNumber: args.appInfo?.versionNumber,
        minimumSupportedOperatingSystem: args.appInfo?.minimumSupportedOperatingSystem
      };

      result = await graphClient.api(apiPath).patch(updatePayload);
      break;

    case 'remove':
      if (!args.appId) {
        throw new McpError(ErrorCode.InvalidParams, 'appId is required for remove action');
      }

      apiPath = `/deviceAppManagement/mobileApps/${args.appId}`;
      await graphClient.api(apiPath).delete();
      result = { message: 'App removed successfully' };
      break;

    case 'sync_status':
      if (!args.appId) {
        throw new McpError(ErrorCode.InvalidParams, 'appId is required for sync_status action');
      }

      // Get app installation status across devices
      apiPath = `/deviceAppManagement/mobileApps/${args.appId}/deviceStatuses`;
      result = await graphClient.api(apiPath).get();
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}

// Intune macOS Compliance Monitoring Handler
export async function handleIntuneMacOSCompliance(
  graphClient: Client,
  args: IntuneMacOSComplianceArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let apiPath = '';
  let result: any;

  switch (args.action) {
    case 'get_status':
      if (args.deviceId) {
        // Get compliance status for specific device
        apiPath = `/deviceManagement/managedDevices/${args.deviceId}/deviceCompliancePolicyStates`;
      } else {
        // Get overall compliance status for macOS devices
        apiPath = '/deviceManagement/deviceCompliancePolicyDeviceStateSummary';
        apiPath += `?$filter=platformType eq 'macOS'`;
      }
      result = await graphClient.api(apiPath).get();
      break;

    case 'get_details':
      if (!args.deviceId) {
        throw new McpError(ErrorCode.InvalidParams, 'deviceId is required for get_details action');
      }

      // Get detailed compliance information for device
      apiPath = `/deviceManagement/managedDevices/${args.deviceId}/deviceCompliancePolicyStates`;
      const complianceStates = await graphClient.api(apiPath).get();

      // Get device configuration states
      const configApiPath = `/deviceManagement/managedDevices/${args.deviceId}/deviceConfigurationStates`;
      const configStates = await graphClient.api(configApiPath).get();

      result = {
        deviceId: args.deviceId,
        compliancePolicyStates: complianceStates,
        configurationStates: configStates
      };
      break;

    case 'update_policy':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for update_policy action');
      }

      apiPath = `/deviceManagement/deviceCompliancePolicies/${args.policyId}`;
      const updatePayload = {
        passwordRequired: args.complianceData?.passwordCompliant,
        storageRequireEncryption: args.complianceData?.encryptionCompliant,
        systemIntegrityProtectionEnabled: args.complianceData?.systemIntegrityCompliant,
        firewallEnabled: args.complianceData?.firewallCompliant
      };

      result = await graphClient.api(apiPath).patch(updatePayload);
      break;

    case 'force_evaluation':
      if (!args.deviceId) {
        throw new McpError(ErrorCode.InvalidParams, 'deviceId is required for force_evaluation action');
      }

      // Trigger compliance evaluation on device
      apiPath = `/deviceManagement/managedDevices/${args.deviceId}/syncDevice`;
      await graphClient.api(apiPath).post({});

      // Also trigger compliance policy evaluation
      const evalApiPath = `/deviceManagement/managedDevices/${args.deviceId}/triggerConfigurationManagerAction`;
      await graphClient.api(evalApiPath).post({
        action: {
          actionType: 'evaluateCompliance'
        }
      });

      result = { message: 'Compliance evaluation triggered successfully' };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}
