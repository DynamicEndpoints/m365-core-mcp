import { z } from 'zod';

// Enhanced settings schemas for each policy type and platform

// macOS specific schemas
const macOSConfigurationSettingsSchema = z.object({
  customConfiguration: z.object({
    payloadFileName: z.string().describe('Filename for the configuration profile (e.g., "config.mobileconfig")'),
    payload: z.string().describe('Base64 encoded .mobileconfig file content')
  }).describe('Custom configuration profile settings')
});

const macOSComplianceSettingsSchema = z.object({
  compliance: z.object({
    passwordRequired: z.boolean().optional().describe('Require password'),
    passwordMinimumLength: z.number().min(4).max(14).optional().describe('Minimum password length (4-14)'),
    passwordMinutesOfInactivityBeforeLock: z.number().min(1).max(60).optional().describe('Minutes before lock (1-60)'),
    passwordMinutesOfInactivityBeforeScreenTimeout: z.number().min(1).max(60).optional().describe('Minutes before screen timeout'),
    passwordPreviousPasswordBlockCount: z.number().min(0).max(24).optional().describe('Password history count'),
    passwordRequiredType: z.enum(['deviceDefault', 'alphanumeric', 'numeric']).optional().describe('Password type requirement'),
    deviceThreatProtectionEnabled: z.boolean().optional().describe('Enable Mobile Threat Defense'),
    deviceThreatProtectionRequiredSecurityLevel: z.enum(['unavailable', 'secured', 'low', 'medium', 'high', 'notSet']).optional(),
    storageRequireEncryption: z.boolean().optional().describe('Require FileVault encryption'),
    osMinimumVersion: z.string().optional().describe('Minimum macOS version (e.g., "14.0")'),
    osMaximumVersion: z.string().optional().describe('Maximum macOS version'),
    systemIntegrityProtectionEnabled: z.boolean().optional().describe('Require System Integrity Protection'),
    firewallEnabled: z.boolean().optional().describe('Require firewall enabled'),
    firewallBlockAllIncoming: z.boolean().optional().describe('Block all incoming connections'),
    firewallEnableStealthMode: z.boolean().optional().describe('Enable firewall stealth mode'),
    gatekeeperAllowedAppSource: z.enum(['notConfigured', 'macAppStore', 'macAppStoreAndIdentifiedDevelopers', 'anywhere']).optional()
  }).optional().describe('macOS compliance settings')
});

// Windows specific schemas
const windowsConfigurationSettingsSchema = z.object({
  passwordRequired: z.boolean().optional().describe('Require password'),
  passwordBlockSimple: z.boolean().optional().describe('Block simple passwords'),
  passwordMinimumLength: z.number().min(0).max(127).optional().describe('Minimum password length'),
  defenderMonitorFileActivity: z.boolean().optional().describe('Monitor file activity'),
  defenderScanNetworkFiles: z.boolean().optional().describe('Scan network files'),
  defenderEnableScanIncomingMail: z.boolean().optional().describe('Scan incoming mail'),
  defenderEnableScanMappedNetworkDrivesDuringFullScan: z.boolean().optional().describe('Scan mapped drives'),
  smartScreenEnabled: z.boolean().optional().describe('Enable SmartScreen'),
  firewallProfileDomain: z.object({
    firewallEnabled: z.enum(['notConfigured', 'blocked', 'allowed']).optional(),
    stealthModeRequired: z.boolean().optional(),
    incomingTrafficBlocked: z.boolean().optional(),
    outboundConnectionsRequired: z.boolean().optional()
  }).optional().describe('Domain firewall profile'),
  bluetoothBlocked: z.boolean().optional().describe('Block Bluetooth'),
  cameraBlocked: z.boolean().optional().describe('Block camera'),
  cortanaBlocked: z.boolean().optional().describe('Block Cortana'),
  windowsSpotlightBlocked: z.boolean().optional().describe('Block Windows Spotlight')
}).describe('Windows configuration settings');

const windowsComplianceSettingsSchema = z.object({
  passwordRequired: z.boolean().optional().describe('Require password'),
  passwordBlockSimple: z.boolean().optional().describe('Block simple passwords'),
  passwordMinimumLength: z.number().min(0).max(127).optional().describe('Minimum password length'),
  passwordMinutesOfInactivityBeforeLock: z.number().min(0).max(999).optional().describe('Minutes before lock'),
  passwordExpirationDays: z.number().min(0).max(730).optional().describe('Password expiration days'),
  passwordPreviousPasswordBlockCount: z.number().min(0).max(50).optional().describe('Password history'),
  passwordMinimumCharacterSetCount: z.number().min(0).max(4).optional().describe('Character set count'),
  passwordRequiredType: z.enum(['deviceDefault', 'alphanumeric', 'numeric']).optional(),
  requireHealthyDeviceReport: z.boolean().optional().describe('Require healthy device report'),
  osMinimumVersion: z.string().optional().describe('Minimum OS version (e.g., "10.0.19041")'),
  osMaximumVersion: z.string().optional().describe('Maximum OS version'),
  earlyLaunchAntiMalwareDriverEnabled: z.boolean().optional().describe('Require ELAM driver'),
  bitLockerEnabled: z.boolean().optional().describe('Require BitLocker'),
  secureBootEnabled: z.boolean().optional().describe('Require Secure Boot'),
  codeIntegrityEnabled: z.boolean().optional().describe('Require Code Integrity'),
  storageRequireEncryption: z.boolean().optional().describe('Require storage encryption'),
  activeFirewallRequired: z.boolean().optional().describe('Require active firewall'),
  defenderEnabled: z.boolean().optional().describe('Require Windows Defender'),
  antivirusRequired: z.boolean().optional().describe('Require antivirus'),
  antiSpywareRequired: z.boolean().optional().describe('Require anti-spyware'),
  rtpEnabled: z.boolean().optional().describe('Require real-time protection'),
  signatureOutOfDate: z.boolean().optional().describe('Allow outdated signatures')
}).describe('Windows compliance settings');

const windowsUpdateSettingsSchema = z.object({
  automaticUpdateMode: z.enum([
    'userDefined', 'notifyDownload', 'autoInstallAtMaintenanceTime', 
    'autoInstallAndRebootAtMaintenanceTime', 'autoInstallAndRebootAtScheduledTime', 
    'autoInstallAndRebootWithoutEndUserControl', 'windowsDefault'
  ]).optional().describe('Automatic update mode'),
  microsoftUpdateServiceAllowed: z.boolean().optional().describe('Allow Microsoft Update service'),
  driversExcluded: z.boolean().optional().describe('Exclude drivers from updates'),
  qualityUpdatesDeferralPeriodInDays: z.number().min(0).max(30).optional().describe('Quality updates deferral (0-30 days)'),
  featureUpdatesDeferralPeriodInDays: z.number().min(0).max(365).optional().describe('Feature updates deferral (0-365 days)'),
  businessReadyUpdatesOnly: z.enum(['userDefined', 'all', 'businessReadyOnly']).optional(),
  prereleaseFeatures: z.enum(['userDefined', 'settingsOnly', 'settingsAndExperimentations', 'notAllowed']).optional(),
  deliveryOptimizationMode: z.enum([
    'userDefined', 'httpOnly', 'httpWithPeeringNat', 'httpWithPeeringPrivateGroup', 
    'httpWithInternetPeering', 'simpleDownload', 'bypassMode'
  ]).optional(),
  installationSchedule: z.object({
    scheduledInstallDay: z.enum([
      'userDefined', 'everyday', 'sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'
    ]).optional(),
    scheduledInstallTime: z.string().optional().describe('Time in HH:MM format'),
    activeHoursStart: z.string().optional().describe('Active hours start in HH:MM format'),
    activeHoursEnd: z.string().optional().describe('Active hours end in HH:MM format')
  }).optional()
}).describe('Windows Update settings');

// Assignment schema
const assignmentSchema = z.object({
  target: z.object({
    deviceAndAppManagementAssignmentFilterId: z.string().optional().describe('Assignment filter ID'),
    deviceAndAppManagementAssignmentFilterType: z.enum(['none', 'include', 'exclude']).optional(),
    groupId: z.string().optional().describe('Azure AD group ID'),
    collectionId: z.string().optional().describe('Collection ID')
  }).describe('Assignment target'),
  intent: z.enum(['apply', 'remove']).optional().describe('Assignment intent'),
  settings: z.object({
    installIntent: z.enum(['available', 'required', 'uninstall', 'availableWithoutEnrollment']).optional(),
    notificationSettings: z.object({
      showInCompanyPortal: z.boolean().optional(),
      showInNotificationCenter: z.boolean().optional(),
      alertType: z.enum(['showAll', 'showRebootsOnly', 'hideAll']).optional()
    }).optional(),
    restartSettings: z.object({
      restartNotificationSnoozeDurationInMinutes: z.number().optional(),
      restartCountdownDisplayDurationInMinutes: z.number().optional()
    }).optional(),
    deadlineSettings: z.object({
      useLocalTime: z.boolean().optional(),
      deadlineDateTime: z.string().optional(),
      gracePeriodInMinutes: z.number().optional()
    }).optional()
  }).optional()
});

// Main enhanced schema with discriminated union for platform-specific settings
const enhancedCreateIntunePolicySchema = z.object({
  platform: z.enum(['windows', 'macos']).describe('The platform for the policy'),
  policyType: z.enum(['Configuration', 'Compliance', 'Security', 'Update', 'AppProtection', 'EndpointSecurity']).describe('The type of policy to create'),
  displayName: z.string().min(1).max(256).describe('The name of the policy (1-256 characters)'),
  description: z.string().max(1024).optional().describe('The description of the policy (max 1024 characters)'),
  useTemplate: z.string().optional().describe('Use a predefined template (basicSecurity, strictSecurity, windowsUpdate)'),
  assignments: z.array(assignmentSchema).optional().describe('Policy assignments to groups or filters')
}).and(
  z.discriminatedUnion('platform', [
    // macOS platform
    z.object({
      platform: z.literal('macos'),
      settings: z.discriminatedUnion('policyType', [
        z.object({
          policyType: z.literal('Configuration'),
          settings: macOSConfigurationSettingsSchema.optional()
        }),
        z.object({
          policyType: z.literal('Compliance'),
          settings: macOSComplianceSettingsSchema.optional()
        }),
        z.object({
          policyType: z.enum(['Security', 'Update', 'AppProtection']),
          settings: z.object({}).optional().describe('Settings not fully implemented for this policy type')
        })
      ]).optional()
    }),
    // Windows platform
    z.object({
      platform: z.literal('windows'),
      settings: z.discriminatedUnion('policyType', [
        z.object({
          policyType: z.literal('Configuration'),
          settings: windowsConfigurationSettingsSchema.optional()
        }),
        z.object({
          policyType: z.literal('Compliance'),
          settings: windowsComplianceSettingsSchema.optional()
        }),
        z.object({
          policyType: z.literal('Update'),
          settings: windowsUpdateSettingsSchema.optional()
        }),
        z.object({
          policyType: z.enum(['Security', 'AppProtection', 'EndpointSecurity']),
          settings: z.object({
            templateId: z.string().optional().describe('Template ID for security policies')
          }).optional()
        })
      ]).optional()
    })
  ])
);

export const enhancedIntuneTools = [
  {
    name: 'create_intune_policy',
    description: 'Create accurate and complete Intune policies for Windows or macOS with validated settings and proper structure',
    inputSchema: enhancedCreateIntunePolicySchema
  }
];

// Also export a simplified version that maintains backwards compatibility
export const createIntunePolicySchema = z.object({
  platform: z.enum(['windows', 'macos']).describe('The platform for the policy.'),
  policyType: z.enum(['Configuration', 'Compliance', 'Security', 'Update', 'AppProtection', 'EndpointSecurity']).describe('The type of policy to create.'),
  displayName: z.string().describe('The name of the policy.'),
  description: z.string().optional().describe('The description of the policy.'),
  settings: z.any().describe('The policy settings. The structure of this object depends on the platform and policyType.'),
  assignments: z.array(z.any()).optional().describe('The assignments for the policy.'),
  useTemplate: z.string().optional().describe('Use a predefined template for common configurations.')
});

export const intuneTools = [
  {
    name: 'createIntunePolicy',
    description: 'Creates a new Intune policy for either Windows or macOS.',
    inputSchema: createIntunePolicySchema
  },
  {
    name: 'create_intune_policy',
    description: 'Create accurate and complete Intune policies for Windows or macOS with validated settings and proper structure',
    inputSchema: enhancedCreateIntunePolicySchema
  }
];
