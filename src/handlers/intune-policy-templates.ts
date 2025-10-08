/**
 * Intune Policy Templates and Helpers
 * Provides templates and utilities for creating Settings Catalog and PPC policies
 */

// Settings Catalog Template IDs
export const SETTINGS_CATALOG_TEMPLATES = {
  // Windows Security Baselines
  SECURITY_BASELINE_WINDOWS: '034ccd46-190c-4afc-adf1-ad7cc11262eb',
  SECURITY_BASELINE_EDGE: '87e5e4ad-6f6c-4cdf-9c7a-5c6f8e1e8e8e',
  SECURITY_BASELINE_DEFENDER: 'e044e60e-5901-41ea-92c5-87e8b6edd6bb',
  
  // Device Configuration
  DEVICE_RESTRICTIONS: 'd1174162-1dd2-4976-affc-6667049ab0ae',
  ENDPOINT_PROTECTION: '0e237410-1367-4844-bd7f-15fb0f08943b',
  
  // Application Management
  APP_CONFIGURATION: '95d6e8e0-0f9e-4e5a-9c7e-3c3f0f2f1f1f',
};

// Platform Protection Configuration Templates
export const PPC_TEMPLATES = {
  ATTACK_SURFACE_REDUCTION: '9dc5088e-2e9e-4d98-bc1f-89c6a6f0e6e6',
  EXPLOIT_PROTECTION: '15e3f5e0-9c3e-4a8e-9e0a-7c8e0e9e0e9e',
  WEB_PROTECTION: '2e9e8e0e-9e0e-4e0e-9e0e-8e0e9e0e9e0e',
};

/**
 * Settings Catalog Policy Structure
 */
export interface SettingsCatalogPolicy {
  name: string;
  description?: string;
  platforms: 'windows10' | 'macOS' | 'iOS' | 'android';
  technologies: 'mdm' | 'windows10Endpointprotection' | 'configManager';
  settings: SettingsCatalogSetting[];
  templateReference?: {
    templateId: string;
    templateFamily?: string;
  };
}

export interface SettingsCatalogSetting {
  '@odata.type': string;
  settingInstance: {
    '@odata.type': string;
    settingDefinitionId: string;
    [key: string]: any;
  };
}

/**
 * Common Settings Catalog Setting Definitions
 */
export const COMMON_SETTINGS_DEFINITIONS = {
  // BitLocker Settings
  BITLOCKER_REQUIRE_DEVICE_ENCRYPTION: 'device_vendor_msft_bitlocker_requiredeviceencryption',
  BITLOCKER_FIXED_DRIVE_ENCRYPTION_TYPE: 'device_vendor_msft_bitlocker_fixeddrivesencryptiontype',
  BITLOCKER_REMOVABLE_DRIVE_ENCRYPTION_TYPE: 'device_vendor_msft_bitlocker_removabledrivesencryptiontype',
  
  // Windows Defender Settings
  DEFENDER_REAL_TIME_PROTECTION: 'device_vendor_msft_policy_config_defender_allowrealtimemonitoring',
  DEFENDER_CLOUD_PROTECTION: 'device_vendor_msft_policy_config_defender_allowcloudprotection',
  DEFENDER_BEHAVIOR_MONITORING: 'device_vendor_msft_policy_config_defender_allowbehaviormonitoring',
  DEFENDER_SCAN_TYPE: 'device_vendor_msft_policy_config_defender_scanparameter',
  
  // Windows Update Settings
  UPDATE_BRANCH_READINESS_LEVEL: 'device_vendor_msft_policy_config_update_branchreadinesslevel',
  UPDATE_DEFER_QUALITY_UPDATES: 'device_vendor_msft_policy_config_update_deferqualityupdatesperiodindays',
  UPDATE_DEFER_FEATURE_UPDATES: 'device_vendor_msft_policy_config_update_deferfeatureupdatesperiodindays',
  
  // Firewall Settings  
  FIREWALL_DOMAIN_PROFILE_ENABLED: 'vendor_msft_firewall_mdmstore_domainprofile_enablefirewall',
  FIREWALL_PUBLIC_PROFILE_ENABLED: 'vendor_msft_firewall_mdmstore_publicprofile_enablefirewall',
  FIREWALL_PRIVATE_PROFILE_ENABLED: 'vendor_msft_firewall_mdmstore_privateprofile_enablefirewall',
  
  // Password Policy Settings
  PASSWORD_MIN_LENGTH: 'device_vendor_msft_policy_config_devicelock_mindevicepasswordlength',
  PASSWORD_COMPLEXITY: 'device_vendor_msft_policy_config_devicelock_mindevicepasswordcomplexcharacters',
  PASSWORD_EXPIRATION: 'device_vendor_msft_policy_config_devicelock_devicepasswordexpiration',
  PASSWORD_HISTORY: 'device_vendor_msft_policy_config_devicelock_devicepasswordhistory',
  
  // Application Control
  APPLOCKER_EXE_RULES: 'device_vendor_msft_applocker_applicationlaunchrestrictions_groupname_exe',
  APPLOCKER_DLL_RULES: 'device_vendor_msft_applocker_applicationlaunchrestrictions_groupname_dll',
  
  // Attack Surface Reduction
  ASR_RULES: 'device_vendor_msft_policy_config_defender_attacksurfacereductionrules',
  ASR_ONLY_EXCLUSIONS: 'device_vendor_msft_policy_config_defender_attacksurfacereductiononlyexclusions',
};

/**
 * Create a Settings Catalog setting
 */
export function createSettingsCatalogSetting(
  settingDefinitionId: string,
  value: any,
  valueType: 'string' | 'int' | 'boolean' | 'collection' = 'string'
): SettingsCatalogSetting {
  const baseType = '#microsoft.graph.deviceManagementConfiguration';
  
  let settingValueType: string;
  let settingValue: any;
  
  switch (valueType) {
    case 'int':
      settingValueType = `${baseType}IntegerSettingValue`;
      settingValue = { value: parseInt(value) };
      break;
    case 'boolean':
      settingValueType = `${baseType}ChoiceSettingValue`;
      settingValue = { 
        value: `${settingDefinitionId}_${value ? '1' : '0'}`,
        children: []
      };
      break;
    case 'collection':
      settingValueType = `${baseType}GroupSettingCollectionInstance`;
      settingValue = { groupSettingCollectionValue: value };
      break;
    default:
      settingValueType = `${baseType}StringSettingValue`;
      settingValue = { value: String(value) };
  }
  
  return {
    '@odata.type': `${baseType}SettingInstance`,
    settingInstance: {
      '@odata.type': `${baseType}SimpleSettingInstance`,
      settingDefinitionId: settingDefinitionId,
      simpleSettingValue: {
        '@odata.type': settingValueType,
        ...settingValue
      }
    }
  };
}

/**
 * Pre-built Settings Catalog Policy Templates
 */
export const SETTINGS_CATALOG_POLICY_TEMPLATES = {
  /**
   * BitLocker Encryption Policy
   */
  BITLOCKER_ENCRYPTION: (): SettingsCatalogPolicy => ({
    name: 'BitLocker Disk Encryption',
    description: 'Enforce BitLocker encryption on Windows devices',
    platforms: 'windows10',
    technologies: 'mdm',
    settings: [
      createSettingsCatalogSetting(
        COMMON_SETTINGS_DEFINITIONS.BITLOCKER_REQUIRE_DEVICE_ENCRYPTION,
        true,
        'boolean'
      ),
      createSettingsCatalogSetting(
        COMMON_SETTINGS_DEFINITIONS.BITLOCKER_FIXED_DRIVE_ENCRYPTION_TYPE,
        1, // Full encryption
        'int'
      ),
      createSettingsCatalogSetting(
        COMMON_SETTINGS_DEFINITIONS.BITLOCKER_REMOVABLE_DRIVE_ENCRYPTION_TYPE,
        1,
        'int'
      )
    ]
  }),
  
  /**
   * Windows Defender Antivirus Policy
   */
  DEFENDER_ANTIVIRUS: (): SettingsCatalogPolicy => ({
    name: 'Windows Defender Antivirus Protection',
    description: 'Configure Windows Defender antivirus protection settings',
    platforms: 'windows10',
    technologies: 'mdm',
    settings: [
      createSettingsCatalogSetting(
        COMMON_SETTINGS_DEFINITIONS.DEFENDER_REAL_TIME_PROTECTION,
        true,
        'boolean'
      ),
      createSettingsCatalogSetting(
        COMMON_SETTINGS_DEFINITIONS.DEFENDER_CLOUD_PROTECTION,
        true,
        'boolean'
      ),
      createSettingsCatalogSetting(
        COMMON_SETTINGS_DEFINITIONS.DEFENDER_BEHAVIOR_MONITORING,
        true,
        'boolean'
      ),
      createSettingsCatalogSetting(
        COMMON_SETTINGS_DEFINITIONS.DEFENDER_SCAN_TYPE,
        2, // Full scan
        'int'
      )
    ]
  }),
  
  /**
   * Windows Update Policy
   */
  WINDOWS_UPDATE: (deferQualityDays: number = 7, deferFeatureDays: number = 14): SettingsCatalogPolicy => ({
    name: 'Windows Update Configuration',
    description: 'Configure Windows Update settings and deferral periods',
    platforms: 'windows10',
    technologies: 'mdm',
    settings: [
      createSettingsCatalogSetting(
        COMMON_SETTINGS_DEFINITIONS.UPDATE_BRANCH_READINESS_LEVEL,
        16, // Current Branch for Business
        'int'
      ),
      createSettingsCatalogSetting(
        COMMON_SETTINGS_DEFINITIONS.UPDATE_DEFER_QUALITY_UPDATES,
        deferQualityDays,
        'int'
      ),
      createSettingsCatalogSetting(
        COMMON_SETTINGS_DEFINITIONS.UPDATE_DEFER_FEATURE_UPDATES,
        deferFeatureDays,
        'int'
      )
    ]
  }),
  
  /**
   * Firewall Policy
   */
  FIREWALL_CONFIGURATION: (): SettingsCatalogPolicy => ({
    name: 'Windows Firewall Configuration',
    description: 'Enable and configure Windows Firewall for all network profiles',
    platforms: 'windows10',
    technologies: 'mdm',
    settings: [
      createSettingsCatalogSetting(
        COMMON_SETTINGS_DEFINITIONS.FIREWALL_DOMAIN_PROFILE_ENABLED,
        true,
        'boolean'
      ),
      createSettingsCatalogSetting(
        COMMON_SETTINGS_DEFINITIONS.FIREWALL_PUBLIC_PROFILE_ENABLED,
        true,
        'boolean'
      ),
      createSettingsCatalogSetting(
        COMMON_SETTINGS_DEFINITIONS.FIREWALL_PRIVATE_PROFILE_ENABLED,
        true,
        'boolean'
      )
    ]
  }),
  
  /**
   * Password Policy
   */
  PASSWORD_POLICY: (minLength: number = 8, complexity: number = 2): SettingsCatalogPolicy => ({
    name: 'Device Password Policy',
    description: 'Configure password requirements for Windows devices',
    platforms: 'windows10',
    technologies: 'mdm',
    settings: [
      createSettingsCatalogSetting(
        COMMON_SETTINGS_DEFINITIONS.PASSWORD_MIN_LENGTH,
        minLength,
        'int'
      ),
      createSettingsCatalogSetting(
        COMMON_SETTINGS_DEFINITIONS.PASSWORD_COMPLEXITY,
        complexity,
        'int'
      ),
      createSettingsCatalogSetting(
        COMMON_SETTINGS_DEFINITIONS.PASSWORD_EXPIRATION,
        90, // 90 days
        'int'
      ),
      createSettingsCatalogSetting(
        COMMON_SETTINGS_DEFINITIONS.PASSWORD_HISTORY,
        24, // Remember 24 previous passwords
        'int'
      )
    ]
  }),
  
  /**
   * Attack Surface Reduction Rules
   */
  ATTACK_SURFACE_REDUCTION: (): SettingsCatalogPolicy => ({
    name: 'Attack Surface Reduction Rules',
    description: 'Configure Attack Surface Reduction rules to protect against threats',
    platforms: 'windows10',
    technologies: 'mdm',
    settings: [
      createSettingsCatalogSetting(
        COMMON_SETTINGS_DEFINITIONS.ASR_RULES,
        [
          // Block executable content from email and webmail
          'BE9BA2D9-53EA-4CDC-84E5-9B1EEEE46550=1',
          // Block Office applications from creating executable content
          '3B576869-A4EC-4529-8536-B80A7769E899=1',
          // Block Office applications from injecting code into other processes
          '75668C1F-73B5-4CF0-BB93-3ECF5CB7CC84=1',
          // Block JavaScript or VBScript from launching downloaded executable content
          'D3E037E1-3EB8-44C8-A917-57927947596D=1',
          // Block execution of potentially obfuscated scripts
          '5BEB7EFE-FD9A-4556-801D-275E5FFC04CC=1'
        ],
        'collection'
      )
    ]
  })
};

/**
 * Platform Protection Configuration (PPC) Helper
 */
export interface PPCPolicyConfig {
  name: string;
  description?: string;
  templateId: string;
  settings: PPCSetting[];
  assignments?: any[];
}

export interface PPCSetting {
  id: string;
  value: any;
  valueState?: 'configured' | 'notConfigured';
}

/**
 * Create a PPC policy configuration
 */
export function createPPCPolicy(
  name: string,
  templateId: string,
  settings: Record<string, any>,
  description?: string
): PPCPolicyConfig {
  const ppcSettings: PPCSetting[] = Object.entries(settings).map(([id, value]) => ({
    id,
    value,
    valueState: 'configured'
  }));
  
  return {
    name,
    description: description || '',
    templateId,
    settings: ppcSettings
  };
}

/**
 * Pre-built PPC Policy Templates
 */
export const PPC_POLICY_TEMPLATES = {
  /**
   * Attack Surface Reduction PPC Policy
   */
  ATTACK_SURFACE_REDUCTION_PPC: (): PPCPolicyConfig => createPPCPolicy(
    'Attack Surface Reduction',
    PPC_TEMPLATES.ATTACK_SURFACE_REDUCTION,
    {
      'blockExecutableContentFromEmailAndWebmail': 'block',
      'blockOfficeAppsFromCreatingExecutableContent': 'block',
      'blockOfficeAppsFromInjectingIntoOtherProcesses': 'block',
      'blockJavaScriptOrVBScriptFromLaunchingContent': 'block',
      'blockExecutionOfPotentiallyObfuscatedScripts': 'block',
      'blockWin32ApiCallsFromOfficeMacros': 'block',
      'blockUntrustedUnsignedProcesses': 'block',
      'blockCredentialStealingFromWindowsLsass': 'block',
      'blockAdobeReaderFromCreatingChildProcesses': 'block',
      'blockPersistenceThroughWMIEventSubscription': 'block'
    },
    'Configure Attack Surface Reduction rules for endpoint protection'
  ),
  
  /**
   * Exploit Protection PPC Policy
   */
  EXPLOIT_PROTECTION_PPC: (): PPCPolicyConfig => createPPCPolicy(
    'Exploit Protection',
    PPC_TEMPLATES.EXPLOIT_PROTECTION,
    {
      'dataExecutionPrevention': 'on',
      'controlFlowGuard': 'on',
      'randomizeMemoryAllocations': 'on',
      'validateExceptionChains': 'on',
      'validateStackIntegrity': 'on',
      'disableExtensionPoints': 'on',
      'disableWin32kSystemCalls': 'on',
      'blockUntrustedFonts': 'block',
      'codeIntegrityGuard': 'on',
      'blockRemoteImageLoads': 'on'
    },
    'Configure exploit protection settings for Windows devices'
  ),
  
  /**
   * Web Protection PPC Policy
   */
  WEB_PROTECTION_PPC: (): PPCPolicyConfig => createPPCPolicy(
    'Web Protection',
    PPC_TEMPLATES.WEB_PROTECTION,
    {
      'enableNetworkProtection': 'enabled',
      'networkProtectionLevel': 'block',
      'enableSmartScreenForEdge': 'enabled',
      'preventSmartScreenPromptOverride': 'required',
      'preventSmartScreenPromptOverrideForFiles': 'required',
      'allowUserFeedback': 'notAllowed',
      'allowUserToBlockMaliciousSites': 'notAllowed'
    },
    'Configure web protection settings including SmartScreen'
  )
};

/**
 * Validate Settings Catalog policy structure
 */
export function validateSettingsCatalogPolicy(policy: SettingsCatalogPolicy): { valid: boolean; errors: string[] } {
  const errors: string[] = [];
  
  if (!policy.name || policy.name.trim() === '') {
    errors.push('Policy name is required');
  }
  
  if (!policy.platforms) {
    errors.push('Platform specification is required');
  }
  
  if (!policy.technologies) {
    errors.push('Technology specification is required');
  }
  
  if (!policy.settings || policy.settings.length === 0) {
    errors.push('At least one setting is required');
  }
  
  // Validate each setting
  policy.settings?.forEach((setting, index) => {
    if (!setting.settingInstance?.settingDefinitionId) {
      errors.push(`Setting at index ${index} is missing settingDefinitionId`);
    }
  });
  
  return {
    valid: errors.length === 0,
    errors
  };
}

/**
 * Validate PPC policy structure
 */
export function validatePPCPolicy(policy: PPCPolicyConfig): { valid: boolean; errors: string[] } {
  const errors: string[] = [];
  
  if (!policy.name || policy.name.trim() === '') {
    errors.push('Policy name is required');
  }
  
  if (!policy.templateId) {
    errors.push('Template ID is required');
  }
  
  if (!policy.settings || policy.settings.length === 0) {
    errors.push('At least one setting is required');
  }
  
  // Validate each setting
  policy.settings?.forEach((setting, index) => {
    if (!setting.id) {
      errors.push(`Setting at index ${index} is missing id`);
    }
    if (setting.value === undefined || setting.value === null) {
      errors.push(`Setting at index ${index} is missing value`);
    }
  });
  
  return {
    valid: errors.length === 0,
    errors
  };
}
