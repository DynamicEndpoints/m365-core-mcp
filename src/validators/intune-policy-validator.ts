import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';

interface PolicyValidationResult {
  isValid: boolean;
  errors: string[];
  warnings: string[];
}

// Policy validation rules by platform and type
const POLICY_VALIDATION_RULES = {
  macos: {
    Configuration: {
      requiredFields: ['customConfiguration.payloadFileName', 'customConfiguration.payload'],
      validate: (settings: any): PolicyValidationResult => {
        const errors: string[] = [];
        const warnings: string[] = [];
        
        if (!settings.customConfiguration) {
          errors.push('macOS Configuration policies require a customConfiguration object');
        } else {
          if (!settings.customConfiguration.payloadFileName) {
            errors.push('customConfiguration.payloadFileName is required');
          }
          if (!settings.customConfiguration.payload) {
            errors.push('customConfiguration.payload is required (base64 encoded .mobileconfig)');
          }
          
          // Validate payload is base64
          if (settings.customConfiguration.payload && !isValidBase64(settings.customConfiguration.payload)) {
            warnings.push('customConfiguration.payload should be base64 encoded');
          }
        }
        
        return { isValid: errors.length === 0, errors, warnings };
      }
    },
    Compliance: {
      requiredFields: [],
      validate: (settings: any): PolicyValidationResult => {
        const errors: string[] = [];
        const warnings: string[] = [];
        
        if (!settings.compliance) {
          warnings.push('macOS Compliance policies should have a compliance object with settings');
        } else {
          // Validate compliance settings
          if (settings.compliance.passwordMinimumLength && settings.compliance.passwordMinimumLength < 4) {
            errors.push('passwordMinimumLength must be at least 4');
          }
          
          if (settings.compliance.passwordRequiredType && 
              !['deviceDefault', 'alphanumeric', 'numeric'].includes(settings.compliance.passwordRequiredType)) {
            errors.push('passwordRequiredType must be one of: deviceDefault, alphanumeric, numeric');
          }
          
          if (settings.compliance.gatekeeperAllowedAppSource && 
              !['notConfigured', 'macAppStore', 'macAppStoreAndIdentifiedDevelopers', 'anywhere'].includes(
                settings.compliance.gatekeeperAllowedAppSource)) {
            errors.push('gatekeeperAllowedAppSource has invalid value');
          }
        }
        
        return { isValid: errors.length === 0, errors, warnings };
      }
    },
    Security: {
      requiredFields: [],
      validate: (settings: any): PolicyValidationResult => {
        const errors: string[] = [];
        const warnings: string[] = [];
        
        warnings.push('macOS Security policies are not fully implemented. Consider using Configuration policies with security-focused .mobileconfig');
        
        return { isValid: errors.length === 0, errors, warnings };
      }
    },
    Update: {
      requiredFields: [],
      validate: (settings: any): PolicyValidationResult => {
        const errors: string[] = [];
        const warnings: string[] = [];
        
        warnings.push('macOS Update policies are not fully implemented. Consider using Configuration policies with software update .mobileconfig');
        
        return { isValid: errors.length === 0, errors, warnings };
      }
    },
    AppProtection: {
      requiredFields: [],
      validate: (settings: any): PolicyValidationResult => {
        const errors: string[] = [];
        const warnings: string[] = [];
        
        warnings.push('macOS AppProtection policies are not fully implemented');
        
        return { isValid: errors.length === 0, errors, warnings };
      }
    }
  },
  windows: {
    Configuration: {
      requiredFields: [],
      validate: (settings: any): PolicyValidationResult => {
        const errors: string[] = [];
        const warnings: string[] = [];
        
        // Windows Configuration policies use direct settings
        if (Object.keys(settings).length === 0) {
          warnings.push('No configuration settings provided. Policy will use defaults.');
        }
        
        return { isValid: errors.length === 0, errors, warnings };
      }
    },
    Compliance: {
      requiredFields: [],
      validate: (settings: any): PolicyValidationResult => {
        const errors: string[] = [];
        const warnings: string[] = [];
        
        // Validate password settings
        if (settings.passwordMinimumLength && settings.passwordMinimumLength < 0) {
          errors.push('passwordMinimumLength cannot be negative');
        }
        
        if (settings.passwordRequiredType && 
            !['deviceDefault', 'alphanumeric', 'numeric'].includes(settings.passwordRequiredType)) {
          errors.push('passwordRequiredType must be one of: deviceDefault, alphanumeric, numeric');
        }
        
        return { isValid: errors.length === 0, errors, warnings };
      }
    },
    Security: {
      requiredFields: [],
      validate: (settings: any): PolicyValidationResult => {
        const errors: string[] = [];
        const warnings: string[] = [];
        
        if (!settings.templateId) {
          warnings.push('Security policies typically require a templateId. Using default security baseline template.');
        }
        
        return { isValid: errors.length === 0, errors, warnings };
      }
    },
    Update: {
      requiredFields: [],
      validate: (settings: any): PolicyValidationResult => {
        const errors: string[] = [];
        const warnings: string[] = [];
        
        if (settings.qualityUpdatesDeferralPeriodInDays && 
            (settings.qualityUpdatesDeferralPeriodInDays < 0 || settings.qualityUpdatesDeferralPeriodInDays > 30)) {
          errors.push('qualityUpdatesDeferralPeriodInDays must be between 0 and 30');
        }
        
        if (settings.featureUpdatesDeferralPeriodInDays && 
            (settings.featureUpdatesDeferralPeriodInDays < 0 || settings.featureUpdatesDeferralPeriodInDays > 365)) {
          errors.push('featureUpdatesDeferralPeriodInDays must be between 0 and 365');
        }
        
        return { isValid: errors.length === 0, errors, warnings };
      }
    },
    AppProtection: {
      requiredFields: [],
      validate: (settings: any): PolicyValidationResult => {
        const errors: string[] = [];
        const warnings: string[] = [];
        
        return { isValid: errors.length === 0, errors, warnings };
      }
    },
    EndpointSecurity: {
      requiredFields: [],
      validate: (settings: any): PolicyValidationResult => {
        const errors: string[] = [];
        const warnings: string[] = [];
        
        if (!settings.templateId) {
          warnings.push('EndpointSecurity policies typically require a templateId. Using default endpoint security template.');
        }
        
        return { isValid: errors.length === 0, errors, warnings };
      }
    }
  }
};

// Default values for policies
export const POLICY_DEFAULTS = {
  macos: {
    Configuration: {
      customConfiguration: {
        payloadFileName: 'configuration.mobileconfig',
        payload: '' // Should be provided by user
      }
    },
    Compliance: {
      compliance: {
        passwordRequired: false,
        passwordMinimumLength: 4,
        passwordMinutesOfInactivityBeforeLock: 15,
        storageRequireEncryption: true,
        systemIntegrityProtectionEnabled: true,
        firewallEnabled: true,
        gatekeeperAllowedAppSource: 'macAppStoreAndIdentifiedDevelopers'
      }
    }
  },
  windows: {
    Configuration: {
      // Basic security defaults
      passwordRequired: true,
      passwordBlockSimple: true,
      passwordMinimumLength: 8,
      defenderMonitorFileActivity: true,
      defenderScanNetworkFiles: true,
      defenderEnableScanIncomingMail: true,
      defenderEnableScanMappedNetworkDrivesDuringFullScan: true
    },
    Compliance: {
      passwordRequired: true,
      passwordBlockSimple: true,
      passwordMinimumLength: 8,
      passwordMinutesOfInactivityBeforeLock: 15,
      osMinimumVersion: '10.0.19041',
      storageRequireEncryption: true,
      activeFirewallRequired: true,
      defenderEnabled: true,
      antivirusRequired: true,
      antiSpywareRequired: true
    },
    Security: {
      templateId: 'd1174162-1dd2-4976-affc-6667049ab0ae' // Default security baseline
    },
    Update: {
      automaticUpdateMode: 'autoInstallAndRebootAtScheduledTime',
      microsoftUpdateServiceAllowed: true,
      driversExcluded: false,
      qualityUpdatesDeferralPeriodInDays: 0,
      featureUpdatesDeferralPeriodInDays: 0,
      businessReadyUpdatesOnly: 'all'
    },
    EndpointSecurity: {
      templateId: 'e044e60e-5901-41ea-92c5-87e8b6edd6bb' // Default endpoint security
    }
  }
};

// Policy templates for common scenarios
export const POLICY_TEMPLATES = {
  macos: {
    basicSecurity: {
      name: 'macOS Basic Security Policy',
      description: 'Basic security configuration for macOS devices',
      policyType: 'Compliance',
      settings: {
        compliance: {
          passwordRequired: true,
          passwordMinimumLength: 8,
          passwordMinutesOfInactivityBeforeLock: 5,
          passwordRequiredType: 'alphanumeric',
          storageRequireEncryption: true,
          osMinimumVersion: '14.0',
          systemIntegrityProtectionEnabled: true,
          firewallEnabled: true,
          firewallBlockAllIncoming: false,
          firewallEnableStealthMode: true,
          gatekeeperAllowedAppSource: 'macAppStoreAndIdentifiedDevelopers'
        }
      }
    },
    strictSecurity: {
      name: 'macOS Strict Security Policy',
      description: 'Strict security configuration for high-security macOS devices',
      policyType: 'Compliance',
      settings: {
        compliance: {
          passwordRequired: true,
          passwordMinimumLength: 12,
          passwordMinutesOfInactivityBeforeLock: 2,
          passwordMinutesOfInactivityBeforeScreenTimeout: 2,
          passwordPreviousPasswordBlockCount: 5,
          passwordRequiredType: 'alphanumeric',
          storageRequireEncryption: true,
          osMinimumVersion: '14.0',
          systemIntegrityProtectionEnabled: true,
          firewallEnabled: true,
          firewallBlockAllIncoming: true,
          firewallEnableStealthMode: true,
          gatekeeperAllowedAppSource: 'macAppStore',
          deviceThreatProtectionEnabled: true,
          deviceThreatProtectionRequiredSecurityLevel: 'secured'
        }
      }
    }
  },
  windows: {
    basicSecurity: {
      name: 'Windows Basic Security Policy',
      description: 'Basic security configuration for Windows devices',
      policyType: 'Compliance',
      settings: {
        passwordRequired: true,
        passwordBlockSimple: true,
        passwordMinimumLength: 8,
        passwordMinutesOfInactivityBeforeLock: 15,
        passwordExpirationDays: 90,
        passwordPreviousPasswordBlockCount: 5,
        passwordRequiredType: 'alphanumeric',
        requireHealthyDeviceReport: true,
        osMinimumVersion: '10.0.19041',
        bitLockerEnabled: true,
        secureBootEnabled: true,
        codeIntegrityEnabled: true,
        storageRequireEncryption: true,
        activeFirewallRequired: true,
        defenderEnabled: true,
        antivirusRequired: true,
        antiSpywareRequired: true,
        rtpEnabled: true
      }
    },
    strictSecurity: {
      name: 'Windows Strict Security Policy',
      description: 'Strict security configuration for high-security Windows devices',
      policyType: 'Compliance',
      settings: {
        passwordRequired: true,
        passwordBlockSimple: true,
        passwordMinimumLength: 14,
        passwordMinutesOfInactivityBeforeLock: 5,
        passwordExpirationDays: 60,
        passwordMinimumCharacterSetCount: 3,
        passwordPreviousPasswordBlockCount: 12,
        passwordRequiredType: 'alphanumeric',
        requireHealthyDeviceReport: true,
        osMinimumVersion: '10.0.22000', // Windows 11
        earlyLaunchAntiMalwareDriverEnabled: true,
        bitLockerEnabled: true,
        secureBootEnabled: true,
        codeIntegrityEnabled: true,
        storageRequireEncryption: true,
        activeFirewallRequired: true,
        defenderEnabled: true,
        antivirusRequired: true,
        antiSpywareRequired: true,
        rtpEnabled: true,
        signatureOutOfDate: false
      }
    },
    windowsUpdate: {
      name: 'Windows Update Management Policy',
      description: 'Standard Windows Update configuration',
      policyType: 'Update',
      settings: {
        automaticUpdateMode: 'autoInstallAndRebootAtScheduledTime',
        microsoftUpdateServiceAllowed: true,
        driversExcluded: false,
        qualityUpdatesDeferralPeriodInDays: 7,
        featureUpdatesDeferralPeriodInDays: 30,
        businessReadyUpdatesOnly: 'businessReadyOnly',
        prereleaseFeatures: 'notAllowed',
        deliveryOptimizationMode: 'httpWithPeeringNat',
        installationSchedule: {
          scheduledInstallDay: 'sunday',
          scheduledInstallTime: '03:00',
          activeHoursStart: '08:00',
          activeHoursEnd: '17:00'
        }
      }
    }
  }
};

// Helper functions
function isValidBase64(str: string): boolean {
  try {
    return btoa(atob(str)) === str;
  } catch (err) {
    return false;
  }
}

function checkRequiredFields(obj: any, paths: string[]): string[] {
  const missing: string[] = [];
  
  for (const path of paths) {
    const parts = path.split('.');
    let current = obj;
    
    for (let i = 0; i < parts.length; i++) {
      if (!current || !(parts[i] in current)) {
        missing.push(path);
        break;
      }
      current = current[parts[i]];
    }
  }
  
  return missing;
}

// Main validation function
export function validatePolicySettings(
  platform: string, 
  policyType: string, 
  settings: any
): PolicyValidationResult {
  const platformRules = POLICY_VALIDATION_RULES[platform as keyof typeof POLICY_VALIDATION_RULES];
  
  if (!platformRules) {
    return {
      isValid: false,
      errors: [`Unsupported platform: ${platform}`],
      warnings: []
    };
  }
  
  const typeRules = platformRules[policyType as keyof typeof platformRules];
  
  if (!typeRules) {
    return {
      isValid: false,
      errors: [`Unsupported policy type for ${platform}: ${policyType}`],
      warnings: []
    };
  }
  
  // Check required fields
  const missingFields = checkRequiredFields(settings, typeRules.requiredFields);
  if (missingFields.length > 0) {
    return {
      isValid: false,
      errors: missingFields.map(field => `Required field missing: ${field}`),
      warnings: []
    };
  }
  
  // Run specific validation
  return typeRules.validate(settings);
}

// Apply defaults to settings
export function applyPolicyDefaults(
  platform: string,
  policyType: string,
  settings: any
): any {
  const platformDefaults = POLICY_DEFAULTS[platform as keyof typeof POLICY_DEFAULTS];
  
  if (!platformDefaults) {
    return settings;
  }
  
  const typeDefaults = platformDefaults[policyType as keyof typeof platformDefaults];
  
  if (!typeDefaults) {
    return settings;
  }
  
  // Deep merge defaults with provided settings
  return deepMerge(typeDefaults, settings);
}

// Deep merge helper
function deepMerge(target: any, source: any): any {
  const output = { ...target };
  
  if (isObject(target) && isObject(source)) {
    Object.keys(source).forEach(key => {
      if (isObject(source[key])) {
        if (!(key in target)) {
          Object.assign(output, { [key]: source[key] });
        } else {
          output[key] = deepMerge(target[key], source[key]);
        }
      } else {
        Object.assign(output, { [key]: source[key] });
      }
    });
  }
  
  return output;
}

function isObject(item: any): boolean {
  return item && typeof item === 'object' && !Array.isArray(item);
}

// Generate policy creation example
export function generatePolicyExample(
  platform: string,
  policyType: string
): string {
  const examples: any = {
    macos: {
      Configuration: {
        platform: 'macos',
        policyType: 'Configuration',
        displayName: 'macOS Configuration Policy',
        description: 'Custom configuration for macOS devices',
        settings: {
          customConfiguration: {
            payloadFileName: 'custom.mobileconfig',
            payload: 'BASE64_ENCODED_MOBILECONFIG_CONTENT'
          }
        },
        assignments: [
          {
            target: {
              groupId: 'GROUP_ID_HERE'
            },
            intent: 'apply'
          }
        ]
      },
      Compliance: {
        platform: 'macos',
        policyType: 'Compliance',
        displayName: 'macOS Compliance Policy',
        description: 'Compliance requirements for macOS devices',
        settings: {
          compliance: {
            passwordRequired: true,
            passwordMinimumLength: 8,
            storageRequireEncryption: true,
            firewallEnabled: true,
            osMinimumVersion: '14.0'
          }
        }
      }
    },
    windows: {
      Configuration: {
        platform: 'windows',
        policyType: 'Configuration',
        displayName: 'Windows Configuration Policy',
        description: 'Configuration settings for Windows devices',
        settings: {
          passwordRequired: true,
          passwordMinimumLength: 8,
          defenderEnabled: true,
          firewallEnabled: true
        }
      },
      Compliance: {
        platform: 'windows',
        policyType: 'Compliance',
        displayName: 'Windows Compliance Policy',
        description: 'Compliance requirements for Windows devices',
        settings: {
          passwordRequired: true,
          passwordMinimumLength: 8,
          bitLockerEnabled: true,
          storageRequireEncryption: true,
          osMinimumVersion: '10.0.19041'
        }
      },
      Update: {
        platform: 'windows',
        policyType: 'Update',
        displayName: 'Windows Update Policy',
        description: 'Update configuration for Windows devices',
        settings: {
          automaticUpdateMode: 'autoInstallAndRebootAtScheduledTime',
          qualityUpdatesDeferralPeriodInDays: 7,
          featureUpdatesDeferralPeriodInDays: 30
        }
      }
    }
  };
  
  const example = examples[platform]?.[policyType];
  
  if (!example) {
    return `No example available for ${platform} ${policyType} policy`;
  }
  
  return JSON.stringify(example, null, 2);
}

// Validate assignments
export function validateAssignments(assignments: any[]): PolicyValidationResult {
  const errors: string[] = [];
  const warnings: string[] = [];
  
  if (!Array.isArray(assignments)) {
    errors.push('Assignments must be an array');
    return { isValid: false, errors, warnings };
  }
  
  assignments.forEach((assignment, index) => {
    if (!assignment.target) {
      errors.push(`Assignment ${index}: missing target`);
    } else {
      if (!assignment.target.groupId && !assignment.target.deviceAndAppManagementAssignmentFilterId) {
        warnings.push(`Assignment ${index}: no groupId or filter specified. Policy may not be assigned.`);
      }
    }
    
    if (assignment.intent && !['apply', 'remove'].includes(assignment.intent)) {
      errors.push(`Assignment ${index}: invalid intent '${assignment.intent}'. Must be 'apply' or 'remove'`);
    }
  });
  
  return { isValid: errors.length === 0, errors, warnings };
}
