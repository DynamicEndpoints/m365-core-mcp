# Intune Policy Creation Enhancement

This document outlines the comprehensive enhancements made to ensure accurate Intune policy creation in the Microsoft 365 Core MCP server.

## Overview

The enhanced policy creation system provides robust validation, default application, template-based creation, and comprehensive error handling to ensure accurate and reliable Intune policy creation for both Windows and macOS platforms.

## Key Components

### 1. Policy Validator (`src/validators/intune-policy-validator.ts`)

**Purpose:** Comprehensive validation system for policy settings, assignments, and platform-specific requirements.

**Features:**
- Platform and policy type validation
- Required field checking
- Settings structure validation
- Assignment validation
- Default value application
- Template management
- Example generation

**Key Functions:**
```typescript
validatePolicySettings(platform, policyType, settings) // Validates policy settings
applyPolicyDefaults(platform, policyType, settings)    // Applies sensible defaults
validateAssignments(assignments)                        // Validates policy assignments
generatePolicyExample(platform, policyType)            // Generates usage examples
```

### 2. Enhanced Policy Handler (`src/handlers/intune-handler-enhanced.ts`)

**Purpose:** Enhanced policy creation handler with comprehensive validation and error handling.

**Features:**
- Pre-flight validation
- Template application
- Enhanced error messages with context
- Validation info in response
- Support for all policy types

**Key Functions:**
```typescript
handleCreateIntunePolicyEnhanced(graphClient, args)  // Main enhanced handler
listPolicyTemplates(platform?)                       // List available templates
getPolicyCreationHelp(platform, policyType)         // Get creation help
```

### 3. Enhanced Tool Definitions (`src/tool-definitions-intune-enhanced.ts`)

**Purpose:** Type-safe tool definitions with discriminated unions for platform-specific validation.

**Features:**
- Platform-specific settings schemas
- Discriminated unions for type safety
- Detailed field validation
- Comprehensive descriptions
- Backwards compatibility

## Policy Types and Platforms

### macOS Support

#### Configuration Policies
- **Required Fields:** `customConfiguration.payloadFileName`, `customConfiguration.payload`
- **Format:** Base64 encoded .mobileconfig files
- **Validation:** Base64 format validation, filename requirements

#### Compliance Policies
- **Settings:** Password, encryption, firewall, SIP, Gatekeeper, OS version
- **Validation:** Range validation for numeric fields, enum validation for options
- **Defaults:** Secure baseline settings applied automatically

#### Security/Update/AppProtection
- **Status:** Basic structure support
- **Recommendation:** Use Configuration policies with appropriate .mobileconfig files

### Windows Support

#### Configuration Policies
- **Settings:** Password, Defender, firewall, device restrictions
- **Validation:** Range validation, boolean flags
- **Defaults:** Security-focused baseline settings

#### Compliance Policies
- **Settings:** Password, BitLocker, OS version, security features
- **Validation:** Comprehensive range and enum validation
- **Defaults:** Enterprise security baseline

#### Update Policies
- **Settings:** Automatic updates, deferral periods, installation schedules
- **Validation:** Range validation (0-30 days quality, 0-365 days feature)
- **Defaults:** Balanced update management

#### Security/EndpointSecurity Policies
- **Implementation:** Template-based using Microsoft Graph intents API
- **Validation:** Template ID validation
- **Defaults:** Standard security baseline templates

## Templates

### macOS Templates

#### basicSecurity
```json
{
  "name": "macOS Basic Security Policy",
  "policyType": "Compliance",
  "settings": {
    "compliance": {
      "passwordRequired": true,
      "passwordMinimumLength": 8,
      "passwordMinutesOfInactivityBeforeLock": 5,
      "storageRequireEncryption": true,
      "firewallEnabled": true,
      "gatekeeperAllowedAppSource": "macAppStoreAndIdentifiedDevelopers"
    }
  }
}
```

#### strictSecurity
```json
{
  "name": "macOS Strict Security Policy",
  "policyType": "Compliance",
  "settings": {
    "compliance": {
      "passwordRequired": true,
      "passwordMinimumLength": 12,
      "passwordMinutesOfInactivityBeforeLock": 2,
      "passwordPreviousPasswordBlockCount": 5,
      "storageRequireEncryption": true,
      "firewallBlockAllIncoming": true,
      "gatekeeperAllowedAppSource": "macAppStore",
      "deviceThreatProtectionEnabled": true
    }
  }
}
```

### Windows Templates

#### basicSecurity
```json
{
  "name": "Windows Basic Security Policy",
  "policyType": "Compliance",
  "settings": {
    "passwordRequired": true,
    "passwordMinimumLength": 8,
    "bitLockerEnabled": true,
    "storageRequireEncryption": true,
    "activeFirewallRequired": true,
    "defenderEnabled": true
  }
}
```

#### strictSecurity
```json
{
  "name": "Windows Strict Security Policy",
  "policyType": "Compliance",
  "settings": {
    "passwordRequired": true,
    "passwordMinimumLength": 14,
    "passwordMinutesOfInactivityBeforeLock": 5,
    "passwordExpirationDays": 60,
    "passwordMinimumCharacterSetCount": 3,
    "osMinimumVersion": "10.0.22000",
    "bitLockerEnabled": true,
    "secureBootEnabled": true
  }
}
```

#### windowsUpdate
```json
{
  "name": "Windows Update Management Policy",
  "policyType": "Update",
  "settings": {
    "automaticUpdateMode": "autoInstallAndRebootAtScheduledTime",
    "qualityUpdatesDeferralPeriodInDays": 7,
    "featureUpdatesDeferralPeriodInDays": 30,
    "businessReadyUpdatesOnly": "businessReadyOnly",
    "installationSchedule": {
      "scheduledInstallDay": "sunday",
      "scheduledInstallTime": "03:00"
    }
  }
}
```

## Validation Rules

### Common Validation Rules
1. **Platform Validation:** Must be 'macos' or 'windows'
2. **Policy Type Validation:** Must be valid for the platform
3. **Display Name:** Required, 1-256 characters
4. **Description:** Optional, max 1024 characters

### Platform-Specific Rules

#### macOS Configuration
- `customConfiguration.payloadFileName` is required
- `customConfiguration.payload` is required and should be base64 encoded
- Warning if payload is not valid base64

#### macOS Compliance
- `passwordMinimumLength` must be 4-14 if specified
- `passwordRequiredType` must be valid enum value
- `gatekeeperAllowedAppSource` must be valid enum value

#### Windows Compliance
- `passwordMinimumLength` must be 0-127 if specified
- `passwordExpirationDays` must be 0-730 if specified
- `passwordPreviousPasswordBlockCount` must be 0-50 if specified

#### Windows Update
- `qualityUpdatesDeferralPeriodInDays` must be 0-30
- `featureUpdatesDeferralPeriodInDays` must be 0-365

### Assignment Validation
- Assignments must be an array if provided
- Each assignment must have a target object
- Target must have either groupId or filter specified
- Intent must be 'apply' or 'remove' if specified

## Error Handling

### Validation Errors
- Clear, specific error messages
- Field-level validation feedback
- Examples of correct usage included in error messages
- Context information (platform, policy type, policy name)

### Enhanced Error Messages
```
Policy validation failed:
Required field missing: customConfiguration.payloadFileName

Example of valid settings:
{
  "platform": "macos",
  "policyType": "Configuration",
  "displayName": "macOS Configuration Policy",
  "settings": {
    "customConfiguration": {
      "payloadFileName": "custom.mobileconfig",
      "payload": "BASE64_ENCODED_MOBILECONFIG_CONTENT"
    }
  }
}
```

## Usage Examples

### Basic Policy Creation
```javascript
// Using the enhanced create_intune_policy tool
{
  "platform": "windows",
  "policyType": "Compliance",
  "displayName": "Windows Security Policy",
  "description": "Basic security requirements for Windows devices",
  "settings": {
    "passwordRequired": true,
    "passwordMinimumLength": 8,
    "bitLockerEnabled": true,
    "storageRequireEncryption": true
  },
  "assignments": [
    {
      "target": {
        "groupId": "12345678-1234-1234-1234-123456789012"
      },
      "intent": "apply"
    }
  ]
}
```

### Template-Based Creation
```javascript
// Using a predefined template
{
  "platform": "macos",
  "policyType": "Compliance",
  "displayName": "macOS Basic Security",
  "useTemplate": "basicSecurity",
  "assignments": [
    {
      "target": {
        "groupId": "87654321-4321-4321-4321-210987654321"
      }
    }
  ]
}
```

### macOS Configuration Policy
```javascript
{
  "platform": "macos",
  "policyType": "Configuration",
  "displayName": "Custom macOS Configuration",
  "settings": {
    "customConfiguration": {
      "payloadFileName": "security-config.mobileconfig",
      "payload": "PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz4..."
    }
  }
}
```

## Default Values

### Automatic Defaults Applied

#### macOS Compliance Defaults
- `passwordRequired`: false
- `passwordMinimumLength`: 4
- `passwordMinutesOfInactivityBeforeLock`: 15
- `storageRequireEncryption`: true
- `systemIntegrityProtectionEnabled`: true
- `firewallEnabled`: true
- `gatekeeperAllowedAppSource`: "macAppStoreAndIdentifiedDevelopers"

#### Windows Configuration Defaults
- `passwordRequired`: true
- `passwordBlockSimple`: true
- `passwordMinimumLength`: 8
- `defenderMonitorFileActivity`: true
- `defenderScanNetworkFiles`: true
- `defenderEnableScanIncomingMail`: true

#### Windows Compliance Defaults
- `passwordRequired`: true
- `passwordBlockSimple`: true
- `passwordMinimumLength`: 8
- `passwordMinutesOfInactivityBeforeLock`: 15
- `osMinimumVersion`: "10.0.19041"
- `storageRequireEncryption`: true
- `activeFirewallRequired`: true
- `defenderEnabled`: true
- `antivirusRequired`: true
- `antiSpywareRequired`: true

## Testing

### Test Coverage
- ✅ Validation function testing
- ✅ Default application testing
- ✅ Template functionality testing
- ✅ Example generation testing
- ✅ Error scenario testing
- ✅ Assignment validation testing
- ✅ Platform-specific validation testing

### Running Tests
```bash
node test-enhanced-policy-creation.mjs
```

## Implementation Status

### ✅ Complete Features
- [x] Comprehensive validation system
- [x] Default value application
- [x] Template-based policy creation
- [x] Enhanced error handling
- [x] Assignment validation
- [x] Example generation
- [x] Help text generation
- [x] Type-safe tool definitions
- [x] Test suite

### 🔄 Ongoing Improvements
- Enhanced template library
- Additional policy type support
- Custom validation rules
- Policy migration utilities

## Benefits

### For Developers
1. **Type Safety:** Comprehensive TypeScript definitions prevent runtime errors
2. **Clear Feedback:** Detailed validation messages guide correct usage
3. **Templates:** Pre-built configurations for common scenarios
4. **Examples:** Auto-generated examples for each policy type

### For End Users
1. **Reliability:** Policies are validated before creation
2. **Consistency:** Default values ensure consistent baseline security
3. **Flexibility:** Support for custom configurations and templates
4. **Error Prevention:** Validation prevents invalid policy creation

### For Operations
1. **Reduced Errors:** Pre-flight validation catches issues early
2. **Standardization:** Templates promote consistent policy deployment
3. **Documentation:** Built-in help and examples reduce support burden
4. **Monitoring:** Enhanced logging and error reporting

## Migration Guide

### From Basic to Enhanced Handler

1. **Update Imports:**
```typescript
// Old
import { handleCreateIntunePolicy } from './src/handlers/intune-handler.js';

// New
import { handleCreateIntunePolicy } from './src/handlers/intune-handler-enhanced.js';
```

2. **Use Enhanced Tool Definition:**
```typescript
// Import enhanced tool definitions
import { enhancedIntuneTools } from './src/tool-definitions-intune-enhanced.js';
```

3. **Leverage New Features:**
- Use `useTemplate` parameter for quick policy creation
- Check `validationInfo` in responses for warnings
- Use enhanced error messages for troubleshooting

## Future Enhancements

1. **Policy Comparison:** Compare existing policies with templates
2. **Bulk Operations:** Create multiple policies from templates
3. **Policy Testing:** Validate policies against test devices
4. **Compliance Reporting:** Generate compliance reports based on policies
5. **Custom Templates:** Allow users to create and save custom templates

## Conclusion

The enhanced Intune policy creation system provides a robust, validated, and user-friendly approach to managing device policies. With comprehensive validation, sensible defaults, template support, and clear error handling, it ensures accurate policy creation while reducing the complexity and potential for errors.
