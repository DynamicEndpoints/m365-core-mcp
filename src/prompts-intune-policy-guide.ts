/**
 * Intune Policy Creation Prompts
 * Specialized prompts to guide accurate policy creation
 */

import { Client } from '@microsoft/microsoft-graph-client';
import {
  SETTINGS_CATALOG_POLICY_TEMPLATES,
  PPC_POLICY_TEMPLATES,
  COMMON_SETTINGS_DEFINITIONS
} from './handlers/intune-policy-templates.js';

export interface IntunePolicyPrompt {
  name: string;
  description: string;
  arguments?: {
    name: string;
    description: string;
    required: boolean;
  }[];
  handler: (graphClient: Client, args: Record<string, string>) => Promise<string>;
}

export const intunePolicyPrompts: IntunePolicyPrompt[] = [
  {
    name: 'intune_policy_wizard',
    description: 'Interactive wizard to guide creation of Intune policies with correct structure and settings',
    arguments: [
      {
        name: 'policy_goal',
        description: 'What you want to accomplish (e.g., "enforce encryption", "configure updates", "block threats")',
        required: true
      },
      {
        name: 'platform',
        description: 'Target platform: windows or macos',
        required: true
      },
      {
        name: 'security_level',
        description: 'Security level: basic, standard, high, or critical',
        required: false
      }
    ],
    handler: async (graphClient: Client, args: Record<string, string>) => {
      const { policy_goal, platform, security_level = 'standard' } = args;
      
      // Analyze the goal and recommend templates
      const recommendations: string[] = [];
      let templateGuide = '';
      
      // Match goal to templates
      if (policy_goal.toLowerCase().includes('encrypt') || policy_goal.toLowerCase().includes('bitlocker')) {
        recommendations.push('BITLOCKER_ENCRYPTION');
        templateGuide += `
## BitLocker Encryption Policy

**Recommended Template**: BITLOCKER_ENCRYPTION

**Usage**:
\`\`\`javascript
{
  action: 'create_settings_catalog',
  settingsCatalogTemplate: 'BITLOCKER_ENCRYPTION',
  name: 'Corporate BitLocker Policy',
  description: 'Enforce BitLocker encryption on Windows devices',
  assignments: [
    {
      target: { groupId: '<your-group-id>' }
    }
  ]
}
\`\`\`

**What this configures**:
- Requires device encryption
- Full disk encryption for fixed drives
- Encryption for removable drives
- Automatic recovery key backup

**Graph API Endpoint**: \`/deviceManagement/configurationPolicies\`
`;
      }
      
      if (policy_goal.toLowerCase().includes('update') || policy_goal.toLowerCase().includes('patch')) {
        recommendations.push('WINDOWS_UPDATE');
        templateGuide += `
## Windows Update Policy

**Recommended Template**: WINDOWS_UPDATE

**Usage**:
\`\`\`javascript
{
  action: 'create_settings_catalog',
  settingsCatalogTemplate: 'WINDOWS_UPDATE',
  settingsCatalogParams: {
    deferQualityDays: 7,      // Defer quality updates by 7 days
    deferFeatureDays: 14       // Defer feature updates by 14 days
  },
  name: 'Windows Update Policy',
  description: 'Configure Windows Update with deferral periods'
}
\`\`\`

**Customizable Parameters**:
- \`deferQualityDays\` (0-30): Days to defer security updates
- \`deferFeatureDays\` (0-365): Days to defer feature updates

**Security Recommendations by Level**:
- Basic: 7 days quality, 30 days feature
- Standard: 5 days quality, 14 days feature  
- High: 3 days quality, 7 days feature
- Critical: 0 days quality, 0 days feature
`;
      }
      
      if (policy_goal.toLowerCase().includes('threat') || policy_goal.toLowerCase().includes('attack') || 
          policy_goal.toLowerCase().includes('malware') || policy_goal.toLowerCase().includes('virus')) {
        recommendations.push('DEFENDER_ANTIVIRUS', 'ATTACK_SURFACE_REDUCTION', 'ATTACK_SURFACE_REDUCTION_PPC');
        templateGuide += `
## Threat Protection Policies

**Recommended Templates**:
1. DEFENDER_ANTIVIRUS (Settings Catalog)
2. ATTACK_SURFACE_REDUCTION (Settings Catalog)
3. ATTACK_SURFACE_REDUCTION_PPC (Platform Protection)

### 1. Windows Defender Antivirus
\`\`\`javascript
{
  action: 'create_settings_catalog',
  settingsCatalogTemplate: 'DEFENDER_ANTIVIRUS',
  name: 'Defender Antivirus Policy'
}
\`\`\`

**Configures**:
- Real-time protection
- Cloud-delivered protection
- Behavior monitoring
- Full scan schedule

### 2. Attack Surface Reduction (Settings Catalog)
\`\`\`javascript
{
  action: 'create_settings_catalog',
  settingsCatalogTemplate: 'ATTACK_SURFACE_REDUCTION',
  name: 'ASR Rules Policy'
}
\`\`\`

**ASR Rules Included**:
- Block executable content from email/webmail
- Block Office apps from creating executables
- Block Office apps from injecting code
- Block JavaScript/VBScript from launching downloads
- Block obfuscated scripts

### 3. Attack Surface Reduction (PPC)
\`\`\`javascript
{
  action: 'create_ppc',
  ppcTemplate: 'ATTACK_SURFACE_REDUCTION_PPC',
  name: 'ASR Protection Policy'
}
\`\`\`

**Additional PPC Settings**:
- Block Win32 API calls from Office macros
- Block untrusted/unsigned processes
- Block credential stealing from LSASS
- Block Adobe Reader child processes
- Block persistence through WMI events
`;
      }
      
      if (policy_goal.toLowerCase().includes('firewall') || policy_goal.toLowerCase().includes('network')) {
        recommendations.push('FIREWALL_CONFIGURATION');
        templateGuide += `
## Firewall Configuration

**Recommended Template**: FIREWALL_CONFIGURATION

**Usage**:
\`\`\`javascript
{
  action: 'create_settings_catalog',
  settingsCatalogTemplate: 'FIREWALL_CONFIGURATION',
  name: 'Windows Firewall Policy'
}
\`\`\`

**Configures**:
- Enables firewall for Domain profile
- Enables firewall for Public profile
- Enables firewall for Private profile

**Important**: This is a basic configuration. For advanced rules, use custom settings.
`;
      }
      
      if (policy_goal.toLowerCase().includes('password') || policy_goal.toLowerCase().includes('authentication')) {
        recommendations.push('PASSWORD_POLICY');
        templateGuide += `
## Password Policy

**Recommended Template**: PASSWORD_POLICY

**Usage**:
\`\`\`javascript
{
  action: 'create_settings_catalog',
  settingsCatalogTemplate: 'PASSWORD_POLICY',
  settingsCatalogParams: {
    minLength: 12,           // Minimum password length
    complexity: 3            // Complexity level (0-4)
  },
  name: 'Password Requirements Policy'
}
\`\`\`

**Complexity Levels**:
- 0: No complexity requirements
- 1: Require digits
- 2: Require lowercase and uppercase
- 3: Require digits, lowercase, and uppercase
- 4: Require digits, lowercase, uppercase, and special characters

**Default Settings**:
- Password expiration: 90 days
- Password history: 24 passwords
`;
      }
      
      // Add exploit protection if high security
      if (security_level === 'high' || security_level === 'critical') {
        recommendations.push('EXPLOIT_PROTECTION_PPC', 'WEB_PROTECTION_PPC');
        templateGuide += `
## High Security: Additional Protections

### Exploit Protection (PPC)
\`\`\`javascript
{
  action: 'create_ppc',
  ppcTemplate: 'EXPLOIT_PROTECTION_PPC',
  name: 'Exploit Protection Policy'
}
\`\`\`

**Protections Enabled**:
- Data Execution Prevention (DEP)
- Control Flow Guard (CFG)
- Randomize memory allocations (ASLR)
- Validate exception chains (SEHOP)
- Validate stack integrity (StackPivot)
- Code Integrity Guard
- Block remote image loads

### Web Protection (PPC)
\`\`\`javascript
{
  action: 'create_ppc',
  ppcTemplate: 'WEB_PROTECTION_PPC',
  name: 'Web Protection Policy'
}
\`\`\`

**Protections Enabled**:
- Network Protection (blocks malicious sites)
- SmartScreen for Edge
- Prevent SmartScreen override
- Block user feedback on warnings
`;
      }
      
      // Build the complete guide
      let guide = `# Intune Policy Creation Guide

## Goal Analysis
**Your Goal**: ${policy_goal}
**Platform**: ${platform}
**Security Level**: ${security_level}

## Recommended Templates
${recommendations.length > 0 ? recommendations.map(r => `- ${r}`).join('\n') : 'No specific templates match your goal. See custom policy section below.'}

${templateGuide}

## Important Notes

### Settings Catalog vs PPC Policies

**Settings Catalog**:
- More granular control
- Individual CSP settings
- Easier to understand and troubleshoot
- Best for: Configuration, Updates, Firewall, Passwords

**Platform Protection Configuration (PPC)**:
- Pre-configured security bundles
- Template-based approach
- Faster deployment
- Best for: Attack Surface Reduction, Exploit Protection, Web Protection

### Assignment Best Practices

1. **Start with Pilot Group**:
   \`\`\`javascript
   assignments: [{
     target: { groupId: 'pilot-group-id' }
   }]
   \`\`\`

2. **Phased Rollout**:
   - Week 1: IT Department (10-20 devices)
   - Week 2: Department leads (50-100 devices)
   - Week 3-4: Full organization

3. **Monitor Compliance**:
   Use \`action: 'get_status'\` in compliance tools

### Common Mistakes to Avoid

1. ❌ **Don't mix policy types**: Use Settings Catalog OR PPC, not both for same settings
2. ❌ **Don't over-configure**: Start with templates, customize only if needed
3. ❌ **Don't skip testing**: Always test in pilot group first
4. ❌ **Don't forget assignments**: Policies without assignments do nothing

### Validation Checklist

Before creating a policy, verify:
- [ ] Template name is correct (case-sensitive)
- [ ] Parameters match template requirements
- [ ] Assignment groups exist in Azure AD
- [ ] You have appropriate permissions
- [ ] Policy name is unique and descriptive

### Getting More Information

List all available templates:
\`\`\`javascript
{
  action: 'list_templates'
}
\`\`\`

This returns both Settings Catalog and PPC templates with descriptions.

## Custom Policy Creation

If templates don't meet your needs, you can create custom policies:

### Custom Settings Catalog
\`\`\`javascript
{
  action: 'create_settings_catalog',
  customSettingsCatalogPolicy: {
    name: 'Custom Policy',
    description: 'Custom configuration',
    platforms: 'windows10',
    technologies: 'mdm',
    settings: [
      {
        '@odata.type': '#microsoft.graph.deviceManagementConfigurationSettingInstance',
        settingInstance: {
          '@odata.type': '#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance',
          settingDefinitionId: 'device_vendor_msft_policy_config_...',
          simpleSettingValue: {
            '@odata.type': '#microsoft.graph.deviceManagementConfigurationIntegerSettingValue',
            value: 1
          }
        }
      }
    ]
  }
}
\`\`\`

### Finding Setting Definition IDs

Available in the templates module under \`COMMON_SETTINGS_DEFINITIONS\`:
- BitLocker settings
- Windows Defender settings
- Windows Update settings
- Firewall settings
- Password policy settings
- Application control settings

## Troubleshooting

### "Invalid policy structure"
- Check template name spelling (case-sensitive)
- Verify parameters match template requirements
- Use \`list_templates\` to see available options

### "Assignment failed"
- Verify group ID exists
- Check group is security-enabled
- Ensure proper permissions

### "Policy not applying"
- Check device is enrolled in Intune
- Verify device meets platform requirements
- Force sync with \`action: 'sync'\`
- Check compliance status

## Next Steps

1. Choose appropriate template(s) from recommendations above
2. Customize parameters if needed
3. Create policy with pilot assignment
4. Monitor deployment
5. Gradually expand to full organization

Need help? Ask follow-up questions about specific templates or requirements.
`;
      
      return guide;
    }
  },
  
  {
    name: 'intune_policy_troubleshoot',
    description: 'Troubleshoot common Intune policy creation and deployment issues',
    arguments: [
      {
        name: 'issue_description',
        description: 'Description of the issue you\'re experiencing',
        required: true
      },
      {
        name: 'policy_type',
        description: 'Type of policy: settings_catalog, ppc, or unknown',
        required: false
      }
    ],
    handler: async (graphClient: Client, args: Record<string, string>) => {
      const { issue_description, policy_type = 'unknown' } = args;
      
      let troubleshootingGuide = `# Intune Policy Troubleshooting Guide

## Issue Description
${issue_description}

## Policy Type
${policy_type}

## Common Issues and Solutions

`;

      // Analyze the issue and provide specific guidance
      const issue = issue_description.toLowerCase();
      
      if (issue.includes('invalid') || issue.includes('error') || issue.includes('failed')) {
        troubleshootingGuide += `
### Issue: Policy Creation Failed

**Possible Causes**:
1. **Incorrect Template Name**: Template names are case-sensitive
   - ✅ Correct: \`BITLOCKER_ENCRYPTION\`
   - ❌ Incorrect: \`bitlocker_encryption\` or \`BitLocker_Encryption\`

2. **Invalid Parameters**: Some templates require specific parameters
   - WINDOWS_UPDATE: needs \`deferQualityDays\` and \`deferFeatureDays\`
   - PASSWORD_POLICY: needs \`minLength\` and \`complexity\`

3. **Incorrect Policy Structure**: Custom policies must follow exact schema
   - Use validation functions before creation
   - Check \`@odata.type\` values are correct

4. **Permission Issues**: Check your app registration has these permissions:
   - DeviceManagementConfiguration.ReadWrite.All
   - DeviceManagementManagedDevices.ReadWrite.All
   - DeviceManagementServiceConfig.ReadWrite.All

**Solutions**:
\`\`\`javascript
// 1. Verify template exists
{
  action: 'list_templates'
}

// 2. Use correct case and structure
{
  action: 'create_settings_catalog',
  settingsCatalogTemplate: 'BITLOCKER_ENCRYPTION',  // Exact case
  name: 'Test Policy',
  description: 'Test description'
}

// 3. Include required parameters
{
  action: 'create_settings_catalog',
  settingsCatalogTemplate: 'WINDOWS_UPDATE',
  settingsCatalogParams: {
    deferQualityDays: 7,
    deferFeatureDays: 14
  }
}
\`\`\`
`;
      }
      
      if (issue.includes('assign') || issue.includes('deployment')) {
        troubleshootingGuide += `
### Issue: Assignment/Deployment Problems

**Possible Causes**:
1. **Invalid Group ID**: Group doesn't exist or ID is incorrect
2. **Group Type**: Must be security-enabled group
3. **Timing**: Assignment takes 5-10 minutes to process
4. **Device Not Enrolled**: Device must be enrolled in Intune

**Solutions**:
\`\`\`javascript
// 1. Verify group exists (use Graph Explorer or Azure Portal)
// 2. Create policy first, then assign separately
{
  action: 'create_settings_catalog',
  settingsCatalogTemplate: 'BITLOCKER_ENCRYPTION',
  name: 'BitLocker Policy'
  // Don't include assignments yet
}

// 3. After policy creation, assign to specific group
{
  action: 'assign',
  policyId: '<policy-id-from-creation>',
  assignments: [{
    target: {
      groupId: '<verified-group-id>'
    }
  }]
}

// 4. Check device enrollment status
{
  action: 'list',
  filter: 'deviceName eq \\'device-name\\''
}
\`\`\`
`;
      }
      
      if (issue.includes('not applying') || issue.includes('not working')) {
        troubleshootingGuide += `
### Issue: Policy Not Applying to Devices

**Diagnostic Steps**:

1. **Check Device Sync Status**:
\`\`\`javascript
{
  action: 'get',
  deviceId: '<device-id>'
}
// Look for lastSyncDateTime - should be recent
\`\`\`

2. **Force Device Sync**:
\`\`\`javascript
{
  action: 'sync',
  deviceId: '<device-id>'
}
\`\`\`

3. **Check Compliance Status**:
\`\`\`javascript
{
  action: 'get_status',
  deviceId: '<device-id>'
}
\`\`\`

4. **Check Policy Assignment**:
\`\`\`javascript
{
  action: 'get',
  policyId: '<policy-id>'
}
// Verify assignments array includes target device's group
\`\`\`

**Common Fixes**:
- Wait 8 hours for natural sync cycle
- Force sync from device (Settings > Accounts > Access work or school > Info > Sync)
- Check device meets policy requirements (OS version, etc.)
- Verify no conflicting policies
- Check device is in assigned group
`;
      }
      
      if (issue.includes('settings catalog') || issue.includes('setting definition')) {
        troubleshootingGuide += `
### Issue: Settings Catalog Configuration

**Finding Correct Setting Definition IDs**:

Available pre-defined settings in \`COMMON_SETTINGS_DEFINITIONS\`:

**BitLocker**:
- BITLOCKER_REQUIRE_DEVICE_ENCRYPTION
- BITLOCKER_FIXED_DRIVE_ENCRYPTION_TYPE
- BITLOCKER_REMOVABLE_DRIVE_ENCRYPTION_TYPE

**Windows Defender**:
- DEFENDER_REAL_TIME_PROTECTION
- DEFENDER_CLOUD_PROTECTION
- DEFENDER_BEHAVIOR_MONITORING
- DEFENDER_SCAN_TYPE

**Windows Update**:
- UPDATE_BRANCH_READINESS_LEVEL
- UPDATE_DEFER_QUALITY_UPDATES
- UPDATE_DEFER_FEATURE_UPDATES

**Firewall**:
- FIREWALL_DOMAIN_PROFILE_ENABLED
- FIREWALL_PUBLIC_PROFILE_ENABLED
- FIREWALL_PRIVATE_PROFILE_ENABLED

**Password Policy**:
- PASSWORD_MIN_LENGTH
- PASSWORD_COMPLEXITY
- PASSWORD_EXPIRATION
- PASSWORD_HISTORY

**Application Control**:
- APPLOCKER_EXE_RULES
- APPLOCKER_DLL_RULES

**Attack Surface Reduction**:
- ASR_RULES
- ASR_ONLY_EXCLUSIONS

**Usage Example**:
\`\`\`javascript
import { COMMON_SETTINGS_DEFINITIONS } from './handlers/intune-policy-templates';

// Use in custom policy
customSettingsCatalogPolicy: {
  settings: [
    {
      settingDefinitionId: COMMON_SETTINGS_DEFINITIONS.DEFENDER_REAL_TIME_PROTECTION,
      value: true
    }
  ]
}
\`\`\`
`;
      }
      
      if (issue.includes('ppc') || issue.includes('platform protection')) {
        troubleshootingGuide += `
### Issue: Platform Protection Configuration (PPC)

**PPC vs Settings Catalog**:
- PPC uses \`/deviceManagement/intents\` API
- Settings Catalog uses \`/deviceManagement/configurationPolicies\` API
- Don't mix both for same settings

**Available PPC Templates**:
1. ATTACK_SURFACE_REDUCTION_PPC
2. EXPLOIT_PROTECTION_PPC
3. WEB_PROTECTION_PPC

**Common PPC Issues**:

1. **Template ID Not Found**:
   - PPC templates are specific to tenant
   - Template IDs in code may need updating
   - Solution: Use pre-built templates

2. **Setting Value Type Mismatch**:
   - PPC expects specific value formats
   - Use templates to ensure correct structure

3. **Conflict with Settings Catalog**:
   - If you have both PPC and Settings Catalog configuring same setting
   - Solution: Use one or the other, not both
`;
      }
      
      troubleshootingGuide += `
## General Troubleshooting Commands

### List All Policies
\`\`\`javascript
{
  action: 'list',
  policyType: 'Configuration'  // or 'Compliance', 'Security', etc.
}
\`\`\`

### Get Specific Policy Details
\`\`\`javascript
{
  action: 'get',
  policyId: '<policy-id>',
  policyType: 'Configuration'
}
\`\`\`

### Check Device Status
\`\`\`javascript
{
  action: 'list',
  filter: 'operatingSystem eq \\'Windows\\''
}
\`\`\`

### Force Policy Refresh
\`\`\`javascript
{
  action: 'sync',
  deviceId: '<device-id>'
}
\`\`\`

## Getting Additional Help

1. **Check Microsoft Documentation**:
   - Settings Catalog: https://learn.microsoft.com/mem/intune/configuration/settings-catalog
   - Policy CSPs: https://learn.microsoft.com/windows/client-management/mdm/policy-csps

2. **Use Graph Explorer**:
   - Test API calls: https://developer.microsoft.com/graph/graph-explorer

3. **Check Intune Admin Center**:
   - Devices > Monitor > Device assignment failures
   - Devices > Configuration profiles > Monitor

4. **Review Permissions**:
   Required Microsoft Graph permissions:
   - DeviceManagementConfiguration.ReadWrite.All
   - DeviceManagementManagedDevices.ReadWrite.All
   - DeviceManagementServiceConfig.ReadWrite.All

## Still Having Issues?

If the issue persists:
1. Enable verbose logging in Intune
2. Check Windows Event Viewer on client device
3. Review device diagnostic logs
4. Contact Microsoft Support with policy ID and error details
`;
      
      return troubleshootingGuide;
    }
  }
];
