import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';

// Policy types that can be backed up
export type PolicyBackupType = 
  | 'conditionalAccess'
  | 'namedLocations' 
  | 'authenticationStrengths'
  | 'deviceCompliancePolicies'
  | 'deviceConfigurationPolicies'
  | 'appProtectionPolicies'
  | 'dlpPolicies'
  | 'retentionPolicies'
  | 'sensitivityLabels'
  | 'all';

export interface PolicyBackupArgs {
  action: 'backup' | 'list' | 'restore';
  policyTypes?: PolicyBackupType[];
  outputFormat?: 'json' | 'summary';
  includeMetadata?: boolean;
  policyData?: string; // For restore action - JSON string of policy data
}

export interface PolicyBackupResult {
  timestamp: string;
  tenantId?: string;
  policies: {
    type: PolicyBackupType;
    count: number;
    data: any[];
  }[];
  metadata?: {
    backupVersion: string;
    exportedBy?: string;
    exportedAt: string;
  };
}

// Backup handler for Microsoft 365 policies
export async function handlePolicyBackup(
  graphClient: Client,
  args: PolicyBackupArgs
): Promise<{ content: { type: string; text: string }[] }> {
  
  switch (args.action) {
    case 'backup':
      return await backupPolicies(graphClient, args);
    case 'list':
      return await listBackupableTypes();
    case 'restore':
      throw new McpError(
        ErrorCode.InvalidParams,
        'Restore functionality requires careful review. Export the backup JSON first, then use individual policy management tools to restore specific policies.'
      );
    default:
      throw new McpError(ErrorCode.InvalidParams, `Unknown action: ${args.action}`);
  }
}

async function listBackupableTypes(): Promise<{ content: { type: string; text: string }[] }> {
  const backupableTypes = [
    {
      type: 'conditionalAccess',
      description: 'Conditional Access Policies - Access control policies for Azure AD',
      endpoint: '/identity/conditionalAccess/policies',
      permissions: 'Policy.Read.All'
    },
    {
      type: 'namedLocations',
      description: 'Named Locations - IP ranges and countries/regions for Conditional Access',
      endpoint: '/identity/conditionalAccess/namedLocations',
      permissions: 'Policy.Read.All'
    },
    {
      type: 'authenticationStrengths',
      description: 'Authentication Strength Policies - MFA method requirements',
      endpoint: '/identity/conditionalAccess/authenticationStrength/policies',
      permissions: 'Policy.Read.All'
    },
    {
      type: 'deviceCompliancePolicies',
      description: 'Intune Device Compliance Policies - Device health requirements',
      endpoint: '/deviceManagement/deviceCompliancePolicies',
      permissions: 'DeviceManagementConfiguration.Read.All'
    },
    {
      type: 'deviceConfigurationPolicies',
      description: 'Intune Device Configuration Policies - Device settings profiles',
      endpoint: '/deviceManagement/deviceConfigurations',
      permissions: 'DeviceManagementConfiguration.Read.All'
    },
    {
      type: 'appProtectionPolicies',
      description: 'Intune App Protection Policies - MAM policies for mobile apps',
      endpoint: '/deviceAppManagement/managedAppPolicies',
      permissions: 'DeviceManagementApps.Read.All'
    },
    {
      type: 'dlpPolicies',
      description: 'Data Loss Prevention Policies - Information protection policies (beta)',
      endpoint: '/informationProtection/policy/labels',
      permissions: 'InformationProtectionPolicy.Read.All'
    },
    {
      type: 'sensitivityLabels',
      description: 'Sensitivity Labels - Data classification labels',
      endpoint: '/security/informationProtection/sensitivityLabels',
      permissions: 'InformationProtectionPolicy.Read.All'
    }
  ];

  return {
    content: [{
      type: 'text',
      text: `# Available Policy Types for Backup\n\n${backupableTypes.map(t => 
        `## ${t.type}\n- **Description**: ${t.description}\n- **Graph Endpoint**: ${t.endpoint}\n- **Required Permissions**: ${t.permissions}\n`
      ).join('\n')}\n\n## Usage\nTo backup policies, use:\n\`\`\`json\n{\n  "action": "backup",\n  "policyTypes": ["conditionalAccess", "namedLocations"],\n  "outputFormat": "json",\n  "includeMetadata": true\n}\n\`\`\``
    }]
  };
}

async function backupPolicies(
  graphClient: Client,
  args: PolicyBackupArgs
): Promise<{ content: { type: string; text: string }[] }> {
  const policyTypes = args.policyTypes || ['all'];
  const includeAll = policyTypes.includes('all');
  const includeMetadata = args.includeMetadata ?? true;
  
  const backupResult: PolicyBackupResult = {
    timestamp: new Date().toISOString(),
    policies: [],
    metadata: includeMetadata ? {
      backupVersion: '1.0.0',
      exportedAt: new Date().toISOString()
    } : undefined
  };

  // Try to get tenant info
  try {
    const org = await graphClient.api('/organization').select('id,displayName').get();
    if (org.value && org.value.length > 0) {
      backupResult.tenantId = org.value[0].id;
      if (backupResult.metadata) {
        backupResult.metadata.exportedBy = org.value[0].displayName;
      }
    }
  } catch (e) {
    // Continue without tenant info
  }

  const errors: string[] = [];

  // Backup Conditional Access Policies
  if (includeAll || policyTypes.includes('conditionalAccess')) {
    try {
      const caPolicies = await graphClient
        .api('/identity/conditionalAccess/policies')
        .get();
      
      backupResult.policies.push({
        type: 'conditionalAccess',
        count: caPolicies.value?.length || 0,
        data: caPolicies.value || []
      });
    } catch (e: any) {
      errors.push(`conditionalAccess: ${e.message || 'Permission denied or not available'}`);
    }
  }

  // Backup Named Locations
  if (includeAll || policyTypes.includes('namedLocations')) {
    try {
      const namedLocations = await graphClient
        .api('/identity/conditionalAccess/namedLocations')
        .get();
      
      backupResult.policies.push({
        type: 'namedLocations',
        count: namedLocations.value?.length || 0,
        data: namedLocations.value || []
      });
    } catch (e: any) {
      errors.push(`namedLocations: ${e.message || 'Permission denied or not available'}`);
    }
  }

  // Backup Authentication Strength Policies
  if (includeAll || policyTypes.includes('authenticationStrengths')) {
    try {
      const authStrengths = await graphClient
        .api('/identity/conditionalAccess/authenticationStrength/policies')
        .get();
      
      backupResult.policies.push({
        type: 'authenticationStrengths',
        count: authStrengths.value?.length || 0,
        data: authStrengths.value || []
      });
    } catch (e: any) {
      errors.push(`authenticationStrengths: ${e.message || 'Permission denied or not available'}`);
    }
  }

  // Backup Device Compliance Policies
  if (includeAll || policyTypes.includes('deviceCompliancePolicies')) {
    try {
      const compliancePolicies = await graphClient
        .api('/deviceManagement/deviceCompliancePolicies')
        .expand('assignments')
        .get();
      
      backupResult.policies.push({
        type: 'deviceCompliancePolicies',
        count: compliancePolicies.value?.length || 0,
        data: compliancePolicies.value || []
      });
    } catch (e: any) {
      errors.push(`deviceCompliancePolicies: ${e.message || 'Permission denied or not available'}`);
    }
  }

  // Backup Device Configuration Policies
  if (includeAll || policyTypes.includes('deviceConfigurationPolicies')) {
    try {
      const configPolicies = await graphClient
        .api('/deviceManagement/deviceConfigurations')
        .expand('assignments')
        .get();
      
      backupResult.policies.push({
        type: 'deviceConfigurationPolicies',
        count: configPolicies.value?.length || 0,
        data: configPolicies.value || []
      });
    } catch (e: any) {
      errors.push(`deviceConfigurationPolicies: ${e.message || 'Permission denied or not available'}`);
    }
  }

  // Backup App Protection Policies
  if (includeAll || policyTypes.includes('appProtectionPolicies')) {
    try {
      // Get iOS policies
      const iosPolicies = await graphClient
        .api('/deviceAppManagement/iosManagedAppProtections')
        .expand('assignments')
        .get();
      
      // Get Android policies
      const androidPolicies = await graphClient
        .api('/deviceAppManagement/androidManagedAppProtections')
        .expand('assignments')
        .get();
      
      const allAppPolicies = [
        ...(iosPolicies.value || []).map((p: any) => ({ ...p, platform: 'iOS' })),
        ...(androidPolicies.value || []).map((p: any) => ({ ...p, platform: 'Android' }))
      ];
      
      backupResult.policies.push({
        type: 'appProtectionPolicies',
        count: allAppPolicies.length,
        data: allAppPolicies
      });
    } catch (e: any) {
      errors.push(`appProtectionPolicies: ${e.message || 'Permission denied or not available'}`);
    }
  }

  // Backup Sensitivity Labels (if requested)
  if (includeAll || policyTypes.includes('sensitivityLabels')) {
    try {
      const labels = await graphClient
        .api('/security/informationProtection/sensitivityLabels')
        .get();
      
      backupResult.policies.push({
        type: 'sensitivityLabels',
        count: labels.value?.length || 0,
        data: labels.value || []
      });
    } catch (e: any) {
      errors.push(`sensitivityLabels: ${e.message || 'Permission denied or not available'}`);
    }
  }

  // Format output based on requested format
  if (args.outputFormat === 'summary') {
    const summary = backupResult.policies.map(p => 
      `- **${p.type}**: ${p.count} policies`
    ).join('\n');

    let output = `# Policy Backup Summary\n\n**Timestamp**: ${backupResult.timestamp}\n**Tenant ID**: ${backupResult.tenantId || 'Unknown'}\n\n## Policies Backed Up\n${summary}`;
    
    if (errors.length > 0) {
      output += `\n\n## Errors/Warnings\n${errors.map(e => `- ${e}`).join('\n')}`;
    }

    output += `\n\n## Total Policies\n${backupResult.policies.reduce((sum, p) => sum + p.count, 0)} policies across ${backupResult.policies.length} categories`;

    return {
      content: [{
        type: 'text',
        text: output
      }]
    };
  }

  // JSON output (default)
  let output = `# Policy Backup Export\n\nBackup completed at: ${backupResult.timestamp}\n\n`;
  
  if (errors.length > 0) {
    output += `## Warnings\nSome policy types could not be exported:\n${errors.map(e => `- ${e}`).join('\n')}\n\n`;
  }

  output += `## Backup Data (JSON)\n\nCopy the JSON below to save your policy backup:\n\n\`\`\`json\n${JSON.stringify(backupResult, null, 2)}\n\`\`\``;

  return {
    content: [{
      type: 'text',
      text: output
    }]
  };
}
