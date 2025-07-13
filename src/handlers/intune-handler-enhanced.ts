import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { handleIntuneMacOSPolicies } from './intune-macos-handler.js';
import { handleIntuneWindowsPolicies } from './intune-windows-handler.js';
import { 
  validatePolicySettings, 
  applyPolicyDefaults, 
  generatePolicyExample,
  validateAssignments,
  POLICY_TEMPLATES
} from '../validators/intune-policy-validator.js';

export async function handleCreateIntunePolicyEnhanced(
  graphClient: Client, 
  args: any
): Promise<any> {
  const { platform, policyType, displayName, description, settings, assignments, useTemplate } = args;

  // Basic parameter validation
  if (!platform || !policyType || !displayName) {
    throw new McpError(
      ErrorCode.InvalidParams, 
      'Missing required parameters: platform, policyType, and displayName are required.'
    );
  }

  // Validate platform
  if (!['macos', 'windows'].includes(platform)) {
    throw new McpError(
      ErrorCode.InvalidParams, 
      `Invalid platform: ${platform}. Must be 'macos' or 'windows'.`
    );
  }

  // Validate policy type
  const validPolicyTypes = ['Configuration', 'Compliance', 'Security', 'Update', 'AppProtection', 'EndpointSecurity'];
  if (!validPolicyTypes.includes(policyType)) {
    throw new McpError(
      ErrorCode.InvalidParams, 
      `Invalid policyType: ${policyType}. Must be one of: ${validPolicyTypes.join(', ')}`
    );
  }

  let finalSettings = settings || {};

  // Apply template if requested
  if (useTemplate) {
    const platformTemplates = POLICY_TEMPLATES[platform as 'macos' | 'windows'];
    if (platformTemplates && useTemplate in platformTemplates) {
      const template = platformTemplates[useTemplate as keyof typeof platformTemplates];
      if (template && 'settings' in template) {
        finalSettings = template.settings;
        console.debug(`Applied template: ${useTemplate} for ${platform} ${policyType}`);
      }
    } else {
      console.warn(`Template '${useTemplate}' not found for ${platform}`);
    }
  }

  // Apply defaults
  finalSettings = applyPolicyDefaults(platform, policyType, finalSettings);

  // Validate settings
  const validationResult = validatePolicySettings(platform, policyType, finalSettings);
  
  // Log warnings
  if (validationResult.warnings.length > 0) {
    console.warn('Policy validation warnings:', validationResult.warnings);
  }

  // Check for validation errors
  if (!validationResult.isValid) {
    const errorMessage = `Policy validation failed:\n${validationResult.errors.join('\n')}`;
    const example = generatePolicyExample(platform, policyType);
    throw new McpError(
      ErrorCode.InvalidParams,
      `${errorMessage}\n\nExample of valid settings:\n${example}`
    );
  }

  // Validate assignments if provided
  let assignmentValidation = { warnings: [] as string[], errors: [] as string[], isValid: true };
  if (assignments) {
    assignmentValidation = validateAssignments(assignments);
    
    if (assignmentValidation.warnings.length > 0) {
      console.warn('Assignment validation warnings:', assignmentValidation.warnings);
    }

    if (!assignmentValidation.isValid) {
      throw new McpError(
        ErrorCode.InvalidParams,
        `Assignment validation failed:\n${assignmentValidation.errors.join('\n')}`
      );
    }
  }

  // Prepare the policy arguments
  const policyArgs = {
    action: 'create' as const,
    policyType: policyType,
    name: displayName,
    description: description,
    settings: finalSettings,
    assignments: assignments
  };

  try {
    // Route to appropriate handler based on platform
    let result;
    if (platform === 'macos') {
      result = await handleIntuneMacOSPolicies(graphClient, policyArgs as any);
    } else if (platform === 'windows') {
      result = await handleIntuneWindowsPolicies(graphClient, policyArgs as any);
    }

    // Add validation info to result
    const enhancedResult = {
      ...result,
      validationInfo: {
        platform,
        policyType,
        settingsValidated: true,
        defaultsApplied: true,
        warnings: [...validationResult.warnings, ...assignmentValidation.warnings]
      }
    };

    return enhancedResult;
  } catch (error) {
    // Enhance error message with context
    if (error instanceof Error) {
      const enhancedError = new McpError(
        ErrorCode.InternalError,
        `Failed to create ${platform} ${policyType} policy: ${error.message}\n\n` +
        `Policy Name: ${displayName}\n` +
        `Platform: ${platform}\n` +
        `Type: ${policyType}`
      );
      throw enhancedError;
    }
    throw error;
  }
}

// Export the enhanced handler as the main handler
export async function handleCreateIntunePolicy(
  graphClient: Client, 
  args: any
): Promise<any> {
  return handleCreateIntunePolicyEnhanced(graphClient, args);
}

// Helper function to list available templates
export function listPolicyTemplates(platform?: string): any {
  if (platform && (platform === 'macos' || platform === 'windows')) {
    return POLICY_TEMPLATES[platform] || {};
  }
  return POLICY_TEMPLATES;
}

// Helper function to get policy creation help
export function getPolicyCreationHelp(platform: string, policyType: string): string {
  const example = generatePolicyExample(platform, policyType);
  const templates = (platform === 'macos' || platform === 'windows') ? POLICY_TEMPLATES[platform] : {};
  const templateNames = templates ? Object.keys(templates) : [];

  return `
# Creating ${platform} ${policyType} Policy

## Required Parameters:
- platform: '${platform}'
- policyType: '${policyType}'
- displayName: string (required)
- description: string (optional)
- settings: object (structure depends on policy type)
- assignments: array (optional)

## Available Templates for ${platform}:
${templateNames.length > 0 ? templateNames.join(', ') : 'No templates available'}

## Example:
${example}

## Using a Template:
Add "useTemplate": "templateName" to use a pre-configured template.
`;
}
