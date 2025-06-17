import { z } from 'zod';

export const createIntunePolicySchema = z.object({
  platform: z.enum(['windows', 'macos']).describe('The platform for the policy.'),
  policyType: z.enum(['Configuration', 'Compliance', 'Security', 'Update', 'AppProtection', 'EndpointSecurity']).describe('The type of policy to create.'),
  displayName: z.string().describe('The name of the policy.'),
  description: z.string().optional().describe('The description of the policy.'),
  settings: z.any().describe('The policy settings. The structure of this object depends on the platform and policyType.'),
  assignments: z.array(z.any()).optional().describe('The assignments for the policy.')
});

export const intuneTools = [
    {
        name: 'createIntunePolicy',
        description: 'Creates a new Intune policy for either Windows or macOS.',
        inputSchema: createIntunePolicySchema
    }
];
