import { z } from 'zod';

// Policy Backup Schema
export const policyBackupArgsSchema = z.object({
  action: z.enum(['backup', 'list', 'restore']).describe('Action to perform - backup exports policies, list shows available types'),
  policyTypes: z.array(z.enum([
    'conditionalAccess',
    'namedLocations',
    'authenticationStrengths',
    'deviceCompliancePolicies',
    'deviceConfigurationPolicies',
    'appProtectionPolicies',
    'dlpPolicies',
    'retentionPolicies',
    'sensitivityLabels',
    'all'
  ])).optional().describe('Types of policies to backup. Use "all" to backup all types'),
  outputFormat: z.enum(['json', 'summary']).optional().describe('Output format - json for full backup, summary for overview'),
  includeMetadata: z.boolean().optional().describe('Include metadata like tenant ID, export timestamp'),
});

// Named Locations Management Schema
export const namedLocationsArgsSchema = z.object({
  action: z.enum(['list', 'get', 'create', 'update', 'delete']).describe('Action to perform on named locations'),
  locationId: z.string().optional().describe('Named location ID for specific operations'),
  displayName: z.string().optional().describe('Display name for the location'),
  locationType: z.enum(['ipNamedLocation', 'countryNamedLocation']).optional().describe('Type of named location'),
  isTrusted: z.boolean().optional().describe('Whether to mark IP ranges as trusted'),
  ipRanges: z.array(z.object({
    cidrAddress: z.string().describe('CIDR notation for IP range (e.g., 192.168.1.0/24)')
  })).optional().describe('IP ranges for IP-based named location'),
  countriesAndRegions: z.array(z.string()).optional().describe('ISO 3166-1 alpha-2 country codes (e.g., ["US", "CA"])'),
  includeUnknownCountriesAndRegions: z.boolean().optional().describe('Include unknown countries/regions'),
});

// Authentication Strength Policy Schema
export const authenticationStrengthArgsSchema = z.object({
  action: z.enum(['list', 'get', 'listCombinations', 'listMethods']).describe('Action to perform on authentication strength policies'),
  policyId: z.string().optional().describe('Authentication strength policy ID for specific operations'),
  policyType: z.enum(['builtIn', 'custom', 'all']).optional().describe('Type of policies to list'),
});

// Cross-Tenant Access Settings Schema
export const crossTenantAccessArgsSchema = z.object({
  action: z.enum(['getDefault', 'listPartners', 'getPartner', 'updateDefault']).describe('Action to perform on cross-tenant access settings'),
  tenantId: z.string().optional().describe('Partner tenant ID for specific operations'),
  inboundTrust: z.object({
    isMfaAccepted: z.boolean().optional().describe('Accept MFA from partner tenant'),
    isCompliantDeviceAccepted: z.boolean().optional().describe('Accept compliant device status from partner'),
    isHybridAzureADJoinedDeviceAccepted: z.boolean().optional().describe('Accept hybrid Azure AD joined device status'),
  }).optional().describe('Inbound trust settings'),
  b2bCollaborationInbound: z.object({
    applications: z.object({
      accessType: z.enum(['allowed', 'blocked']).describe('Access type for applications'),
      targets: z.array(z.object({
        target: z.string().describe('Application ID or "All"'),
        targetType: z.enum(['application', 'group']).describe('Type of target')
      })).optional().describe('Application targets')
    }).optional().describe('Application access settings'),
    usersAndGroups: z.object({
      accessType: z.enum(['allowed', 'blocked']).describe('Access type for users/groups'),
      targets: z.array(z.object({
        target: z.string().describe('User/Group ID or "All"'),
        targetType: z.enum(['user', 'group']).describe('Type of target')
      })).optional().describe('User/group targets')
    }).optional().describe('User and group access settings'),
  }).optional().describe('B2B collaboration inbound settings'),
});

// Identity Protection Risk Policies Schema
export const identityProtectionArgsSchema = z.object({
  action: z.enum(['listRiskDetections', 'getRiskDetection', 'listRiskyUsers', 'getRiskyUser', 'dismissRiskyUser', 'confirmRiskyUserCompromised']).describe('Action to perform'),
  riskDetectionId: z.string().optional().describe('Risk detection ID'),
  userId: z.string().optional().describe('User ID for risky user operations'),
  riskLevel: z.enum(['low', 'medium', 'high', 'hidden', 'none']).optional().describe('Filter by risk level'),
  riskState: z.enum(['atRisk', 'confirmedCompromised', 'remediated', 'dismissed', 'unknownFutureValue']).optional().describe('Filter by risk state'),
  top: z.number().optional().describe('Number of results to return'),
});

export type PolicyBackupArgs = z.infer<typeof policyBackupArgsSchema>;
export type NamedLocationsArgs = z.infer<typeof namedLocationsArgsSchema>;
export type AuthenticationStrengthArgs = z.infer<typeof authenticationStrengthArgsSchema>;
export type CrossTenantAccessArgs = z.infer<typeof crossTenantAccessArgsSchema>;
export type IdentityProtectionArgs = z.infer<typeof identityProtectionArgsSchema>;
