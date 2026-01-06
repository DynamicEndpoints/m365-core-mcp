import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { NamedLocationsArgs, AuthenticationStrengthArgs, CrossTenantAccessArgs, IdentityProtectionArgs } from '../schemas/identity-security-schemas.js';

// Named Locations Handler
export async function handleNamedLocations(
  graphClient: Client,
  args: NamedLocationsArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let result: any;

  switch (args.action) {
    case 'list':
      result = await graphClient
        .api('/identity/conditionalAccess/namedLocations')
        .get();
      
      const locations = result.value || [];
      const formattedLocations = locations.map((loc: any) => {
        const isIP = loc['@odata.type'] === '#microsoft.graph.ipNamedLocation';
        return {
          id: loc.id,
          displayName: loc.displayName,
          type: isIP ? 'IP Named Location' : 'Country Named Location',
          createdDateTime: loc.createdDateTime,
          modifiedDateTime: loc.modifiedDateTime,
          ...(isIP ? {
            isTrusted: loc.isTrusted,
            ipRangeCount: loc.ipRanges?.length || 0
          } : {
            countriesAndRegions: loc.countriesAndRegions,
            includeUnknownCountriesAndRegions: loc.includeUnknownCountriesAndRegions
          })
        };
      });

      return {
        content: [{
          type: 'text',
          text: `# Named Locations\n\nFound ${formattedLocations.length} named locations:\n\n${JSON.stringify(formattedLocations, null, 2)}`
        }]
      };

    case 'get':
      if (!args.locationId) {
        throw new McpError(ErrorCode.InvalidParams, 'locationId is required for get action');
      }
      result = await graphClient
        .api(`/identity/conditionalAccess/namedLocations/${args.locationId}`)
        .get();
      break;

    case 'create':
      if (!args.displayName || !args.locationType) {
        throw new McpError(ErrorCode.InvalidParams, 'displayName and locationType are required for create action');
      }

      if (args.locationType === 'ipNamedLocation') {
        if (!args.ipRanges || args.ipRanges.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'ipRanges is required for IP named location');
        }
        const ipLocationPayload = {
          '@odata.type': '#microsoft.graph.ipNamedLocation',
          displayName: args.displayName,
          isTrusted: args.isTrusted ?? false,
          ipRanges: args.ipRanges.map(range => ({
            '@odata.type': '#microsoft.graph.iPv4CidrRange',
            cidrAddress: range.cidrAddress
          }))
        };
        result = await graphClient
          .api('/identity/conditionalAccess/namedLocations')
          .post(ipLocationPayload);
      } else {
        if (!args.countriesAndRegions || args.countriesAndRegions.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'countriesAndRegions is required for country named location');
        }
        const countryLocationPayload = {
          '@odata.type': '#microsoft.graph.countryNamedLocation',
          displayName: args.displayName,
          countriesAndRegions: args.countriesAndRegions,
          includeUnknownCountriesAndRegions: args.includeUnknownCountriesAndRegions ?? false
        };
        result = await graphClient
          .api('/identity/conditionalAccess/namedLocations')
          .post(countryLocationPayload);
      }
      break;

    case 'update':
      if (!args.locationId) {
        throw new McpError(ErrorCode.InvalidParams, 'locationId is required for update action');
      }
      
      const updatePayload: any = {};
      if (args.displayName) updatePayload.displayName = args.displayName;
      if (args.isTrusted !== undefined) updatePayload.isTrusted = args.isTrusted;
      if (args.ipRanges) {
        updatePayload.ipRanges = args.ipRanges.map(range => ({
          '@odata.type': '#microsoft.graph.iPv4CidrRange',
          cidrAddress: range.cidrAddress
        }));
      }
      if (args.countriesAndRegions) updatePayload.countriesAndRegions = args.countriesAndRegions;
      if (args.includeUnknownCountriesAndRegions !== undefined) {
        updatePayload.includeUnknownCountriesAndRegions = args.includeUnknownCountriesAndRegions;
      }

      result = await graphClient
        .api(`/identity/conditionalAccess/namedLocations/${args.locationId}`)
        .patch(updatePayload);
      break;

    case 'delete':
      if (!args.locationId) {
        throw new McpError(ErrorCode.InvalidParams, 'locationId is required for delete action');
      }
      await graphClient
        .api(`/identity/conditionalAccess/namedLocations/${args.locationId}`)
        .delete();
      result = { message: `Named location ${args.locationId} deleted successfully` };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Unknown action: ${args.action}`);
  }

  return {
    content: [{
      type: 'text',
      text: `Named Location ${args.action} operation completed:\n\n${JSON.stringify(result, null, 2)}`
    }]
  };
}

// Authentication Strength Policy Handler
export async function handleAuthenticationStrength(
  graphClient: Client,
  args: AuthenticationStrengthArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let result: any;

  switch (args.action) {
    case 'list':
      result = await graphClient
        .api('/identity/conditionalAccess/authenticationStrength/policies')
        .get();
      
      const policies = result.value || [];
      let filteredPolicies = policies;
      
      if (args.policyType && args.policyType !== 'all') {
        filteredPolicies = policies.filter((p: any) => 
          args.policyType === 'builtIn' ? p.policyType === 'builtIn' : p.policyType === 'custom'
        );
      }

      const formattedPolicies = filteredPolicies.map((p: any) => ({
        id: p.id,
        displayName: p.displayName,
        description: p.description,
        policyType: p.policyType,
        requirementsSatisfied: p.requirementsSatisfied,
        allowedCombinations: p.allowedCombinations,
        createdDateTime: p.createdDateTime,
        modifiedDateTime: p.modifiedDateTime
      }));

      return {
        content: [{
          type: 'text',
          text: `# Authentication Strength Policies\n\nFound ${formattedPolicies.length} policies:\n\n${JSON.stringify(formattedPolicies, null, 2)}`
        }]
      };

    case 'get':
      if (!args.policyId) {
        throw new McpError(ErrorCode.InvalidParams, 'policyId is required for get action');
      }
      result = await graphClient
        .api(`/identity/conditionalAccess/authenticationStrength/policies/${args.policyId}`)
        .get();
      break;

    case 'listCombinations':
      // List all possible authentication method combinations
      result = await graphClient
        .api('/identity/conditionalAccess/authenticationStrength/authenticationMethodModes')
        .get();
      
      return {
        content: [{
          type: 'text',
          text: `# Available Authentication Method Modes\n\nThese are the authentication method modes that can be combined in authentication strength policies:\n\n${JSON.stringify(result.value || [], null, 2)}`
        }]
      };

    case 'listMethods':
      // List authentication methods configuration
      result = await graphClient
        .api('/policies/authenticationMethodsPolicy/authenticationMethodConfigurations')
        .get();
      
      const methods = (result.value || []).map((m: any) => ({
        id: m.id,
        state: m.state,
        '@odata.type': m['@odata.type']
      }));

      return {
        content: [{
          type: 'text',
          text: `# Configured Authentication Methods\n\n${JSON.stringify(methods, null, 2)}`
        }]
      };

    default:
      throw new McpError(ErrorCode.InvalidParams, `Unknown action: ${args.action}`);
  }

  return {
    content: [{
      type: 'text',
      text: `Authentication Strength ${args.action} operation completed:\n\n${JSON.stringify(result, null, 2)}`
    }]
  };
}

// Cross-Tenant Access Settings Handler
export async function handleCrossTenantAccess(
  graphClient: Client,
  args: CrossTenantAccessArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let result: any;

  switch (args.action) {
    case 'getDefault':
      result = await graphClient
        .api('/policies/crossTenantAccessPolicy/default')
        .get();
      break;

    case 'listPartners':
      result = await graphClient
        .api('/policies/crossTenantAccessPolicy/partners')
        .get();
      
      const partners = result.value || [];
      return {
        content: [{
          type: 'text',
          text: `# Cross-Tenant Access Partners\n\nFound ${partners.length} partner configurations:\n\n${JSON.stringify(partners, null, 2)}`
        }]
      };

    case 'getPartner':
      if (!args.tenantId) {
        throw new McpError(ErrorCode.InvalidParams, 'tenantId is required for getPartner action');
      }
      result = await graphClient
        .api(`/policies/crossTenantAccessPolicy/partners/${args.tenantId}`)
        .get();
      break;

    case 'updateDefault':
      const updatePayload: any = {};
      if (args.inboundTrust) {
        updatePayload.inboundTrust = args.inboundTrust;
      }
      if (args.b2bCollaborationInbound) {
        updatePayload.b2bCollaborationInbound = args.b2bCollaborationInbound;
      }
      
      result = await graphClient
        .api('/policies/crossTenantAccessPolicy/default')
        .patch(updatePayload);
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Unknown action: ${args.action}`);
  }

  return {
    content: [{
      type: 'text',
      text: `Cross-Tenant Access ${args.action} operation completed:\n\n${JSON.stringify(result, null, 2)}`
    }]
  };
}

// Identity Protection Handler
export async function handleIdentityProtection(
  graphClient: Client,
  args: IdentityProtectionArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let result: any;
  let apiPath: string;

  switch (args.action) {
    case 'listRiskDetections':
      apiPath = '/identityProtection/riskDetections';
      let query = graphClient.api(apiPath);
      
      const filters: string[] = [];
      if (args.riskLevel) {
        filters.push(`riskLevel eq '${args.riskLevel}'`);
      }
      if (args.riskState) {
        filters.push(`riskState eq '${args.riskState}'`);
      }
      if (filters.length > 0) {
        query = query.filter(filters.join(' and '));
      }
      if (args.top) {
        query = query.top(args.top);
      }
      
      result = await query.get();
      
      const detections = (result.value || []).map((d: any) => ({
        id: d.id,
        riskType: d.riskType,
        riskLevel: d.riskLevel,
        riskState: d.riskState,
        riskDetail: d.riskDetail,
        userDisplayName: d.userDisplayName,
        userPrincipalName: d.userPrincipalName,
        detectedDateTime: d.detectedDateTime,
        ipAddress: d.ipAddress,
        location: d.location
      }));

      return {
        content: [{
          type: 'text',
          text: `# Risk Detections\n\nFound ${detections.length} risk detections:\n\n${JSON.stringify(detections, null, 2)}`
        }]
      };

    case 'getRiskDetection':
      if (!args.riskDetectionId) {
        throw new McpError(ErrorCode.InvalidParams, 'riskDetectionId is required for getRiskDetection action');
      }
      result = await graphClient
        .api(`/identityProtection/riskDetections/${args.riskDetectionId}`)
        .get();
      break;

    case 'listRiskyUsers':
      apiPath = '/identityProtection/riskyUsers';
      let riskyQuery = graphClient.api(apiPath);
      
      const riskyFilters: string[] = [];
      if (args.riskLevel) {
        riskyFilters.push(`riskLevel eq '${args.riskLevel}'`);
      }
      if (args.riskState) {
        riskyFilters.push(`riskState eq '${args.riskState}'`);
      }
      if (riskyFilters.length > 0) {
        riskyQuery = riskyQuery.filter(riskyFilters.join(' and '));
      }
      if (args.top) {
        riskyQuery = riskyQuery.top(args.top);
      }
      
      result = await riskyQuery.get();
      
      const riskyUsers = (result.value || []).map((u: any) => ({
        id: u.id,
        userDisplayName: u.userDisplayName,
        userPrincipalName: u.userPrincipalName,
        riskLevel: u.riskLevel,
        riskState: u.riskState,
        riskDetail: u.riskDetail,
        riskLastUpdatedDateTime: u.riskLastUpdatedDateTime
      }));

      return {
        content: [{
          type: 'text',
          text: `# Risky Users\n\nFound ${riskyUsers.length} risky users:\n\n${JSON.stringify(riskyUsers, null, 2)}`
        }]
      };

    case 'getRiskyUser':
      if (!args.userId) {
        throw new McpError(ErrorCode.InvalidParams, 'userId is required for getRiskyUser action');
      }
      result = await graphClient
        .api(`/identityProtection/riskyUsers/${args.userId}`)
        .get();
      break;

    case 'dismissRiskyUser':
      if (!args.userId) {
        throw new McpError(ErrorCode.InvalidParams, 'userId is required for dismissRiskyUser action');
      }
      await graphClient
        .api('/identityProtection/riskyUsers/dismiss')
        .post({ userIds: [args.userId] });
      result = { message: `Risky user ${args.userId} dismissed successfully` };
      break;

    case 'confirmRiskyUserCompromised':
      if (!args.userId) {
        throw new McpError(ErrorCode.InvalidParams, 'userId is required for confirmRiskyUserCompromised action');
      }
      await graphClient
        .api('/identityProtection/riskyUsers/confirmCompromised')
        .post({ userIds: [args.userId] });
      result = { message: `Risky user ${args.userId} confirmed as compromised` };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Unknown action: ${args.action}`);
  }

  return {
    content: [{
      type: 'text',
      text: `Identity Protection ${args.action} operation completed:\n\n${JSON.stringify(result, null, 2)}`
    }]
  };
}
