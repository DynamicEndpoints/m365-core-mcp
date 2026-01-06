import { Client } from '@microsoft/microsoft-graph-client';
import { PromptHandler } from './prompts.js';

/**
 * Identity and Security Prompts
 * Comprehensive prompts for Conditional Access, Identity Protection, and Policy Management
 */

export const identitySecurityPrompts: PromptHandler[] = [
  // ===== POLICY BACKUP STRATEGY PROMPT =====
  {
    name: 'policy_backup_strategy',
    description: 'Comprehensive guide for backing up and managing Microsoft 365 policy configurations for disaster recovery',
    arguments: [
      {
        name: 'backup_scope',
        description: 'Scope of backup: full, security, compliance, or intune (default: full)',
        required: false
      },
      {
        name: 'output_format',
        description: 'Output format: json, summary, or both (default: both)',
        required: false
      }
    ],
    handler: async (graphClient, args) => {
      const scope = args?.backup_scope || 'full';
      const format = args?.output_format || 'both';

      return `# Microsoft 365 Policy Backup Strategy

## Objective
Create a comprehensive backup of Microsoft 365 policy configurations for disaster recovery, migration, and documentation purposes.

## Backup Scope: ${scope.toUpperCase()}

### 1. Identity & Access Policies

#### Conditional Access Policies
Critical for zero-trust security. Must backup:
- All Conditional Access policies with complete conditions
- Grant and session controls
- Policy state (enabled, disabled, report-only)

**Backup Command:**
\`\`\`json
{
  "action": "backup",
  "policyTypes": ["conditionalAccess"],
  "outputFormat": "${format}",
  "includeMetadata": true
}
\`\`\`
**Tool:** \`backup_policies\`

#### Named Locations
IP ranges and country-based locations used by CA policies.

**Backup Command:**
\`\`\`json
{
  "action": "backup", 
  "policyTypes": ["namedLocations"],
  "outputFormat": "json",
  "includeMetadata": true
}
\`\`\`

#### Authentication Strength Policies
MFA requirements and authentication method combinations.

**View Authentication Strengths:**
\`\`\`json
{
  "action": "list",
  "policyType": "all"
}
\`\`\`
**Tool:** \`manage_authentication_strengths\`

### 2. Device Management Policies (Intune)

${scope === 'full' || scope === 'intune' ? `
#### Device Compliance Policies
- Windows compliance requirements
- macOS compliance requirements
- iOS/Android compliance settings

**Backup Command:**
\`\`\`json
{
  "action": "backup",
  "policyTypes": ["deviceCompliancePolicies"],
  "outputFormat": "json",
  "includeMetadata": true
}
\`\`\`

#### Device Configuration Policies
- Windows configuration profiles
- macOS configuration profiles
- Security baselines

**Backup Command:**
\`\`\`json
{
  "action": "backup",
  "policyTypes": ["deviceConfigurationPolicies"],
  "outputFormat": "json",
  "includeMetadata": true
}
\`\`\`

#### App Protection Policies
- iOS app protection (MAM)
- Android app protection (MAM)

**Backup Command:**
\`\`\`json
{
  "action": "backup",
  "policyTypes": ["appProtectionPolicies"],
  "outputFormat": "json",
  "includeMetadata": true
}
\`\`\`
` : '*Intune policies not included in this scope*'}

### 3. Data Protection Policies

${scope === 'full' || scope === 'compliance' ? `
#### Sensitivity Labels
- Label definitions and scope
- Auto-labeling rules
- Visual markings and protection settings

**Backup Command:**
\`\`\`json
{
  "action": "backup",
  "policyTypes": ["sensitivityLabels"],
  "outputFormat": "json",
  "includeMetadata": true
}
\`\`\`

#### DLP Policies
- Data loss prevention rules
- Sensitive information types
- Policy locations and exceptions
` : '*Compliance policies not included in this scope*'}

## Complete Backup Strategy

### Full Environment Backup
\`\`\`json
{
  "action": "backup",
  "policyTypes": ["all"],
  "outputFormat": "json",
  "includeMetadata": true
}
\`\`\`

### Available Policy Types
Run \`backup_policies\` with action: "list" to see all available policy types.

## Backup Best Practices

### 1. Frequency
- **Daily**: Conditional Access policies (high-change environment)
- **Weekly**: Full policy backup
- **Before changes**: Always backup before major policy modifications
- **Post-implementation**: Backup after successful rollouts

### 2. Storage Recommendations
- Store backups in version-controlled repository (Git)
- Maintain at least 30 days of backup history
- Store in multiple locations (Azure Blob, SharePoint, local)
- Encrypt backups containing sensitive policy details

### 3. Naming Convention
\`\`\`
{tenant}-{date}-{scope}-backup.json
Example: contoso-2026-01-06-full-backup.json
\`\`\`

### 4. Validation
After backup, validate:
- JSON is well-formed
- All expected policy types are included
- Policy counts match expected numbers
- Metadata includes correct tenant information

## Recovery Procedures

### Conditional Access Recovery
1. Review backup JSON for policy details
2. Use \`manage_conditional_access_policies\` with action: "create"
3. Recreate policies one-by-one, starting with most critical
4. Test in report-only mode before enabling

### Named Locations Recovery
1. Use \`manage_named_locations\` with action: "create"
2. Recreate IP ranges and country locations
3. Update CA policies to reference new location IDs

### Important Notes
- Policy IDs will change on recreation
- Some policies may have dependencies (e.g., CA policies referencing named locations)
- Recreate in dependency order: Named Locations → Auth Strengths → CA Policies

## Monitoring & Alerts

Set up monitoring for policy changes:
- \`search_audit_log\` with service: "Conditional Access"
- Alert on policy modifications outside change windows
- Track backup job success/failure

---

**Backup Scope**: ${scope}
**Output Format**: ${format}
**Generated**: ${new Date().toISOString()}`;
    }
  },

  // ===== CONDITIONAL ACCESS DESIGN PROMPT =====
  {
    name: 'conditional_access_design',
    description: 'Design and implement Conditional Access policies following zero-trust principles and Microsoft best practices',
    arguments: [
      {
        name: 'scenario',
        description: 'Deployment scenario: baseline, remote_work, byod, contractors, privileged_access (default: baseline)',
        required: false
      },
      {
        name: 'strictness',
        description: 'Security level: standard, strict, maximum (default: standard)',
        required: false
      }
    ],
    handler: async (graphClient, args) => {
      const scenario = args?.scenario || 'baseline';
      const strictness = args?.strictness || 'standard';

      return `# Conditional Access Policy Design Guide

## Objective
Design and implement Conditional Access policies for: **${scenario.replace('_', ' ').toUpperCase()}**
Security Level: **${strictness.toUpperCase()}**

## Zero Trust Principles

1. **Verify Explicitly** - Always authenticate and authorize based on all available data points
2. **Least Privilege Access** - Limit user access with just-in-time and just-enough-access
3. **Assume Breach** - Minimize blast radius and segment access

## Pre-Implementation Checklist

### 1. Inventory Current State
\`\`\`json
// List existing Conditional Access policies
{
  "action": "list"
}
\`\`\`
**Tool:** \`manage_conditional_access_policies\`

### 2. Review Named Locations
\`\`\`json
{
  "action": "list"
}
\`\`\`
**Tool:** \`manage_named_locations\`

### 3. Check Authentication Methods
\`\`\`json
{
  "action": "listMethods"
}
\`\`\`
**Tool:** \`manage_authentication_strengths\`

## Policy Templates by Scenario

${scenario === 'baseline' || scenario === 'all' ? `
### Baseline Security Policies

#### Policy 1: Require MFA for All Users
\`\`\`json
{
  "action": "create",
  "displayName": "Require MFA for All Users",
  "state": "enabledForReportingButNotEnforced",
  "conditions": {
    "users": {
      "includeUsers": ["All"],
      "excludeUsers": ["<break-glass-account-id>"]
    },
    "applications": {
      "includeApplications": ["All"]
    }
  },
  "grantControls": {
    "operator": "OR",
    "builtInControls": ["mfa"]
  }
}
\`\`\`

#### Policy 2: Block Legacy Authentication
\`\`\`json
{
  "action": "create",
  "displayName": "Block Legacy Authentication",
  "state": "enabledForReportingButNotEnforced",
  "conditions": {
    "users": {
      "includeUsers": ["All"]
    },
    "applications": {
      "includeApplications": ["All"]
    },
    "clientAppTypes": ["exchangeActiveSync", "other"]
  },
  "grantControls": {
    "operator": "OR",
    "builtInControls": ["block"]
  }
}
\`\`\`

#### Policy 3: Require Compliant Device for Office 365
\`\`\`json
{
  "action": "create",
  "displayName": "Require Compliant Device for Office 365",
  "state": "enabledForReportingButNotEnforced",
  "conditions": {
    "users": {
      "includeUsers": ["All"]
    },
    "applications": {
      "includeApplications": ["Office365"]
    }
  },
  "grantControls": {
    "operator": "AND",
    "builtInControls": ["mfa", "compliantDevice"]
  }
}
\`\`\`
` : ''}

${scenario === 'remote_work' ? `
### Remote Work Security Policies

#### Policy 1: Require MFA Outside Corporate Network
First, create a named location for corporate IPs:
\`\`\`json
{
  "action": "create",
  "displayName": "Corporate Network",
  "locationType": "ipNamedLocation",
  "isTrusted": true,
  "ipRanges": [
    {"cidrAddress": "10.0.0.0/8"},
    {"cidrAddress": "192.168.1.0/24"},
    {"cidrAddress": "<your-public-ip>/32"}
  ]
}
\`\`\`
**Tool:** \`manage_named_locations\`

Then create the CA policy:
\`\`\`json
{
  "action": "create",
  "displayName": "Require MFA Outside Corporate Network",
  "state": "enabledForReportingButNotEnforced",
  "conditions": {
    "users": {
      "includeUsers": ["All"]
    },
    "applications": {
      "includeApplications": ["All"]
    },
    "locations": {
      "includeLocations": ["All"],
      "excludeLocations": ["<corporate-network-id>"]
    }
  },
  "grantControls": {
    "operator": "OR",
    "builtInControls": ["mfa"]
  }
}
\`\`\`

#### Policy 2: Block Access from High-Risk Countries
\`\`\`json
{
  "action": "create",
  "displayName": "Blocked Countries",
  "locationType": "countryNamedLocation",
  "countriesAndRegions": ["RU", "CN", "KP", "IR"],
  "includeUnknownCountriesAndRegions": false
}
\`\`\`
**Tool:** \`manage_named_locations\`
` : ''}

${scenario === 'privileged_access' ? `
### Privileged Access Policies

#### Policy 1: Require Phishing-Resistant MFA for Admins
\`\`\`json
{
  "action": "create",
  "displayName": "Require Phishing-Resistant MFA for Admins",
  "state": "enabledForReportingButNotEnforced",
  "conditions": {
    "users": {
      "includeRoles": [
        "62e90394-69f5-4237-9190-012177145e10",
        "194ae4cb-b126-40b2-bd5b-6091b380977d",
        "f28a1f50-f6e7-4571-818b-6a12f2af6b6c"
      ]
    },
    "applications": {
      "includeApplications": ["All"]
    }
  },
  "grantControls": {
    "operator": "OR",
    "builtInControls": ["mfa"]
  }
}
\`\`\`
Note: Role IDs are for Global Admin, Security Admin, Privileged Role Admin

#### Policy 2: Block Admin Access from Non-Trusted Locations
\`\`\`json
{
  "action": "create",
  "displayName": "Block Admin Access from Non-Trusted Locations",
  "state": "enabledForReportingButNotEnforced",
  "conditions": {
    "users": {
      "includeRoles": ["62e90394-69f5-4237-9190-012177145e10"]
    },
    "applications": {
      "includeApplications": ["All"]
    },
    "locations": {
      "includeLocations": ["All"],
      "excludeLocations": ["AllTrusted"]
    }
  },
  "grantControls": {
    "operator": "OR",
    "builtInControls": ["block"]
  }
}
\`\`\`

#### Policy 3: Require Compliant Device for Admin Portals
\`\`\`json
{
  "action": "create",
  "displayName": "Require Compliant Device for Admin Portals",
  "state": "enabledForReportingButNotEnforced",
  "conditions": {
    "users": {
      "includeRoles": ["62e90394-69f5-4237-9190-012177145e10"]
    },
    "applications": {
      "includeApplications": [
        "797f4846-ba00-4fd7-ba43-dac1f8f63013",
        "0000000c-0000-0000-c000-000000000000"
      ]
    }
  },
  "grantControls": {
    "operator": "AND",
    "builtInControls": ["mfa", "compliantDevice"]
  }
}
\`\`\`
` : ''}

## Implementation Workflow

### Phase 1: Report-Only Mode (1-2 weeks)
1. Create policies in "enabledForReportingButNotEnforced" state
2. Monitor sign-in logs for impact
3. Review What-If analysis results
4. Identify false positives and adjust conditions

### Phase 2: Gradual Rollout (2-4 weeks)
1. Enable for pilot group first
2. Collect feedback and adjust
3. Expand to larger groups progressively
4. Document exceptions and reasons

### Phase 3: Full Enforcement
1. Switch to "enabled" state
2. Monitor for authentication failures
3. Have rollback plan ready
4. Communicate to users

## Monitoring & Troubleshooting

### Check Policy Impact
\`\`\`json
{
  "action": "listRiskDetections",
  "top": 50
}
\`\`\`
**Tool:** \`manage_identity_protection\`

### Review Sign-In Logs
\`\`\`json
{
  "endpoint": "/auditLogs/signIns",
  "method": "GET",
  "queryParams": {
    "$filter": "conditionalAccessStatus eq 'failure'"
  }
}
\`\`\`
**Tool:** \`call_microsoft_api\`

## Break-Glass Account Requirements

Always maintain emergency access accounts:
1. Cloud-only accounts (not synced from on-premises)
2. Excluded from ALL Conditional Access policies
3. Strong, unique passwords stored securely
4. MFA configured with hardware tokens
5. Monitor for any usage (should be rare)

---

**Scenario**: ${scenario}
**Security Level**: ${strictness}
**Design Date**: ${new Date().toISOString().split('T')[0]}`;
    }
  },

  // ===== IDENTITY PROTECTION RESPONSE PROMPT =====
  {
    name: 'identity_protection_response',
    description: 'Guide for investigating and responding to identity-based security threats and risky users',
    arguments: [
      {
        name: 'incident_type',
        description: 'Type of incident: risky_user, risky_signin, compromised_account, or investigation (default: investigation)',
        required: false
      },
      {
        name: 'risk_level',
        description: 'Risk level to focus on: low, medium, high (default: high)',
        required: false
      }
    ],
    handler: async (graphClient, args) => {
      const incidentType = args?.incident_type || 'investigation';
      const riskLevel = args?.risk_level || 'high';

      return `# Identity Protection Incident Response Guide

## Objective
Investigate and respond to identity-based security incidents.
**Incident Type**: ${incidentType.replace('_', ' ').toUpperCase()}
**Risk Level Focus**: ${riskLevel.toUpperCase()}

## Initial Assessment

### 1. Gather Risk Detections
\`\`\`json
{
  "action": "listRiskDetections",
  "riskLevel": "${riskLevel}",
  "top": 100
}
\`\`\`
**Tool:** \`manage_identity_protection\`

### 2. Identify Risky Users
\`\`\`json
{
  "action": "listRiskyUsers",
  "riskLevel": "${riskLevel}",
  "riskState": "atRisk"
}
\`\`\`
**Tool:** \`manage_identity_protection\`

### 3. Review Recent Authentications
\`\`\`json
{
  "endpoint": "/auditLogs/signIns",
  "method": "GET",
  "queryParams": {
    "$filter": "riskLevelDuringSignIn eq '${riskLevel}'",
    "$top": "50",
    "$orderby": "createdDateTime desc"
  }
}
\`\`\`
**Tool:** \`call_microsoft_api\`

## Risk Detection Types

### User Risk Detections
| Detection | Description | Severity |
|-----------|-------------|----------|
| Leaked Credentials | User credentials found in dark web breach | High |
| Anomalous User Activity | Unusual behavior patterns | Medium |
| Azure AD Threat Intelligence | User flagged by Microsoft threat intel | High |
| Malware-Linked IP Address | Sign-in from malware-infected IP | Medium |

### Sign-In Risk Detections
| Detection | Description | Severity |
|-----------|-------------|----------|
| Anonymous IP Address | Sign-in from anonymous proxy/VPN | Medium |
| Atypical Travel | Impossible travel between locations | Medium |
| Unfamiliar Sign-In Properties | Unusual device/location/app | Low |
| Malicious IP Address | Sign-in from known malicious IP | High |
| Password Spray | Evidence of password spray attack | High |

## Investigation Workflow

### Step 1: Triage
1. **Identify affected users**
   \`\`\`json
   {
     "action": "listRiskyUsers",
     "riskLevel": "${riskLevel}",
     "top": 50
   }
   \`\`\`

2. **Get risk detection details**
   \`\`\`json
   {
     "action": "getRiskDetection",
     "riskDetectionId": "<detection-id>"
   }
   \`\`\`

### Step 2: Analyze User Activity
\`\`\`json
{
  "endpoint": "/users/<user-id>/activities",
  "method": "GET"
}
\`\`\`

Review:
- Recent sign-in locations and IPs
- Device types and browsers used
- Applications accessed
- Changes to account settings
- Mail forwarding rules added

### Step 3: Check for Compromise Indicators
\`\`\`json
{
  "category": "AuditLogs",
  "filter": {
    "userId": "<user-id>",
    "operations": ["Update user", "Add member to role", "Add OAuth2PermissionGrant"]
  }
}
\`\`\`
**Tool:** \`search_audit_log\`

Look for:
- New MFA methods registered
- Password changes
- New OAuth app consents
- Role assignments
- Mailbox rules created
- Forwarding addresses added

## Response Actions

### Immediate Response (High Risk)

#### 1. Confirm User as Compromised
\`\`\`json
{
  "action": "confirmRiskyUserCompromised",
  "userId": "<user-id>"
}
\`\`\`
**Tool:** \`manage_identity_protection\`

This will:
- Set user risk to high
- Trigger risk-based CA policies
- Force password reset on next sign-in

#### 2. Reset User Password
\`\`\`json
{
  "endpoint": "/users/<user-id>/authentication/methods/<method-id>/resetPassword",
  "method": "POST"
}
\`\`\`

#### 3. Revoke All Sessions
\`\`\`json
{
  "endpoint": "/users/<user-id>/revokeSignInSessions",
  "method": "POST"
}
\`\`\`
**Tool:** \`call_microsoft_api\`

#### 4. Reset MFA Methods
\`\`\`json
{
  "endpoint": "/users/<user-id>/authentication/methods",
  "method": "GET"
}
\`\`\`
Then delete suspicious methods:
\`\`\`json
{
  "endpoint": "/users/<user-id>/authentication/methods/<method-id>",
  "method": "DELETE"
}
\`\`\`

### Post-Incident Remediation

#### 1. Remove Malicious Artifacts
- Delete unauthorized OAuth app consents
- Remove suspicious mailbox rules
- Revoke delegated permissions
- Remove forwarding rules

#### 2. Dismiss Risk (After Remediation)
\`\`\`json
{
  "action": "dismissRiskyUser",
  "userId": "<user-id>"
}
\`\`\`
**Tool:** \`manage_identity_protection\`

Use only after:
- Password has been reset
- MFA methods have been verified
- All sessions have been revoked
- Suspicious activities have been cleaned up

## Conditional Access Risk Policies

### Auto-Remediate Sign-In Risk
\`\`\`json
{
  "action": "create",
  "displayName": "Require MFA for Risky Sign-Ins",
  "state": "enabled",
  "conditions": {
    "users": {
      "includeUsers": ["All"]
    },
    "applications": {
      "includeApplications": ["All"]
    },
    "signInRisk": {
      "riskLevels": ["medium", "high"]
    }
  },
  "grantControls": {
    "operator": "OR",
    "builtInControls": ["mfa"]
  }
}
\`\`\`
**Tool:** \`manage_conditional_access_policies\`

### Block High-Risk Users
\`\`\`json
{
  "action": "create",
  "displayName": "Block High-Risk Users",
  "state": "enabled",
  "conditions": {
    "users": {
      "includeUsers": ["All"]
    },
    "applications": {
      "includeApplications": ["All"]
    },
    "userRisk": {
      "riskLevels": ["high"]
    }
  },
  "grantControls": {
    "operator": "OR",
    "builtInControls": ["block"]
  }
}
\`\`\`

## Monitoring & Alerting

### Set Up Risk Detection Alerts
Monitor for:
- New high-risk detections
- Users moving to high risk state
- Unusual patterns in risk detections
- Failed risky sign-in attempts

### Regular Review
- Daily: Review high-risk users and detections
- Weekly: Review medium-risk activity
- Monthly: Analyze risk trends and patterns

## Escalation Matrix

| Risk Level | Response Time | Actions |
|------------|--------------|---------|
| Low | 72 hours | Review, monitor |
| Medium | 24 hours | Investigate, require MFA |
| High | 1 hour | Confirm compromise, remediate |
| Confirmed | Immediate | Full incident response |

## Documentation Requirements

For each incident, document:
1. Detection timestamp and type
2. Affected users and resources
3. Investigation findings
4. Actions taken
5. Root cause (if determined)
6. Recommendations for prevention

---

**Incident Type**: ${incidentType}
**Risk Level**: ${riskLevel}
**Response Guide Generated**: ${new Date().toISOString()}`;
    }
  },

  // ===== B2B COLLABORATION SETUP PROMPT =====
  {
    name: 'b2b_collaboration_setup',
    description: 'Configure cross-tenant access settings and B2B collaboration policies for external user access',
    arguments: [
      {
        name: 'collaboration_type',
        description: 'Type: guests_only, b2b_direct_connect, or full_collaboration (default: guests_only)',
        required: false
      },
      {
        name: 'security_level',
        description: 'Security posture: permissive, balanced, or restrictive (default: balanced)',
        required: false
      }
    ],
    handler: async (graphClient, args) => {
      const collaborationType = args?.collaboration_type || 'guests_only';
      const securityLevel = args?.security_level || 'balanced';

      return `# B2B Collaboration & Cross-Tenant Access Setup

## Objective
Configure secure external collaboration with partner organizations.
**Collaboration Type**: ${collaborationType.replace('_', ' ').toUpperCase()}
**Security Level**: ${securityLevel.toUpperCase()}

## Understanding Cross-Tenant Access

### Components
1. **B2B Collaboration** - Guest user access to your tenant
2. **B2B Direct Connect** - External users access without guest account
3. **Cross-Tenant Sync** - Synchronize users across tenants
4. **Inbound/Outbound Settings** - Control access direction

## Current Configuration Review

### 1. Check Default Settings
\`\`\`json
{
  "action": "getDefault"
}
\`\`\`
**Tool:** \`manage_cross_tenant_access\`

### 2. List Partner Configurations
\`\`\`json
{
  "action": "listPartners"
}
\`\`\`
**Tool:** \`manage_cross_tenant_access\`

## Configuration by Security Level

${securityLevel === 'restrictive' ? `
### Restrictive Configuration

#### Default Policy (Block All)
Block all external access by default, allow specific partners only.

\`\`\`json
{
  "action": "updateDefault",
  "b2bCollaborationInbound": {
    "applications": {
      "accessType": "blocked",
      "targets": [{"target": "All", "targetType": "application"}]
    },
    "usersAndGroups": {
      "accessType": "blocked",
      "targets": [{"target": "All", "targetType": "group"}]
    }
  }
}
\`\`\`
**Tool:** \`manage_cross_tenant_access\`

#### Trust Settings (Minimal)
\`\`\`json
{
  "action": "updateDefault",
  "inboundTrust": {
    "isMfaAccepted": false,
    "isCompliantDeviceAccepted": false,
    "isHybridAzureADJoinedDeviceAccepted": false
  }
}
\`\`\`

Then add specific partner configurations for approved organizations.
` : ''}

${securityLevel === 'balanced' ? `
### Balanced Configuration

#### Default Policy (Controlled Access)
\`\`\`json
{
  "action": "updateDefault",
  "inboundTrust": {
    "isMfaAccepted": true,
    "isCompliantDeviceAccepted": true,
    "isHybridAzureADJoinedDeviceAccepted": true
  },
  "b2bCollaborationInbound": {
    "applications": {
      "accessType": "allowed",
      "targets": [{"target": "Office365", "targetType": "application"}]
    },
    "usersAndGroups": {
      "accessType": "allowed",
      "targets": [{"target": "All", "targetType": "group"}]
    }
  }
}
\`\`\`
**Tool:** \`manage_cross_tenant_access\`

This allows:
- Trust partner MFA (no re-authentication)
- Trust partner device compliance
- Access to Office 365 apps
- All users from approved partners
` : ''}

${securityLevel === 'permissive' ? `
### Permissive Configuration

#### Default Policy (Open Access)
\`\`\`json
{
  "action": "updateDefault",
  "inboundTrust": {
    "isMfaAccepted": true,
    "isCompliantDeviceAccepted": true,
    "isHybridAzureADJoinedDeviceAccepted": true
  },
  "b2bCollaborationInbound": {
    "applications": {
      "accessType": "allowed",
      "targets": [{"target": "All", "targetType": "application"}]
    },
    "usersAndGroups": {
      "accessType": "allowed",
      "targets": [{"target": "All", "targetType": "group"}]
    }
  }
}
\`\`\`
**Tool:** \`manage_cross_tenant_access\`

⚠️ **Warning**: This configuration allows broad external access. Ensure you have proper Conditional Access policies in place.
` : ''}

## Conditional Access for Guest Users

### Require MFA for Guests
\`\`\`json
{
  "action": "create",
  "displayName": "Require MFA for Guest Users",
  "state": "enabled",
  "conditions": {
    "users": {
      "includeGuestsOrExternalUsers": {
        "guestOrExternalUserTypes": "b2bCollaborationGuest",
        "externalTenants": {
          "membershipKind": "all"
        }
      }
    },
    "applications": {
      "includeApplications": ["All"]
    }
  },
  "grantControls": {
    "operator": "OR",
    "builtInControls": ["mfa"]
  }
}
\`\`\`
**Tool:** \`manage_conditional_access_policies\`

### Block Guests from Sensitive Apps
\`\`\`json
{
  "action": "create",
  "displayName": "Block Guest Access to Admin Portals",
  "state": "enabled",
  "conditions": {
    "users": {
      "includeGuestsOrExternalUsers": {
        "guestOrExternalUserTypes": "b2bCollaborationGuest"
      }
    },
    "applications": {
      "includeApplications": [
        "797f4846-ba00-4fd7-ba43-dac1f8f63013",
        "0000000c-0000-0000-c000-000000000000"
      ]
    }
  },
  "grantControls": {
    "operator": "OR",
    "builtInControls": ["block"]
  }
}
\`\`\`

## Guest User Management

### Review Current Guests
\`\`\`json
{
  "endpoint": "/users",
  "method": "GET",
  "queryParams": {
    "$filter": "userType eq 'Guest'",
    "$select": "id,displayName,mail,createdDateTime,signInActivity"
  }
}
\`\`\`
**Tool:** \`call_microsoft_api\`

### Identify Stale Guest Accounts
\`\`\`json
{
  "endpoint": "/users",
  "method": "GET",
  "queryParams": {
    "$filter": "userType eq 'Guest'",
    "$select": "id,displayName,signInActivity"
  }
}
\`\`\`
Review \`signInActivity.lastSignInDateTime\` for inactive guests.

## External Sharing Settings (SharePoint)

### Review Current Settings
\`\`\`json
{
  "endpoint": "/admin/sharepoint/settings",
  "method": "GET"
}
\`\`\`
**Tool:** \`call_microsoft_api\`

Configure:
- External sharing level (Anyone, New & Existing Guests, Existing Guests Only, Organization Only)
- Default sharing link type
- Sharing restrictions by domain

## Monitoring External Access

### Audit External User Activity
\`\`\`json
{
  "category": "AuditLogs",
  "filter": {
    "targetResources": "external user"
  }
}
\`\`\`
**Tool:** \`search_audit_log\`

### Regular Review Checklist
- [ ] Review guest user list monthly
- [ ] Remove inactive guests (90+ days no sign-in)
- [ ] Audit guest access to sensitive resources
- [ ] Review cross-tenant access logs
- [ ] Validate partner trust settings

## Best Practices

1. **Principle of Least Privilege** - Grant minimum necessary access
2. **Regular Access Reviews** - Quarterly review of guest access
3. **Time-Bound Access** - Use expiration for guest invitations
4. **MFA Requirements** - Always require MFA for external users
5. **Monitoring** - Alert on unusual external access patterns
6. **Domain Restrictions** - Limit to known partner domains

---

**Collaboration Type**: ${collaborationType}
**Security Level**: ${securityLevel}
**Configuration Date**: ${new Date().toISOString().split('T')[0]}`;
    }
  },

  // ===== ZERO TRUST IMPLEMENTATION PROMPT =====
  {
    name: 'zero_trust_implementation',
    description: 'Comprehensive guide for implementing Zero Trust security architecture in Microsoft 365',
    arguments: [
      {
        name: 'maturity_level',
        description: 'Current maturity: initial, developing, or advanced (default: initial)',
        required: false
      },
      {
        name: 'priority_area',
        description: 'Priority: identity, devices, data, network, or all (default: all)',
        required: false
      }
    ],
    handler: async (graphClient, args) => {
      const maturity = args?.maturity_level || 'initial';
      const priority = args?.priority_area || 'all';

      return `# Zero Trust Implementation Guide for Microsoft 365

## Overview
Implement Zero Trust security architecture following Microsoft's Zero Trust model.
**Current Maturity**: ${maturity.toUpperCase()}
**Priority Area**: ${priority.toUpperCase()}

## Zero Trust Principles

1. **Verify Explicitly** - Always authenticate and authorize based on all available data points
2. **Use Least Privilege Access** - Limit user access with Just-In-Time and Just-Enough-Access
3. **Assume Breach** - Minimize blast radius and segment access, verify end-to-end encryption

## Zero Trust Pillars

### 1. Identity (Foundation)
Every access request starts with identity verification.

#### Assessment
\`\`\`json
{
  "action": "listRiskyUsers",
  "riskLevel": "high"
}
\`\`\`
**Tool:** \`manage_identity_protection\`

\`\`\`json
{
  "action": "list"
}
\`\`\`
**Tool:** \`manage_conditional_access_policies\`

#### Required Controls
${maturity === 'initial' ? `
**Initial Maturity:**
- [ ] MFA enabled for all users
- [ ] Legacy authentication blocked
- [ ] Password protection configured
- [ ] Self-service password reset enabled
- [ ] Basic Conditional Access policies
` : ''}
${maturity === 'developing' ? `
**Developing Maturity:**
- [ ] Risk-based Conditional Access
- [ ] Passwordless authentication pilot
- [ ] Privileged Identity Management (PIM)
- [ ] Identity Governance access reviews
- [ ] Named locations configured
` : ''}
${maturity === 'advanced' ? `
**Advanced Maturity:**
- [ ] Continuous access evaluation
- [ ] Authentication strength policies
- [ ] Cross-tenant access controls
- [ ] FIDO2 security keys deployed
- [ ] Automated threat response
` : ''}

#### Implementation
\`\`\`json
{
  "action": "create",
  "displayName": "Zero Trust - Require MFA All Users",
  "state": "enabledForReportingButNotEnforced",
  "conditions": {
    "users": {"includeUsers": ["All"]},
    "applications": {"includeApplications": ["All"]}
  },
  "grantControls": {
    "operator": "OR",
    "builtInControls": ["mfa"]
  }
}
\`\`\`
**Tool:** \`manage_conditional_access_policies\`

### 2. Devices
Verify device health before granting access.

#### Assessment
\`\`\`json
{
  "action": "list",
  "platform": "all"
}
\`\`\`
**Tool:** \`manage_intune_windows_devices\` and \`manage_intune_macos_devices\`

#### Required Controls
- [ ] Device enrollment required
- [ ] Compliance policies configured
- [ ] BitLocker/FileVault encryption
- [ ] Endpoint detection and response
- [ ] Conditional Access requiring compliant devices

#### Implementation
\`\`\`json
{
  "action": "create",
  "displayName": "Zero Trust - Require Compliant Device",
  "state": "enabledForReportingButNotEnforced",
  "conditions": {
    "users": {"includeUsers": ["All"]},
    "applications": {"includeApplications": ["Office365"]},
    "platforms": {"includePlatforms": ["windows", "macOS", "iOS", "android"]}
  },
  "grantControls": {
    "operator": "AND",
    "builtInControls": ["mfa", "compliantDevice"]
  }
}
\`\`\`
**Tool:** \`manage_conditional_access_policies\`

### 3. Data Protection
Protect data at rest and in transit.

#### Assessment
\`\`\`json
{
  "action": "list"
}
\`\`\`
**Tool:** \`manage_sensitivity_labels\` and \`manage_dlp_policies\`

#### Required Controls
- [ ] Sensitivity labels defined and published
- [ ] DLP policies for sensitive data types
- [ ] Encryption for sensitive content
- [ ] External sharing controls
- [ ] Information barriers (if needed)

### 4. Applications
Verify application health and apply appropriate policies.

#### Assessment
\`\`\`json
{
  "action": "list"
}
\`\`\`
**Tool:** \`manage_azure_ad_apps\`

#### Required Controls
- [ ] App consent policies configured
- [ ] OAuth app governance
- [ ] App protection policies (MAM)
- [ ] Session controls for cloud apps

### 5. Network
Assume the network is hostile.

#### Required Controls
- [ ] Named locations for trusted networks
- [ ] Conditional Access based on location
- [ ] VPN/Private Access configuration
- [ ] Network segmentation

## Implementation Roadmap

### Phase 1: Foundation (Weeks 1-4)
1. Enable MFA for all users
2. Block legacy authentication
3. Configure password protection
4. Enable SSPR
5. Create break-glass accounts

### Phase 2: Identity Protection (Weeks 5-8)
1. Deploy Conditional Access policies
2. Configure named locations
3. Enable Identity Protection
4. Implement risk-based policies
5. Configure sign-in frequency

### Phase 3: Device Security (Weeks 9-12)
1. Enroll devices in Intune
2. Create compliance policies
3. Deploy configuration profiles
4. Enable BitLocker/FileVault
5. Configure device-based CA

### Phase 4: Data Protection (Weeks 13-16)
1. Define sensitivity labels
2. Create label policies
3. Configure DLP policies
4. Enable auto-labeling
5. Configure external sharing

### Phase 5: Continuous Improvement (Ongoing)
1. Regular access reviews
2. Policy optimization
3. New feature adoption
4. Security assessment
5. User training

## Monitoring & Metrics

### Key Performance Indicators
- MFA adoption rate (target: 100%)
- Device compliance rate (target: >95%)
- Risky sign-ins blocked (target: 100% high-risk)
- Data classification coverage (target: >80%)
- Legacy auth usage (target: 0%)

### Monitoring Tools
\`\`\`json
{
  "action": "listRiskDetections",
  "riskLevel": "high"
}
\`\`\`
**Tool:** \`manage_identity_protection\`

### Regular Reports
- Weekly: Risk detection summary
- Monthly: Compliance posture
- Quarterly: Full security assessment

## Quick Wins

### Immediate Actions (Day 1)
1. Enable Security Defaults (if no CA policies)
2. Block legacy authentication
3. Enable MFA for admins
4. Create break-glass accounts
5. Enable sign-in and audit logs

### Week 1 Actions
1. Deploy basic CA policies
2. Configure named locations
3. Enable self-service password reset
4. Review privileged accounts
5. Enable Identity Protection

---

**Maturity Level**: ${maturity}
**Priority Area**: ${priority}
**Implementation Guide Generated**: ${new Date().toISOString()}`;
    }
  }
];
