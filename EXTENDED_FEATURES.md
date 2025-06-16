# M365 Core MCP Server - Extended Resources and Prompts

## Overview
We've enhanced the M365 Core MCP (Model Context Protocol) server with 40 additional resources and comprehensive prompts to provide better insights and automation capabilities for Microsoft 365 environments.

## Added Resources (40 total)

### Security Resources (1-20)
1. **security_alerts** - `m365://security/alerts` - Recent security alerts
2. **security_incidents** - `m365://security/incidents` - Security incidents
3. **conditional_access_policies** - `m365://identity/conditionalAccess/policies` - Conditional Access policies
4. **applications** - `m365://applications` - Azure AD applications
5. **service_principals** - `m365://servicePrincipals` - Service principals
6. **directory_roles** - `m365://directoryRoles` - Directory roles
7. **privileged_access** - `m365://privilegedAccess/azureAD/resources` - Privileged access resources
8. **audit_logs_signin_extended** - Enhanced sign-in logs
9. **audit_logs_directory_extended** - Enhanced directory audit logs
10. **risky_users_extended** - Extended risky users information
11. **threat_assessment_extended** - Threat assessment requests
12. **security_score_extended** - Extended secure score data
13. **compliance_policies_dlp_extended** - Extended DLP policies
14. **retention_policies_extended** - Extended retention policies
15. **sensitivity_labels_extended** - Extended sensitivity labels
16. **communication_compliance_extended** - Extended communication compliance
17. **ediscovery_cases_extended** - Extended eDiscovery cases
18. **information_protection_extended** - Extended information protection labels
19. **subscribed_skus_extended** - Extended SKU information
20. **directory_role_assignments** - Directory role member assignments

### Device Management Resources (21-30)
21. **intune_devices_extended** - Extended Intune device information
22. **intune_apps_extended** - Extended mobile apps data
23. **intune_compliance_policies_extended** - Extended compliance policies
24. **intune_configuration_policies_extended** - Extended configuration profiles
25. **device_info_extended** - Detailed device information by ID
26. **app_assignments_extended** - Extended app assignment details
27. **policy_assignments_extended** - Extended policy assignment details
28. **user_licenses_extended** - Extended user license information
29. **user_groups_extended** - Extended user group memberships
30. **group_members_extended** - Extended group member information

### Collaboration Resources (31-40)
31. **teams_list_extended** - Extended Teams information
32. **mail_folders_extended** - Extended mail folder data
33. **calendar_events_extended** - Extended calendar events
34. **onedrive_extended** - Extended OneDrive information
35. **planner_plans_extended** - Extended Planner plans
36. **user_messages_extended** - Extended user messages by ID
37. **user_calendar_extended** - Extended user calendar by ID
38. **user_drive_extended** - Extended user drive by ID
39. **team_channels_extended** - Extended team channels by team ID
40. **team_members_extended** - Extended team members by team ID

## Added Prompts (5 comprehensive prompts)

### 1. Security Assessment (`security_assessment`)
**Purpose**: Analyze M365 security posture and provide recommendations
**Parameters**:
- `scope` (optional): Security assessment scope (identity, data, devices, applications)
- `timeframe` (optional): Assessment timeframe (last 7 days, 30 days, 90 days)

**Features**:
- Gathers security alerts, risky users, conditional access policies, and secure scores
- Provides comprehensive security analysis
- Identifies critical risks and vulnerabilities
- Offers prioritized recommendations
- Includes compliance gap analysis

### 2. Compliance Review (`compliance_review`)
**Purpose**: Generate compliance review and gap analysis
**Parameters**:
- `framework` (optional): Compliance framework (SOC2, ISO27001, NIST, GDPR, HIPAA)
- `scope` (optional): Review scope (policies, controls, data protection)

**Features**:
- Reviews DLP policies, retention labels, sensitivity labels
- Analyzes recent audit events
- Provides framework-specific compliance assessment
- Identifies gaps and non-conformities
- Offers remediation recommendations

### 3. User Access Review (`user_access_review`)
**Purpose**: Analyze user access rights and permissions
**Parameters**:
- `userId` (optional): Specific user ID to review
- `focus` (optional): Review focus (permissions, licenses, group memberships, recent activity)

**Features**:
- Single user or organization-wide analysis
- Reviews licenses, groups, and sign-in patterns
- Identifies excessive permissions
- Suggests license optimizations
- Provides access governance recommendations

### 4. Device Compliance Analysis (`device_compliance_analysis`)
**Purpose**: Analyze device compliance and management status
**Parameters**:
- `platform` (optional): Device platform (Windows, iOS, Android, macOS)
- `complianceStatus` (optional): Filter by compliance status

**Features**:
- Reviews managed devices, apps, and policies
- Analyzes compliance status
- Identifies configuration gaps
- Provides security posture assessment
- Offers device management improvements

### 5. Collaboration Governance (`collaboration_governance`)
**Purpose**: Analyze Teams and collaboration governance
**Parameters**:
- `governanceArea` (optional): Focus area (teams, sharing, guest access, data protection)
- `timeframe` (optional): Analysis timeframe (30 days, 90 days, 6 months)

**Features**:
- Reviews Teams, groups, sites, and applications
- Assesses governance maturity
- Identifies sprawl and proliferation issues
- Analyzes guest access risks
- Provides governance recommendations

## Implementation Details

### File Structure
- **src/extended-resources.ts**: Contains all extended resources and prompts
- **src/server.ts**: Enhanced to import and initialize extended functionality

### Key Features
- **Comprehensive Coverage**: 40 resources covering security, compliance, device management, and collaboration
- **Intelligent Prompts**: 5 specialized prompts for different analysis scenarios
- **Dynamic Resources**: Template-based resources with parameters for specific data retrieval
- **Error Handling**: Robust error handling with meaningful error messages
- **JSON Output**: All resources return structured JSON data for easy parsing

### Usage Examples

#### Using Resources
```typescript
// Access security alerts
GET m365://security/alerts

// Get user-specific information
GET m365://users/{userId}/messages/extended

// Get team channels
GET m365://teams/{teamId}/channels/extended
```

#### Using Prompts
```typescript
// Security assessment for identity scope
security_assessment(scope: "identity", timeframe: "30 days")

// Compliance review for NIST framework
compliance_review(framework: "NIST", scope: "policies")

// User access review for specific user
user_access_review(userId: "user@domain.com", focus: "permissions")
```

## Benefits

1. **Enhanced Visibility**: 40 additional data sources provide comprehensive M365 insights
2. **Intelligent Analysis**: Prompts deliver contextual analysis and recommendations
3. **Automation Ready**: Structured outputs enable automated governance and compliance workflows
4. **Flexible Querying**: Template-based resources allow specific data retrieval
5. **Security Focus**: Strong emphasis on security, compliance, and governance use cases

## Next Steps

1. **Test Integration**: Verify all resources and prompts work correctly
2. **Documentation**: Create detailed API documentation for each resource and prompt
3. **Performance Optimization**: Monitor and optimize resource access patterns
4. **Additional Prompts**: Consider adding more specialized prompts based on user feedback
5. **Custom Dashboards**: Build dashboards leveraging the extended resource data
