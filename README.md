## Latest Enhancements (December 2024)

**Comprehensive Microsoft 365 Policy Management Expansion:**
- **Added 10 new policy management tools** covering all major Microsoft 365 products and services
- **30+ policy types supported** across security, compliance, governance, and productivity
- **Full lifecycle management** with create, read, update, delete, enable/disable operations
- **Enterprise-ready features** including policy assignment, targeting, and multi-location support

**New Policy Management Tools:**
- `manage_retention_policies` - Data retention across SharePoint, Exchange, Teams, OneDrive
- `manage_sensitivity_labels` - Information protection with encryption and content marking
- `manage_information_protection_policies` - Label policies and organization-wide settings
- `manage_conditional_access_policies` - Identity and access security with MFA, device compliance
- `manage_defender_policies` - Advanced threat protection (Safe Attachments, Safe Links, Anti-Phishing)
- `manage_teams_policies` - Teams governance (messaging, meetings, calling, apps)
- `manage_exchange_policies` - Email security (OWA, ActiveSync, address book policies)
- `manage_sharepoint_governance_policies` - Content and sharing governance
- `manage_security_alert_policies` - Security event monitoring and automated responses

**Policy Types Covered:**
- **Security**: Conditional Access, Defender for Office 365 (Safe Attachments/Links, Anti-Phishing/Malware/Spam)
- **Compliance**: DLP, Retention Policies, Sensitivity Labels, Information Protection
- **Governance**: SharePoint Sharing/Access Policies, Information Barriers, Retention Labels
- **Productivity**: Teams (Messaging/Meeting/Calling/App Setup), Exchange (OWA/ActiveSync/Address Book)
- **Monitoring**: Security and Compliance Alert Policies with automated notifications

**Key Features:**
- Granular control with complex conditions and rules
- Multi-location and multi-target support
- Policy assignment to users, groups, and roles
- Enable/disable functionality for testing
- Comprehensive validation with Zod schemas
- Type-safe implementations with full TypeScript support

For complete documentation, examples, and best practices, see:
- [Policy Management Implementation Guide](./POLICY_MANAGEMENT_EXPANSION_COMPLETE.md)
- [Quick Reference Guide](./POLICY_MANAGEMENT_QUICK_REFERENCE.md)

## Previous Enhancements (September 25, 2025)

**Universal Microsoft Graph API Framework - Complete Transformation:**
- **Transformed from specialized tool to universal Graph API gateway** with access to 1000+ Microsoft Graph endpoints
- **Dynamic Tool Generation System**: Automatically discovers and creates tools for all Graph API endpoints at runtime
- **Advanced Graph API Features**: Batch operations, delta queries, webhook subscriptions, and advanced search
- **Comprehensive Service Coverage**: Teams, OneNote, Planner, To Do, Bookings, Security, Analytics, and more
- **Enhanced Authentication**: Multi-scope token caching with automatic scope detection for all Graph categories
- **Real-time Capabilities**: Webhook subscriptions for live change notifications across all Microsoft 365 services

**New Advanced Graph API Tools:**
- `execute_graph_batch` - Execute up to 20 Graph requests in a single high-performance batch operation
- `execute_delta_query` - Efficiently track changes to any Graph resource using delta queries
- `manage_graph_subscriptions` - Create, update, delete, and list webhook subscriptions for real-time notifications
- `execute_graph_search` - Advanced search across Microsoft 365 content with aggregations and filtering

**Dynamic Category Tools (Generated at Runtime):**
- `manage_teams_resources` - Complete Microsoft Teams management (teams, channels, messages, meetings, chat)
- `manage_productivity_resources` - OneNote notebooks/pages, Planner plans/tasks, To Do lists, Bookings appointments
- `manage_security_resources` - Security incidents, threat intelligence, advanced alerts, Defender integration
- `manage_analytics_resources` - Usage reports, activity insights, trending documents, user analytics

**Enhanced Windows Device Management:**
- `manage_intune_windows_devices` - Complete Windows device lifecycle management in Intune
- `manage_intune_windows_policies` - Windows configuration and compliance policy management
- `manage_intune_windows_apps` - Windows application deployment and management
- `manage_intune_windows_compliance` - Windows device compliance assessment and reporting

**Technical Architecture Improvements:**
- **GraphMetadataService**: Auto-discovers Graph endpoints and generates schemas dynamically
- **DynamicToolGenerator**: Creates tools at runtime based on Graph API metadata
- **GraphAdvancedFeatures**: Implements batch operations, webhooks, delta queries, and search
- **Enhanced Error Handling**: Intelligent troubleshooting with Graph-specific error interpretation
- **Performance Optimizations**: Token caching, batch operations, pagination, and retry logic
- **Smithery Integration**: All 40+ tools properly configured for Smithery discovery

**Scope Coverage Expansion:**
- **Microsoft Teams**: Team.ReadBasic.All, Channel.Create, ChannelMessage.Send, OnlineMeetings.ReadWrite
- **Productivity Apps**: Notes.ReadWrite, Tasks.ReadWrite, Bookings.ReadWrite.All
- **Advanced Security**: SecurityIncident.ReadWrite.All, ThreatIntelligence.Read.All
- **Analytics & Reports**: Reports.Read.All, Sites.Read.All for insights and trending data
- **Power Platform**: Power BI API integration for datasets, reports, and dashboards

This transformation makes the M365 MCP server the **definitive solution for Microsoft 365 automation**, providing unprecedented access to the entire Microsoft Graph API ecosystem with advanced features and optimal performance.

**Previous HTTP Transport Migration (September 25, 2025):**
- Migrated M365 Core MCP Server from STDIO to HTTP transport
- Added Express.js HTTP server with `/mcp` endpoint
- Implemented CORS configuration for browser compatibility
- Added configuration parsing from HTTP requests (Smithery integration)
- Updated Dockerfile for HTTP container deployment (port 8081)
- Updated smithery.yaml to use container runtime with HTTP transport
- Added HTTP development and testing scripts
- Created comprehensive HTTP transport test suite
- Maintained backward compatibility with STDIO transport
- Added support for both stateless and stateful HTTP modes
- Added health and capabilities endpoints for monitoring

## Previous Enhancements (June 16, 2025)

**Extended Resources and Prompts (40 Resources + 5 Comprehensive Prompts):**
- Added 40 additional Microsoft 365 resources covering security, compliance, device management, and collaboration
- Implemented 5 intelligent prompts for automated analysis and recommendations:
  - **Security Assessment**: Comprehensive security posture analysis with recommendations
  - **Compliance Review**: Framework-specific compliance gap analysis (SOC2, ISO27001, NIST, GDPR, HIPAA)
  - **User Access Review**: Individual and organization-wide access rights analysis
  - **Device Compliance Analysis**: Intune device management and compliance assessment
  - **Collaboration Governance**: Teams and SharePoint governance analysis
- Enhanced resource coverage includes:
  - Security alerts, incidents, and conditional access policies
  - Intune device management, apps, and compliance policies
  - Extended user, group, and team information
  - Information protection and DLP policies
  - Audit logs and privileged access data

For detailed information about all new resources and prompts, see [EXTENDED_FEATURES.md](./EXTENDED_FEATURES.md).

## Recent Enhancements (June 7, 2025)

**TypeScript Error Resolution & Compliance Module Enhancements:**
- Resolved all TypeScript errors in `src/server.ts` and `src/handlers/compliance-handler.ts` related to incorrect tool registration syntax and type mismatches.
- Enhanced the compliance module to include comprehensive support for CIS (Center for Internet Security) controls.
- Updated `ComplianceFrameworkArgs` to recognize 'cis' as a valid framework.
- Corrected parameter parsing in compliance handler functions to properly handle string-to-number conversions for implementation groups.

**Conditional Access Policy Review & Reporting:**
- Implemented functionality to retrieve and review Microsoft Entra Conditional Access policies.


## Recent Enhancements (May 3, 2025)

**MCP and HTTP Streaming Updates:**
- Updated MCP SDK to version 1.12.0
- Enhanced HTTP streaming support with both stateful and stateless modes
- Added environment variables for configuring HTTP transport options

## Previous Enhancements (April 4, 2025)

Added several new tools to expand Microsoft Entra ID management and Security & Compliance capabilities:

**Entra ID Management:**
- `manage_azuread_roles`: Manage Entra ID directory roles and assignments.
- `manage_azuread_apps`: Manage Entra ID application registrations (list, view, owners).
- `manage_azuread_devices`: Manage Entra ID device objects (list, view, enable/disable/delete).
- `manage_service_principals`: Manage Entra ID Service Principals (list, view, owners).

**Generic API Access:**
- `dynamicendpoints m365 assistant`: Call arbitrary Microsoft Graph (including Entra APIs) or Azure Resource Management API endpoints.

**Security & Compliance:**
- `search_audit_log`: Search the Entra ID Unified Audit Log.
- `manage_alerts`: List and view security alerts from Microsoft security products.

**Note:** Ensure the associated Entra ID App Registration has the necessary Graph API permissions and Azure RBAC roles for these tools to function correctly.

---

# Microsoft 365 Core MCP Server

[![smithery badge](https://smithery.ai/badge/@DynamicEndpoints/m365-core-mcp)](https://smithery.ai/server/@DynamicEndpoints/m365-core-mcp)

An MCP server that provides tools for managing Microsoft 365 core services including:
- Distribution Lists
- Security Groups
- Microsoft 365 Groups
- Exchange Settings
- User Management
- Offboarding Processes
- SharePoint Sites and Lists

## Features

### Core Microsoft 365 Management
- **Distribution Lists**: Create, delete, manage membership and settings
- **Security Groups**: Full lifecycle management with mail-enabled options
- **Microsoft 365 Groups**: Create, configure, and manage owners/members
- **Exchange Settings**: Mailbox, transport, organization, and retention policies
- **User Management**: Get and update user settings and configurations
- **Offboarding Processes**: Automated user offboarding with configurable options

### SharePoint Management
- **Site Management**: Create, update, delete sites with template support
- **List Management**: Create, configure, and manage SharePoint lists
- **Item Management**: Add, update, and retrieve list items
- **Permissions**: Manage site users and permissions
- **Settings**: Configure site-level and organization settings

### Azure AD Management
- **Role Management**: Assign and manage directory roles and role assignments
- **Application Management**: Manage app registrations, owners, and settings
- **Device Management**: Enable, disable, delete Azure AD devices
- **Service Principals**: Manage service principal objects and ownership

### Security & Compliance
- **Audit Logging**: Search and analyze Microsoft 365 Unified Audit Log
- **Security Alerts**: List, view, and manage security alerts across Microsoft products
- **Data Loss Prevention**: Create, configure, and manage DLP policies and incidents
- **Sensitivity Labels**: Manage Microsoft Purview sensitivity labels and policies
- **Compliance Frameworks**: Support for HITRUST, ISO27001, SOC2, CIS Controls
- **Assessment & Monitoring**: Automated compliance assessments and continuous monitoring
- **Evidence Collection**: Automated evidence gathering for compliance audits
- **Gap Analysis**: Cross-framework compliance gap analysis and remediation planning

### Intune Device Management (macOS Focus)
- **Device Inventory**: List, filter, and manage macOS devices in Intune
- **Policy Management**: Create, deploy, and monitor macOS configuration policies
- **Application Management**: Deploy and manage macOS applications via Intune
- **Compliance Monitoring**: Track and enforce macOS device compliance policies

### Advanced Features
- **Dynamic API Access**: Call arbitrary Microsoft Graph and Azure Resource Management APIs
- **Real-time Capabilities**: Server-sent events, progress reporting, streaming responses
- **Intelligent Prompts**: 5 comprehensive analysis prompts for security, compliance, and governance
- **Extended Resources**: 44 resources covering security, compliance, device management, and collaboration
- **Modern MCP Features**: Enhanced error handling, response validation, lazy loading

## Setup

### Installing via Smithery

To install Microsoft 365 Core Server for Claude Desktop automatically via [Smithery](https://smithery.ai/server/@DynamicEndpoints/m365-core-mcp):

```bash
npx -y @smithery/cli install @DynamicEndpoints/m365-core-mcp --client claude
```

### Installing Manually
1. Clone the repository
2. Install dependencies:
   ```bash
   npm install
   ```
3. Create a `.env` file based on `.env.example`:
   ```
   MS_TENANT_ID=your-tenant-id
   MS_CLIENT_ID=your-client-id
   MS_CLIENT_SECRET=your-client-secret
   
   # Optional Configuration
   # LOG_LEVEL=info    # debug, info, warn, error
   # PORT=3000         # Port for HTTP server if needed
   # USE_HTTP=true     # Set to 'true' to use HTTP transport instead of stdio
   # STATELESS=false   # Set to 'true' to use stateless HTTP mode (no session management)
   ```
4. Register an application in Azure AD:
   - **Required Microsoft Graph permissions:**
     - Directory.ReadWrite.All
     - Group.ReadWrite.All
     - User.ReadWrite.All
     - Mail.ReadWrite
     - MailboxSettings.ReadWrite
     - Organization.ReadWrite.All
     - Sites.ReadWrite.All
     - Sites.Manage.All
     - SecurityEvents.ReadWrite.All
     - SecurityActions.ReadWrite.All
     - Device.ReadWrite.All
     - DeviceManagementConfiguration.ReadWrite.All
     - DeviceManagementManagedDevices.ReadWrite.All
     - DeviceManagementApps.ReadWrite.All
     - InformationProtectionPolicy.ReadWrite.All
     - Policy.ReadWrite.ConditionalAccess
     - RoleManagement.ReadWrite.Directory
     - AuditLog.Read.All
     - Reports.Read.All
     - ThreatIndicators.ReadWrite.OwnedBy
     - IdentityRiskyUser.ReadWrite.All
     - IdentityRiskEvent.Read.All

   - **Required Azure RBAC roles** (for Azure Resource Management):
     - Security Admin (for security-related operations)
     - Compliance Administrator (for compliance management)
     - Intune Administrator (for device management)
     - Reports Reader (for audit and reporting functions)

5. Build the server:
   ```bash
   npm run build
   ```

6. Start the server:
   ```bash
   npm start
   ```

## Transport Options

The server supports multiple transport options for MCP communication:

### stdio Transport

By default, the server uses stdio transport, which is ideal for:
- Command-line tools and direct integrations
- Local development and testing
- Integration with Smithery and other MCP clients that support stdio

### HTTP Transport

The server also supports HTTP transport with two modes:

#### Stateful Mode (With Session Management)

This is the default HTTP mode when `USE_HTTP=true` and `STATELESS=false`:
- Maintains session state between requests
- Supports server-to-client notifications via GET requests
- Handles session termination via DELETE requests
- Ideal for long-running sessions and interactive applications
- Provides better performance for multiple requests in the same session

#### Stateless Mode

Enable this mode by setting `USE_HTTP=true` and `STATELESS=true`:
- Creates a new server instance for each request
- No session state is maintained between requests
- Only supports POST requests (GET and DELETE are not supported)
- Ideal for RESTful scenarios where each request is independent
- Better for horizontally scaled deployments without shared session state
- Simpler API wrappers where session management isn't needed

To configure the transport options, set the appropriate environment variables in your `.env` file:
```
USE_HTTP=true     # Use HTTP transport instead of stdio
STATELESS=false   # Use stateful mode with session management (default)
PORT=3000         # Port for the HTTP server
```

## Usage

The server provides MCP tools and resources that can be used to manage various aspects of Microsoft 365. Each tool accepts specific parameters and returns structured responses.

### Tools

The server provides **29 comprehensive tools** for Microsoft 365 management:

#### Core Management Tools
- `manage_distribution_lists` - Create, delete, and manage distribution lists and membership
- `manage_security_groups` - Create, delete, and manage security groups and membership  
- `manage_m365_groups` - Create, delete, and manage Microsoft 365 groups and membership
- `manage_exchange_settings` - Configure mailbox, transport, organization, and retention settings
- `manage_user_settings` - Get and update user settings and configurations
- `manage_offboarding` - Automated user offboarding processes with configurable options

#### SharePoint Management Tools
- `manage_sharepoint_sites` - Create, update, delete SharePoint sites and manage users
- `manage_sharepoint_lists` - Create, update, delete SharePoint lists and manage items

#### Azure AD Management Tools
- `manage_azuread_roles` - Manage Azure AD directory roles and role assignments
- `manage_azuread_apps` - Manage Azure AD application registrations and owners
- `manage_azuread_devices` - Manage Azure AD device objects (enable, disable, delete)
- `manage_service_principals` - Manage Azure AD Service Principals and ownership

#### Security & Compliance Tools
- `search_audit_log` - Search the Microsoft 365 Unified Audit Log
- `manage_alerts` - List and view security alerts from Microsoft security products
- `manage_dlp_policies` - Manage Data Loss Prevention policies and configurations
- `manage_dlp_incidents` - Handle DLP policy violations and incident management
- `manage_sensitivity_labels` - Manage Microsoft Purview sensitivity labels

#### Intune Device Management Tools
- `manage_intune_macos_devices` - Manage Intune macOS devices and enrollment
- `manage_intune_macos_policies` - Configure and deploy macOS device policies
- `manage_intune_macos_apps` - Deploy and manage macOS applications via Intune
- `manage_intune_macos_compliance` - Monitor and enforce macOS device compliance

#### Compliance Framework Tools
- `manage_compliance_frameworks` - Configure compliance frameworks (HITRUST, ISO27001, SOC2)
- `manage_compliance_assessments` - Run and manage compliance assessments
- `manage_compliance_monitoring` - Monitor compliance status and configure alerts
- `manage_evidence_collection` - Collect and manage compliance evidence
- `manage_gap_analysis` - Perform compliance gap analysis and remediation planning
- `manage_cis_compliance` - Manage CIS Controls compliance and benchmarks

#### Audit & Reporting Tools
- `generate_audit_reports` - Generate comprehensive audit reports for various frameworks

#### Dynamic API Access
- `dynamicendpoints m365 assistant` - Call arbitrary Microsoft Graph or Azure Resource Management API endpoints

### Resources

The server provides **44 comprehensive resources** covering security, compliance, device management, and collaboration:

#### Core Resources
- `sharepoint_sites` - SharePoint site information and configuration
- `sharepoint_lists` - SharePoint list structures and metadata  
- `sharepoint_list_items` - Items within SharePoint lists
- `security_incidents` - Microsoft security incidents and details

#### Extended Security Resources (20 resources)
- Security alerts and incidents from Microsoft Defender
- Conditional access policies and assignments
- Privileged access management data
- Threat intelligence and vulnerability assessments
- Identity protection risks and policies
- Authentication methods and security defaults
- Compliance policies and their status
- Data governance and retention policies
- Insider risk management insights
- Security baselines and configurations

#### Device Management Resources (10 resources)
- Intune device inventories and compliance status
- Mobile application management policies
- Device configuration profiles and assignments
- Compliance policies for various platforms
- App protection policies and status
- Device enrollment configurations
- Update policies and deployment rings
- Certificate profiles and management
- Wi-Fi and VPN configuration profiles
- Endpoint protection policies

#### Collaboration Resources (10 resources)
- Microsoft Teams structures and policies
- Exchange Online configurations and settings
- Calendar and scheduling information
- OneDrive storage and sharing policies
- Planner tasks and project management
- Viva Engage (Yammer) communities
- Power Platform environments and apps
- Booking services and appointments
- Whiteboard collaboration data
- Stream video content and policies

#### Extended Dynamic Resources
All resources support URI templates for specific object access:
- `m365://security/alerts/{alertId}` - Specific security alert details
- `m365://devices/{deviceId}` - Individual device information
- `m365://users/{userId}/compliance` - User-specific compliance status
- `m365://teams/{teamId}/governance` - Team governance and policies

### Intelligent Prompts

The server provides **5 comprehensive prompts** for automated analysis and recommendations:

#### Security Assessment Prompt
- **Purpose**: Comprehensive security posture analysis with actionable recommendations
- **Scope**: Security policies, access controls, threat detection, identity protection
- **Output**: Risk assessment, security gaps, remediation roadmap

#### Compliance Review Prompt  
- **Purpose**: Framework-specific compliance gap analysis
- **Frameworks**: SOC2, ISO27001, NIST, GDPR, HIPAA, CIS Controls
- **Scope**: Control implementation status, evidence collection, audit readiness
- **Output**: Compliance dashboard, gap analysis, remediation plans

#### User Access Review Prompt
- **Purpose**: Individual and organization-wide access rights analysis
- **Scope**: Role assignments, group memberships, application access, privileged accounts
- **Output**: Access recommendations, risk-based prioritization, cleanup tasks

#### Device Compliance Analysis Prompt
- **Purpose**: Intune device management and compliance assessment
- **Scope**: Device policies, compliance status, security configurations, app management
- **Output**: Compliance reports, policy recommendations, deployment guidance

#### Collaboration Governance Prompt
- **Purpose**: Teams and SharePoint governance analysis
- **Scope**: Team structures, sharing policies, external access, data governance
- **Output**: Governance recommendations, policy suggestions, compliance alignment

Each prompt provides contextual analysis, actionable insights, and integration with the corresponding management tools for immediate remediation.

### Example Tool Usage

```typescript
// Managing a distribution list
await callTool('manage_distribution_lists', {
  action: 'create',
  displayName: 'Marketing Team',
  emailAddress: 'marketing@company.com',
  members: ['user1@company.com', 'user2@company.com']
});

// Managing security groups
await callTool('manage_security_groups', {
  action: 'create',
  displayName: 'IT Admins',
  description: 'IT Administration Team',
  members: ['admin1@company.com']
});

// Managing Azure AD roles (note: using correct tool name)
await callTool('manage_azuread_roles', {
  action: 'assign_role',
  roleId: 'role-id-here',
  principalId: 'user-id-here'
});

// Managing DLP policies
await callTool('manage_dlp_policies', {
  action: 'create',
  policyName: 'Financial Data Protection',
  rules: [{
    name: 'Block Credit Cards',
    conditions: { contentContainsSensitiveInfo: ['CreditCardNumber'] },
    actions: { blockAccess: true }
  }]
});

// Managing Intune macOS devices
await callTool('manage_intune_macos_devices', {
  action: 'list',
  filters: { complianceState: 'compliant' }
});

// Running compliance assessments
await callTool('manage_compliance_assessments', {
  action: 'run_assessment',
  framework: 'iso27001',
  scope: ['access_control', 'data_protection'],
  settings: {
    automated: true,
    generateRemediation: true
  }
});

// Generating audit reports
await callTool('generate_audit_reports', {
  framework: 'soc2',
  reportType: 'comprehensive',
  dateRange: { start: '2025-01-01', end: '2025-06-16' },
  format: 'pdf',
  includeEvidence: true
});

// Managing Exchange settings
await callTool('manage_exchange_settings', {
  action: 'update',
  settingType: 'mailbox',
  target: 'user@company.com',
  settings: {
    automateProcessing: {
      autoReplyEnabled: true
    }
  }
});

// Managing SharePoint sites
await callTool('manage_sharepoint_sites', {
  action: 'create',
  title: 'Marketing Site',
  description: 'Site for marketing team',
  template: 'STS#0',
  url: 'https://contoso.sharepoint.com/sites/marketing',
  owners: ['user1@company.com'],
  members: ['user2@company.com', 'user3@company.com']
});

// Managing SharePoint lists
await callTool('manage_sharepoint_lists', {
  action: 'create',
  siteId: 'contoso.sharepoint.com,5a14e1cf-e284-4722-8f50-a5e1b2b0a8d6,9528e4bb-7660-4b11-a758-9d8fb3ca295f',
  title: 'Project Tasks',
  description: 'List of project tasks',
  columns: [
    { name: 'Title', type: 'text', required: true },
    { name: 'DueDate', type: 'dateTime' },
    { name: 'Status', type: 'choice', choices: ['Not Started', 'In Progress', 'Completed'] }
  ]
});

// Dynamic API calls for custom scenarios
await callTool('dynamicendpoints m365 assistant', {
  apiType: 'graph',
  path: '/me/messages',
  method: 'get',
  queryParams: { '$top': '10', '$filter': 'isRead eq false' }
});
```

## Implementation Details

### Schema Validation

The server uses Zod for schema validation, providing:
- Runtime type checking for all inputs
- Detailed validation error messages
- Type inference for TypeScript
- Automatic documentation of input schemas

### Error Handling

The server implements comprehensive error handling:
- Input validation for all parameters
- Graph API error handling
- Token refresh management
- Detailed error messages with proper error codes

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

MIT
