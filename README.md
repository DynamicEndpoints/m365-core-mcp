## Recent Enhancements (May 3, 2025)

**MCP and HTTP Streaming Updates:**
- Updated MCP SDK to version 1.12.0
- Enhanced HTTP streaming support with both stateful and stateless modes
- Added environment variables for configuring HTTP transport options

## Previous Enhancements (April 4, 2025)

Added several new tools to expand Microsoft Entra ID management and Security & Compliance capabilities:

**Entra ID Management:**
- `manage_azure_ad_roles`: Manage Entra ID directory roles and assignments.
- `manage_azure_ad_apps`: Manage Entra ID application registrations (list, view, owners).
- `manage_azure_ad_devices`: Manage Entra ID device objects (list, view, enable/disable/delete).
- `manage_service_principals`: Manage Entra ID Service Principals (list, view, owners).

**Generic API Access:**
- `call_microsoft_api`: Call arbitrary Microsoft Graph (including Entra APIs) or Azure Resource Management API endpoints.

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

### SharePoint Management
- Create and manage SharePoint sites
- Configure site settings and permissions
- Create and manage SharePoint lists
- Add, update, and retrieve list items
- Manage site users and permissions

### Distribution List Management
- Create and delete distribution lists
- Add/remove members
- Update list settings
- Configure list properties (visibility, moderation, etc.)

### Security Group Management
- Create and delete security groups
- Manage group membership
- Configure mail-enabled settings
- Update group properties

### Microsoft 365 Group Management
- Create and delete M365 groups
- Manage owners and members
- Configure group settings
- Control external access

### Exchange Settings Management
- Configure mailbox settings
- Manage transport rules
- Set organization-wide policies
- Configure retention policies

### User Management
- Get and update user settings
- Configure mailbox properties
- Manage user permissions

### Offboarding Process
- Disable user accounts
- Revoke access tokens
- Backup user data
- Convert to shared mailbox
- Automated cleanup

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
   - Required permissions:
     - Directory.ReadWrite.All
     - Group.ReadWrite.All
     - User.ReadWrite.All
     - Mail.ReadWrite
     - MailboxSettings.ReadWrite
     - Organization.ReadWrite.All
     - Sites.ReadWrite.All
     - Sites.Manage.All

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

### Resources

The server provides the following resources:

- `m365://users/current` - Information about the currently authenticated user
- `m365://tenant/info` - Information about the Microsoft 365 tenant
- `m365://sharepoint/sites` - List of SharePoint sites in the tenant
- `m365://sharepoint/admin/settings` - SharePoint admin settings

Dynamic resources with URI templates:

- `m365://users/{userId}` - Information about a specific user
- `m365://groups/{groupId}` - Information about a specific group
- `m365://sharepoint/sites/{siteId}` - Information about a specific SharePoint site
- `m365://sharepoint/sites/{siteId}/lists` - Lists in a specific SharePoint site
- `m365://sharepoint/sites/{siteId}/lists/{listId}` - Information about a specific SharePoint list
- `m365://sharepoint/sites/{siteId}/lists/{listId}/items` - Items in a specific SharePoint list

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
