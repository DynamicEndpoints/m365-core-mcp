# Microsoft 365 Core MCP Server

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
