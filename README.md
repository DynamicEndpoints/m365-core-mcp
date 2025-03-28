# Microsoft 365 Core MCP Server

An MCP server that provides tools for managing Microsoft 365 core services including:
- Distribution Lists
- Security Groups
- Microsoft 365 Groups
- Exchange Settings
- User Management
- Offboarding Processes

## Features

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

5. Build the server:
   ```bash
   npm run build
   ```

6. Start the server:
   ```bash
   npm start
   ```

## Usage

The server provides MCP tools that can be used to manage various aspects of Microsoft 365. Each tool accepts specific parameters and returns structured responses.

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
```

## Error Handling

The server implements comprehensive error handling:
- Input validation for all parameters
- Graph API error handling
- Token refresh management
- Detailed error messages

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

MIT
