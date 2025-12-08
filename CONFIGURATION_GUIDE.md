# M365 Core MCP Server - Configuration Guide

## Connection Error: "Not connected"

This error occurs when the MCP server doesn't have valid Microsoft 365 credentials configured.

## Required Environment Variables

You need to configure three environment variables:

```bash
MS_TENANT_ID=your-tenant-id
MS_CLIENT_ID=your-application-id
MS_CLIENT_SECRET=your-client-secret
```

## Step-by-Step Setup

### 1. Register an Azure AD Application

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Fill in:
   - **Name**: M365 Core MCP Server
   - **Supported account types**: Single tenant
   - **Redirect URI**: Leave blank
5. Click **Register**

### 2. Get Application (Client) ID and Tenant ID

After registration:
- **Application (client) ID**: Copy this value → Use as `MS_CLIENT_ID`
- **Directory (tenant) ID**: Copy this value → Use as `MS_TENANT_ID`

### 3. Create Client Secret

1. In your app registration, go to **Certificates & secrets**
2. Click **New client secret**
3. Add description: "MCP Server Secret"
4. Set expiration (recommend: 24 months)
5. Click **Add**
6. **IMPORTANT**: Copy the **Value** immediately (you can't see it again!) → Use as `MS_CLIENT_SECRET`

### 4. Configure API Permissions

In your app registration, go to **API permissions**:

#### Required Microsoft Graph Permissions (Application):

**User Management**:
- `User.ReadWrite.All`
- `Directory.ReadWrite.All`

**Group Management**:
- `Group.ReadWrite.All`
- `GroupMember.ReadWrite.All`

**Device Management (Intune)**:
- `DeviceManagementConfiguration.ReadWrite.All`
- `DeviceManagementManagedDevices.ReadWrite.All`
- `DeviceManagementServiceConfig.ReadWrite.All`
- `DeviceManagementApps.ReadWrite.All`

**Security & Compliance**:
- `SecurityEvents.ReadWrite.All`
- `ThreatIndicators.ReadWrite.OwnedBy`
- `InformationProtectionPolicy.Read.All`

**SharePoint**:
- `Sites.ReadWrite.All`
- `Sites.FullControl.All`

**Exchange**:
- `Mail.ReadWrite`
- `MailboxSettings.ReadWrite`

**Reports & Audit**:
- `AuditLog.Read.All`
- `Reports.Read.All`

#### Add Permissions:
1. Click **Add a permission**
2. Select **Microsoft Graph**
3. Choose **Application permissions**
4. Search for and add each permission above
5. Click **Add permissions**
6. **IMPORTANT**: Click **Grant admin consent** (requires Global Admin)

### 5. Configure Environment Variables

Choose one of these methods:

#### Option A: .env File (Recommended for Development)

Create a `.env` file in the project root:

```bash
# Azure AD Configuration
MS_TENANT_ID=your-tenant-id-here
MS_CLIENT_ID=your-client-id-here
MS_CLIENT_SECRET=your-client-secret-here

# Optional: Logging
LOG_LEVEL=info
```

Then use with your MCP client that supports .env files.

#### Option B: System Environment Variables (Windows)

```powershell
# PowerShell
[System.Environment]::SetEnvironmentVariable('MS_TENANT_ID', 'your-tenant-id', 'User')
[System.Environment]::SetEnvironmentVariable('MS_CLIENT_ID', 'your-client-id', 'User')
[System.Environment]::SetEnvironmentVariable('MS_CLIENT_SECRET', 'your-client-secret', 'User')
```

```cmd
# Command Prompt
setx MS_TENANT_ID "your-tenant-id"
setx MS_CLIENT_ID "your-client-id"
setx MS_CLIENT_SECRET "your-client-secret"
```

#### Option C: System Environment Variables (macOS/Linux)

Add to `~/.bashrc`, `~/.zshrc`, or `~/.profile`:

```bash
export MS_TENANT_ID="your-tenant-id"
export MS_CLIENT_ID="your-client-id"
export MS_CLIENT_SECRET="your-client-secret"
```

Then reload:
```bash
source ~/.bashrc  # or ~/.zshrc
```

#### Option D: MCP Client Configuration

For Cline or other MCP clients, add to the MCP settings:

```json
{
  "mcpServers": {
    "m365-core": {
      "command": "node",
      "args": ["path/to/m365-core-mcp/build/index.js"],
      "env": {
        "MS_TENANT_ID": "your-tenant-id",
        "MS_CLIENT_ID": "your-client-id",
        "MS_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

### 6. Verify Configuration

Test the connection:

```bash
# Build the project
npm run build

# Test a simple tool
node -e "
process.env.MS_TENANT_ID = 'your-tenant-id';
process.env.MS_CLIENT_ID = 'your-client-id';
process.env.MS_CLIENT_SECRET = 'your-client-secret';
require('./build/server.js');
"
```

## Troubleshooting

### "Not connected" Error
**Cause**: Missing or invalid credentials
**Solution**: 
1. Verify all three environment variables are set
2. Check values are correct (no extra spaces/quotes)
3. Ensure client secret hasn't expired
4. Restart your MCP client after setting variables

### "Insufficient privileges" Error
**Cause**: Missing API permissions or admin consent
**Solution**:
1. Go to Azure Portal > App registrations > Your app
2. Check API permissions tab
3. Ensure all required permissions are added
4. Click "Grant admin consent for [Your Org]"
5. Wait 5-10 minutes for permissions to propagate

### "Invalid client secret" Error
**Cause**: Wrong secret or expired
**Solution**:
1. Go to Azure Portal > App registrations > Your app
2. Certificates & secrets > Client secrets
3. Create new client secret
4. Update `MS_CLIENT_SECRET` environment variable
5. Restart MCP client

### Token Acquisition Errors
**Cause**: Network issues or Azure AD problems
**Solution**:
1. Check internet connectivity
2. Verify tenant ID is correct
3. Try from different network
4. Check Azure AD service status

### Specific Tool Errors

#### Intune Tools Not Working
**Required Permissions**:
- `DeviceManagementConfiguration.ReadWrite.All`
- `DeviceManagementManagedDevices.ReadWrite.All`
- `DeviceManagementServiceConfig.ReadWrite.All`

Verify these are added with admin consent.

#### SharePoint Tools Not Working
**Required Permissions**:
- `Sites.ReadWrite.All` or `Sites.FullControl.All`

#### Security/Compliance Tools Not Working
**Required Permissions**:
- `SecurityEvents.ReadWrite.All`
- `InformationProtectionPolicy.Read.All`

## Security Best Practices

1. **Store Secrets Securely**:
   - Never commit secrets to Git
   - Use `.env` files (add to `.gitignore`)
   - Consider Azure Key Vault for production

2. **Rotate Secrets Regularly**:
   - Set client secret expiration
   - Rotate every 6-12 months
   - Update immediately if compromised

3. **Principle of Least Privilege**:
   - Only add permissions you need
   - Review permissions periodically
   - Create separate apps for different purposes

4. **Monitor Usage**:
   - Review Azure AD sign-in logs
   - Monitor for unusual activity
   - Set up alerts for permission changes

## Testing Your Configuration

Once configured, test with these tools:

1. **List users** (basic test):
```javascript
call_microsoft_api({
  apiType: "graph",
  method: "get",
  path: "/users",
  queryParams: { "$top": "5" }
})
```

2. **List Intune devices**:
```javascript
manage_intune_windows_devices({
  action: "list"
})
```

3. **Check permissions**:
```javascript
call_microsoft_api({
  apiType: "graph",
  method: "get",
  path: "/me/oauth2PermissionGrants"
})
```

## Getting Help

If you continue to have issues:

1. Check the logs for detailed error messages
2. Verify Azure AD app configuration
3. Test with Microsoft Graph Explorer: https://developer.microsoft.com/graph/graph-explorer
4. Review Microsoft Graph documentation: https://learn.microsoft.com/graph/

## Quick Reference

| Variable | Where to Find |
|----------|--------------|
| `MS_TENANT_ID` | Azure Portal > Azure AD > Overview > Tenant ID |
| `MS_CLIENT_ID` | Azure Portal > App registrations > Your app > Application ID |
| `MS_CLIENT_SECRET` | Azure Portal > App registrations > Your app > Certificates & secrets |

## Next Steps

After configuration:
1. Restart your MCP client
2. Try the test commands above
3. Explore available tools with `list_tools`
4. Review prompts with `list_prompts`
5. Access resources with resource URIs (e.g., `m365://security/alerts`)

For Intune policy creation, use the wizard:
```javascript
intune_policy_wizard({
  policy_goal: "your goal here",
  platform: "windows",
  security_level: "standard"
})
