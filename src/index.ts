#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';
import {
  UserManagementArgs,
  OffboardingArgs,
  DistributionListArgs,
  SecurityGroupArgs,
  M365GroupArgs,
  ExchangeSettingsArgs,
} from './types.js';
import { m365CoreTools } from './tool-definitions.js';

// Environment validation
const MS_TENANT_ID = process.env.MS_TENANT_ID ?? '';
const MS_CLIENT_ID = process.env.MS_CLIENT_ID ?? '';
const MS_CLIENT_SECRET = process.env.MS_CLIENT_SECRET ?? '';

if (!MS_TENANT_ID || !MS_CLIENT_ID || !MS_CLIENT_SECRET) {
  throw new Error('Required environment variables (MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET) are missing');
}

class M365CoreServer {
  private server: Server;
  private graphClient: Client;
  private token: string | null = null;

  constructor() {
    this.server = new Server(
      {
        name: 'm365-core-server',
        version: '1.0.0',
      },
      {
        capabilities: {
          tools: {},
          resources: {},
        },
      }
    );

    this.graphClient = Client.init({
      authProvider: async (callback) => {
        try {
          if (!this.token) {
            this.token = await this.getAccessToken();
          }
          callback(null, this.token);
        } catch (error) {
          callback(error as Error, null);
        }
      }
    });

    this.setupHandlers();
    
    // Error handling
    this.server.onerror = (error) => console.error('[MCP Error]', error);
    process.on('SIGINT', async () => {
      await this.server.close();
      process.exit(0);
    });
  }

  private async getAccessToken(): Promise<string> {
    const tokenEndpoint = `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`;
    const scope = 'https://graph.microsoft.com/.default';

    const params = new URLSearchParams();
    params.append('client_id', MS_CLIENT_ID);
    params.append('client_secret', MS_CLIENT_SECRET);
    params.append('grant_type', 'client_credentials');
    params.append('scope', scope);

    const response = await fetch(tokenEndpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: params,
    });

    if (!response.ok) {
      throw new Error('Failed to get access token');
    }

    const data = await response.json();
    return data.access_token;
  }

  private setupHandlers(): void {
    // List available tools
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: m365CoreTools
    }));

    // Handle tool calls
    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      try {
        const { name, arguments: args } = request.params;

        if (!args || typeof args !== 'object') {
          throw new McpError(ErrorCode.InvalidParams, 'Missing or invalid arguments');
        }

        switch (name) {
          case 'manage_distribution_lists': {
            const dlArgs = this.validateDistributionListArgs(args);
            return await this.handleDistributionList(dlArgs);
          }

          case 'manage_security_groups': {
            const sgArgs = this.validateSecurityGroupArgs(args);
            return await this.handleSecurityGroup(sgArgs);
          }

          case 'manage_m365_groups': {
            const m365Args = this.validateM365GroupArgs(args);
            return await this.handleM365Group(m365Args);
          }

          case 'manage_exchange_settings': {
            const exchangeArgs = this.validateExchangeSettingsArgs(args);
            return await this.handleExchangeSettings(exchangeArgs);
          }

          case 'manage_user_settings': {
            const userArgs = this.validateUserManagementArgs(args);
            return await this.handleUserSettings(userArgs);
          }

          case 'manage_offboarding': {
            const offboardingArgs = this.validateOffboardingArgs(args);
            return await this.handleOffboarding(offboardingArgs);
          }
        }

        throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${name}`);
      } catch (error) {
        if (error instanceof McpError) {
          throw error;
        }
        throw new McpError(
          ErrorCode.InternalError,
          `Error executing tool: ${error instanceof Error ? error.message : 'Unknown error'}`
        );
      }
    });
  }

  private validateDistributionListArgs(args: Record<string, unknown>): DistributionListArgs {
    if (!args.action || typeof args.action !== 'string') {
      throw new McpError(ErrorCode.InvalidParams, 'Missing required distribution list parameters');
    }
    return {
      action: args.action as DistributionListArgs['action'],
      listId: typeof args.listId === 'string' ? args.listId : undefined,
      displayName: typeof args.displayName === 'string' ? args.displayName : undefined,
      emailAddress: typeof args.emailAddress === 'string' ? args.emailAddress : undefined,
      members: Array.isArray(args.members) ? args.members.map(String) : undefined,
      settings: typeof args.settings === 'object' ? args.settings as DistributionListArgs['settings'] : undefined,
    };
  }

  private validateSecurityGroupArgs(args: Record<string, unknown>): SecurityGroupArgs {
    if (!args.action || typeof args.action !== 'string') {
      throw new McpError(ErrorCode.InvalidParams, 'Missing required security group parameters');
    }
    return {
      action: args.action as SecurityGroupArgs['action'],
      groupId: typeof args.groupId === 'string' ? args.groupId : undefined,
      displayName: typeof args.displayName === 'string' ? args.displayName : undefined,
      description: typeof args.description === 'string' ? args.description : undefined,
      members: Array.isArray(args.members) ? args.members.map(String) : undefined,
      settings: typeof args.settings === 'object' ? args.settings as SecurityGroupArgs['settings'] : undefined,
    };
  }

  private validateM365GroupArgs(args: Record<string, unknown>): M365GroupArgs {
    if (!args.action || typeof args.action !== 'string') {
      throw new McpError(ErrorCode.InvalidParams, 'Missing required M365 group parameters');
    }
    return {
      action: args.action as M365GroupArgs['action'],
      groupId: typeof args.groupId === 'string' ? args.groupId : undefined,
      displayName: typeof args.displayName === 'string' ? args.displayName : undefined,
      description: typeof args.description === 'string' ? args.description : undefined,
      owners: Array.isArray(args.owners) ? args.owners.map(String) : undefined,
      members: Array.isArray(args.members) ? args.members.map(String) : undefined,
      settings: typeof args.settings === 'object' ? args.settings as M365GroupArgs['settings'] : undefined,
    };
  }

  private validateExchangeSettingsArgs(args: Record<string, unknown>): ExchangeSettingsArgs {
    if (!args.action || !args.settingType || typeof args.action !== 'string' || typeof args.settingType !== 'string') {
      throw new McpError(ErrorCode.InvalidParams, 'Missing required Exchange settings parameters');
    }
    return {
      action: args.action as ExchangeSettingsArgs['action'],
      settingType: args.settingType as ExchangeSettingsArgs['settingType'],
      target: typeof args.target === 'string' ? args.target : undefined,
      settings: typeof args.settings === 'object' ? args.settings as ExchangeSettingsArgs['settings'] : undefined,
    };
  }

  private validateUserManagementArgs(args: Record<string, unknown>): UserManagementArgs {
    if (!args.action || !args.userId || typeof args.action !== 'string' || typeof args.userId !== 'string') {
      throw new McpError(ErrorCode.InvalidParams, 'Missing required user management parameters');
    }
    return {
      action: args.action as UserManagementArgs['action'],
      userPrincipalName: args.userId,
      settings: typeof args.settings === 'object' ? args.settings as Record<string, unknown> : undefined,
    };
  }

  private validateOffboardingArgs(args: Record<string, unknown>): OffboardingArgs {
    if (!args.action || !args.userId || typeof args.action !== 'string' || typeof args.userId !== 'string') {
      throw new McpError(ErrorCode.InvalidParams, 'Missing required offboarding parameters');
    }
    return {
      action: args.action as OffboardingArgs['action'],
      userId: args.userId,
      options: typeof args.options === 'object' ? args.options as OffboardingArgs['options'] : undefined,
    };
  }

  private async handleDistributionList(args: DistributionListArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    switch (args.action) {
      case 'get': {
        const list = await this.graphClient
          .api(`/groups/${args.listId}`)
          .get();
        return { content: [{ type: 'text', text: JSON.stringify(list, null, 2) }] };
      }
      case 'create': {
        const list = await this.graphClient
          .api('/groups')
          .post({
            displayName: args.displayName,
            mailEnabled: true,
            securityEnabled: false,
            mailNickname: args.emailAddress?.split('@')[0],
            ...args.settings,
          });
        return { content: [{ type: 'text', text: JSON.stringify(list, null, 2) }] };
      }
      case 'update': {
        await this.graphClient
          .api(`/groups/${args.listId}`)
          .patch({
            displayName: args.displayName,
            ...args.settings,
          });
        return { content: [{ type: 'text', text: 'Distribution list updated successfully' }] };
      }
      case 'delete': {
        await this.graphClient
          .api(`/groups/${args.listId}`)
          .delete();
        return { content: [{ type: 'text', text: 'Distribution list deleted successfully' }] };
      }
      case 'add_members': {
        if (!args.members?.length) {
          throw new McpError(ErrorCode.InvalidParams, 'No members specified to add');
        }
        for (const member of args.members) {
          await this.graphClient
            .api(`/groups/${args.listId}/members/$ref`)
            .post({
              '@odata.id': `https://graph.microsoft.com/v1.0/users/${member}`,
            });
        }
        return { content: [{ type: 'text', text: 'Members added successfully' }] };
      }
      case 'remove_members': {
        if (!args.members?.length) {
          throw new McpError(ErrorCode.InvalidParams, 'No members specified to remove');
        }
        for (const member of args.members) {
          await this.graphClient
            .api(`/groups/${args.listId}/members/${member}/$ref`)
            .delete();
        }
        return { content: [{ type: 'text', text: 'Members removed successfully' }] };
      }
      default:
        throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
    }
  }

  private async handleSecurityGroup(args: SecurityGroupArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    switch (args.action) {
      case 'get': {
        const group = await this.graphClient
          .api(`/groups/${args.groupId}`)
          .get();
        return { content: [{ type: 'text', text: JSON.stringify(group, null, 2) }] };
      }
      case 'create': {
        const group = await this.graphClient
          .api('/groups')
          .post({
            displayName: args.displayName,
            description: args.description,
            securityEnabled: true,
            mailEnabled: args.settings?.mailEnabled ?? false,
            mailNickname: args.displayName?.replace(/\s+/g, '').toLowerCase(),
          });
        return { content: [{ type: 'text', text: JSON.stringify(group, null, 2) }] };
      }
      case 'update': {
        await this.graphClient
          .api(`/groups/${args.groupId}`)
          .patch({
            displayName: args.displayName,
            description: args.description,
            ...args.settings,
          });
        return { content: [{ type: 'text', text: 'Security group updated successfully' }] };
      }
      case 'delete': {
        await this.graphClient
          .api(`/groups/${args.groupId}`)
          .delete();
        return { content: [{ type: 'text', text: 'Security group deleted successfully' }] };
      }
      case 'add_members':
      case 'remove_members': {
        if (!args.members?.length) {
          throw new McpError(ErrorCode.InvalidParams, 'No members specified');
        }
        for (const member of args.members) {
          if (args.action === 'add_members') {
            await this.graphClient
              .api(`/groups/${args.groupId}/members/$ref`)
              .post({
                '@odata.id': `https://graph.microsoft.com/v1.0/users/${member}`,
              });
          } else {
            await this.graphClient
              .api(`/groups/${args.groupId}/members/${member}/$ref`)
              .delete();
          }
        }
        return { content: [{ type: 'text', text: `Members ${args.action === 'add_members' ? 'added' : 'removed'} successfully` }] };
      }
      default:
        throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
    }
  }

  private async handleM365Group(args: M365GroupArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    switch (args.action) {
      case 'get': {
        const group = await this.graphClient
          .api(`/groups/${args.groupId}`)
          .get();
        return { content: [{ type: 'text', text: JSON.stringify(group, null, 2) }] };
      }
      case 'create': {
        const group = await this.graphClient
          .api('/groups')
          .post({
            displayName: args.displayName,
            description: args.description,
            groupTypes: ['Unified'],
            mailEnabled: true,
            securityEnabled: false,
            mailNickname: args.displayName?.replace(/\s+/g, '').toLowerCase(),
            visibility: args.settings?.visibility?.toLowerCase(),
            ...args.settings,
          });
        return { content: [{ type: 'text', text: JSON.stringify(group, null, 2) }] };
      }
      case 'update': {
        await this.graphClient
          .api(`/groups/${args.groupId}`)
          .patch({
            displayName: args.displayName,
            description: args.description,
            ...args.settings,
          });
        return { content: [{ type: 'text', text: 'M365 group updated successfully' }] };
      }
      case 'delete': {
        await this.graphClient
          .api(`/groups/${args.groupId}`)
          .delete();
        return { content: [{ type: 'text', text: 'M365 group deleted successfully' }] };
      }
      case 'add_members':
      case 'remove_members': {
        if (!args.members?.length) {
          throw new McpError(ErrorCode.InvalidParams, 'No members specified');
        }
        for (const member of args.members) {
          if (args.action === 'add_members') {
            await this.graphClient
              .api(`/groups/${args.groupId}/members/$ref`)
              .post({
                '@odata.id': `https://graph.microsoft.com/v1.0/users/${member}`,
              });
          } else {
            await this.graphClient
              .api(`/groups/${args.groupId}/members/${member}/$ref`)
              .delete();
          }
        }
        return { content: [{ type: 'text', text: `Members ${args.action === 'add_members' ? 'added' : 'removed'} successfully` }] };
      }
      default:
        throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
    }
  }

  private async handleExchangeSettings(args: ExchangeSettingsArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    switch (args.settingType) {
      case 'mailbox': {
        if (args.action === 'get') {
          const settings = await this.graphClient
            .api(`/users/${args.target}/mailboxSettings`)
            .get();
          return { content: [{ type: 'text', text: JSON.stringify(settings, null, 2) }] };
        } else {
          await this.graphClient
            .api(`/users/${args.target}/mailboxSettings`)
            .patch(args.settings);
          return { content: [{ type: 'text', text: 'Mailbox settings updated successfully' }] };
        }
      }
      case 'transport': {
        if (args.action === 'get') {
          const rules = await this.graphClient
            .api('/admin/transportRules')
            .get();
          return { content: [{ type: 'text', text: JSON.stringify(rules, null, 2) }] };
        } else {
          await this.graphClient
            .api('/admin/transportRules')
            .post(args.settings?.rules);
          return { content: [{ type: 'text', text: 'Transport rules updated successfully' }] };
        }
      }
      case 'organization': {
        if (args.action === 'get') {
          const settings = await this.graphClient
            .api('/admin/organization/settings')
            .get();
          return { content: [{ type: 'text', text: JSON.stringify(settings, null, 2) }] };
        } else {
          await this.graphClient
            .api('/admin/organization/settings')
            .patch(args.settings);
          return { content: [{ type: 'text', text: 'Organization settings updated successfully' }] };
        }
      }
      case 'retention': {
        if (args.action === 'get') {
          const policies = await this.graphClient
            .api('/admin/retentionPolicies')
            .get();
          return { content: [{ type: 'text', text: JSON.stringify(policies, null, 2) }] };
        } else {
          await this.graphClient
            .api('/admin/retentionPolicies')
            .post(args.settings?.retentionTags);
          return { content: [{ type: 'text', text: 'Retention policies updated successfully' }] };
        }
      }
      default:
        throw new McpError(ErrorCode.InvalidParams, `Invalid setting type: ${args.settingType}`);
    }
  }

  private async handleUserSettings(args: UserManagementArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    if (args.action === 'get') {
      const settings = await this.graphClient
        .api(`/users/${args.userPrincipalName}`)
        .get();
      return { content: [{ type: 'text', text: JSON.stringify(settings, null, 2) }] };
    } else {
      await this.graphClient
        .api(`/users/${args.userPrincipalName}`)
        .patch(args.settings);
      return { content: [{ type: 'text', text: 'User settings updated successfully' }] };
    }
  }

  private async handleOffboarding(args: OffboardingArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    switch (args.action) {
      case 'start': {
        // Block sign-ins
        await this.graphClient
          .api(`/users/${args.userId}`)
          .patch({ accountEnabled: false });

        if (args.options?.revokeAccess) {
          // Revoke all refresh tokens
          await this.graphClient
            .api(`/users/${args.userId}/revokeSignInSessions`)
            .post({});
        }

        if (args.options?.backupData) {
          // Trigger backup
          await this.graphClient
            .api(`/users/${args.userId}/drive/content`)
            .get();
        }

        return { content: [{ type: 'text', text: 'Offboarding process started successfully' }] };
      }
      case 'check': {
        const status = await this.graphClient
          .api(`/users/${args.userId}`)
          .get();
        return { content: [{ type: 'text', text: JSON.stringify(status, null, 2) }] };
      }
      case 'complete': {
        if (args.options?.convertToShared) {
          // Convert to shared mailbox
          await this.graphClient
            .api(`/users/${args.userId}/mailbox/convert`)
            .post({});
        } else if (!args.options?.retainMailbox) {
          // Delete user if not retaining mailbox
          await this.graphClient
            .api(`/users/${args.userId}`)
            .delete();
        }
        return { content: [{ type: 'text', text: 'Offboarding process completed successfully' }] };
      }
      default:
        throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
    }
  }

  async run(): Promise<void> {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('Microsoft 365 Core MCP server running on stdio');
  }
}

const server = new M365CoreServer();
server.run().catch(console.error);
