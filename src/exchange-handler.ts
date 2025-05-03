import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { ExchangeSettingsArgs } from './types.js';

export async function handleExchangeSettings(
  graphClient: Client,
  args: ExchangeSettingsArgs
): Promise<{ content: { type: string; text: string; }[]; }> {
  switch (args.settingType) {
    case 'mailbox': {
      if (args.action === 'get') {
        const settings = await graphClient
          .api(`/users/${args.target}/mailboxSettings`)
          .get();
        return { content: [{ type: 'text', text: JSON.stringify(settings, null, 2) }] };
      } else {
        await graphClient
          .api(`/users/${args.target}/mailboxSettings`)
          .patch(args.settings);
        return { content: [{ type: 'text', text: 'Mailbox settings updated successfully' }] };
      }
    }
    case 'transport': {
      if (args.action === 'get') {
        const rules = await graphClient
          .api('/admin/transportRules')
          .get();
        return { content: [{ type: 'text', text: JSON.stringify(rules, null, 2) }] };
      } else {
        await graphClient
          .api('/admin/transportRules')
          .post(args.settings?.rules);
        return { content: [{ type: 'text', text: 'Transport rules updated successfully' }] };
      }
    }
    case 'organization': {
      if (args.action === 'get') {
        const settings = await graphClient
          .api('/admin/organization/settings')
          .get();
        return { content: [{ type: 'text', text: JSON.stringify(settings, null, 2) }] };
      } else {
        await graphClient
          .api('/admin/organization/settings')
          .patch(args.settings?.sharingPolicy);
        return { content: [{ type: 'text', text: 'Organization settings updated successfully' }] };
      }
    }
    case 'retention': {
      if (args.action === 'get') {
        const tags = await graphClient
          .api('/admin/retentionTags')
          .get();
        return { content: [{ type: 'text', text: JSON.stringify(tags, null, 2) }] };
      } else {
        if (!args.settings?.retentionTags?.length) {
          throw new McpError(ErrorCode.InvalidParams, 'No retention tags specified');
        }
        for (const tag of args.settings.retentionTags) {
          await graphClient
            .api('/admin/retentionTags')
            .post(tag);
        }
        return { content: [{ type: 'text', text: 'Retention tags updated successfully' }] };
      }
    }
    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid setting type: ${args.settingType}`);
  }
}
