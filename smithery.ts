/**
 * Smithery TypeScript configuration for M365 Core MCP Server
 * https://smithery.ai/docs/build/deployments/typescript
 */

import { z } from 'zod';

// Configuration schema for Smithery
export const configSchema = z.object({
  msTenantId: z.string()
    .min(1, "Tenant ID is required")
    .describe('Microsoft Entra ID (Azure AD) Tenant ID - Found in Azure Portal > Entra ID > Overview'),
  msClientId: z.string()
    .min(1, "Client ID is required")
    .describe('Application (Client) ID - Found in Azure Portal > App Registration > Overview'),
  msClientSecret: z.string()
    .min(1, "Client Secret is required")
    .describe('Client Secret Value - Created in Azure Portal > App Registration > Certificates & secrets'),
  msRedirectUri: z.string()
    .optional()
    .describe('OAuth Redirect URI for user-delegated auth (default: http://localhost:3000/auth/callback)'),
  logLevel: z.enum(['debug', 'info', 'warn', 'error'])
    .optional()
    .describe('Log verbosity level (default: info)')
});

// Server factory function - default export required by Smithery
export default function createServer({ config }: { config: z.infer<typeof configSchema> }) {
  // Set environment variables from config
  process.env.MS_TENANT_ID = config.msTenantId;
  process.env.MS_CLIENT_ID = config.msClientId;
  process.env.MS_CLIENT_SECRET = config.msClientSecret;
  process.env.MS_REDIRECT_URI = config.msRedirectUri || 'http://localhost:3000/auth/callback';
  process.env.LOG_LEVEL = config.logLevel || 'info';

  // Dynamically require server to avoid side effects at module load
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  const { M365CoreServer } = require('./src/server.js');
  const server = new M365CoreServer();
  
  return server.server;
}
