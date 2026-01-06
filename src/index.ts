#!/usr/bin/env node
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { isInitializeRequest } from '@modelcontextprotocol/sdk/types.js';
import express, { Request, Response, NextFunction } from 'express';
import { randomUUID } from 'node:crypto';
import { M365CoreServer } from './server.js';

// Import authentication middleware (MCP SDK latest auth features)
import { mcpAuthMiddleware, optionalAuth, getAuthInfoFromRequest } from './auth/index.js';
import { getOAuthProvider, resetOAuthProvider } from './auth/index.js';

// DNS rebinding protection middleware (MCP SDK best practice for security)
// In production (Smithery), we trust the reverse proxy to handle host validation
// For local development, we validate allowed hosts
function hostHeaderValidation(allowedHosts: string[]) {
  return (req: Request, res: Response, next: NextFunction) => {
    // Skip validation in production environments (Smithery handles this via reverse proxy)
    if (process.env.NODE_ENV === 'production' || process.env.SMITHERY_DEPLOYMENT === 'true') {
      next();
      return;
    }
    
    const host = req.headers.host?.split(':')[0]; // Remove port if present
    
    // Allow Smithery deployment hosts
    if (host?.endsWith('.smithery.ai') || host?.endsWith('.smithery.app')) {
      next();
      return;
    }
    
    if (!host || allowedHosts.includes(host)) {
      next();
    } else {
      res.status(403).json({ error: 'Forbidden: Invalid host header' });
    }
  };
}

// Environment validation
// Default port 8080 for Smithery compatibility
const PORT = process.env.PORT ? parseInt(process.env.PORT, 10) : 8080;
const LOG_LEVEL = process.env.LOG_LEVEL ?? 'info';
const USE_HTTP = process.env.USE_HTTP === 'true';
const STATELESS = process.env.STATELESS === 'true';

// Configuration parsing for HTTP requests (Smithery integration)
interface M365Config {
  msTenantId?: string;
  msClientId?: string;
  msClientSecret?: string;
  stateless?: boolean;
  logLevel?: string;
}

function parseConfigFromRequest(req: Request): M365Config {
  const config: M365Config = {};
  
  // Parse from query parameters (Smithery passes config this way)
  if (req.query.msTenantId) config.msTenantId = req.query.msTenantId as string;
  if (req.query.msClientId) config.msClientId = req.query.msClientId as string;
  if (req.query.msClientSecret) config.msClientSecret = req.query.msClientSecret as string;
  if (req.query.stateless) config.stateless = req.query.stateless === 'true';
  if (req.query.logLevel) config.logLevel = req.query.logLevel as string;
  
  // Parse from headers (alternative method)
  if (req.headers['x-ms-tenant-id']) config.msTenantId = req.headers['x-ms-tenant-id'] as string;
  if (req.headers['x-ms-client-id']) config.msClientId = req.headers['x-ms-client-id'] as string;
  if (req.headers['x-ms-client-secret']) config.msClientSecret = req.headers['x-ms-client-secret'] as string;
  if (req.headers['x-stateless']) config.stateless = req.headers['x-stateless'] === 'true';
  if (req.headers['x-log-level']) config.logLevel = req.headers['x-log-level'] as string;
  
  // Parse from request body if it contains config
  if (req.body && typeof req.body === 'object' && req.body.config) {
    const bodyConfig = req.body.config;
    if (bodyConfig.msTenantId) config.msTenantId = bodyConfig.msTenantId;
    if (bodyConfig.msClientId) config.msClientId = bodyConfig.msClientId;
    if (bodyConfig.msClientSecret) config.msClientSecret = bodyConfig.msClientSecret;
    if (bodyConfig.stateless !== undefined) config.stateless = bodyConfig.stateless;
    if (bodyConfig.logLevel) config.logLevel = bodyConfig.logLevel;
  }
  
  return config;
}

async function startServer() {
  const server = new M365CoreServer();

  if (USE_HTTP) {
    // Setup Express app for HTTP transport
    const app = express();
    
    // DNS rebinding protection (MCP SDK best practice for security)
    // Allow localhost variants and container networking
    app.use(hostHeaderValidation(['localhost', '127.0.0.1', '0.0.0.0', 'host.docker.internal']));
    
    // CORS configuration for browser-based MCP clients
    app.use((req: Request, res: Response, next: NextFunction) => {
      res.header('Access-Control-Allow-Origin', '*');
      res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
      res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization, mcp-session-id, mcp-protocol-version');
      res.header('Access-Control-Expose-Headers', 'mcp-session-id, mcp-protocol-version');
      
      if (req.method === 'OPTIONS') {
        res.sendStatus(200);
        return;
      }
      next();
    });
    
    app.use(express.json());

    // ============================================
    // OAuth 2.0 Endpoints (MCP SDK latest auth features)
    // ============================================
    
    // OAuth authorization endpoint - redirect to Azure AD
    app.get('/oauth/authorize', (req: Request, res: Response) => {
      const provider = getOAuthProvider();
      const state = req.query.state as string || randomUUID();
      const scopes = (req.query.scope as string)?.split(' ');
      
      const authUrl = provider.getAuthorizationUrl(state, scopes);
      res.redirect(authUrl);
    });

    // OAuth callback endpoint - handle Azure AD response
    app.get('/oauth/callback', async (req: Request, res: Response) => {
      try {
        const code = req.query.code as string;
        const error = req.query.error as string;
        const state = req.query.state as string;

        if (error) {
          // Return error page for OAuth errors
          res.status(400).send(`
            <html>
              <body>
                <h1>Authorization Failed</h1>
                <p>Error: ${error}</p>
                <script>
                  if (window.opener) {
                    window.opener.postMessage({ type: 'oauth-error', error: '${error}' }, '*');
                    window.close();
                  }
                </script>
              </body>
            </html>
          `);
          return;
        }

        if (!code) {
          res.status(400).json({ error: 'Missing authorization code' });
          return;
        }

        // Exchange code for tokens
        const provider = getOAuthProvider();
        const tokens = await provider.exchangeCode(code);

        // Return success page with code for client-side handling
        res.send(`
          <html>
            <body>
              <h1>Authorization Successful!</h1>
              <p>You can close this window and return to the app.</p>
              <script>
                if (window.opener) {
                  window.opener.postMessage({ 
                    type: 'oauth-success', 
                    code: '${code}',
                    state: '${state || ''}'
                  }, '*');
                  window.close();
                } else {
                  // Fallback: store in session for retrieval
                  sessionStorage.setItem('oauth_code', '${code}');
                  window.location.href = '/?oauth=success';
                }
              </script>
            </body>
          </html>
        `);
      } catch (error) {
        console.error('OAuth callback error:', error);
        res.status(500).json({ 
          error: 'OAuth callback failed',
          message: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    });

    // Token endpoint - for token exchange and refresh
    app.post('/oauth/token', async (req: Request, res: Response) => {
      try {
        const { grant_type, code, refresh_token, code_verifier } = req.body;
        const provider = getOAuthProvider();

        let tokens;
        if (grant_type === 'authorization_code' && code) {
          tokens = await provider.exchangeCode(code, code_verifier);
        } else if (grant_type === 'refresh_token' && refresh_token) {
          tokens = await provider.refreshToken(refresh_token);
        } else {
          res.status(400).json({ error: 'invalid_grant', error_description: 'Unsupported grant type' });
          return;
        }

        res.json(tokens);
      } catch (error) {
        console.error('Token endpoint error:', error);
        res.status(500).json({ 
          error: 'server_error',
          error_description: error instanceof Error ? error.message : 'Token exchange failed'
        });
      }
    });

    // OAuth metadata endpoint (RFC 8414 - OAuth Server Metadata)
    app.get('/.well-known/oauth-authorization-server', (req: Request, res: Response) => {
      const provider = getOAuthProvider();
      const baseUrl = `${req.protocol}://${req.get('host')}`;
      
      res.json({
        issuer: baseUrl,
        authorization_endpoint: `${baseUrl}/oauth/authorize`,
        token_endpoint: `${baseUrl}/oauth/token`,
        token_endpoint_auth_methods_supported: ['client_secret_post', 'client_secret_basic'],
        grant_types_supported: ['authorization_code', 'refresh_token', 'client_credentials'],
        response_types_supported: ['code'],
        scopes_supported: [
          'openid', 'profile', 'email', 'offline_access',
          'User.Read', 'User.ReadWrite.All', 'Group.ReadWrite.All',
          'Directory.ReadWrite.All', 'Mail.ReadWrite', 'Sites.ReadWrite.All'
        ],
        code_challenge_methods_supported: ['S256', 'plain'],
        service_documentation: 'https://github.com/your-org/m365-core-mcp'
      });
    });

    // Apply optional auth middleware to MCP endpoints
    app.use('/mcp', optionalAuth);

    if (STATELESS) {
      // Stateless mode - create a new instance for each request
      console.log('Running in stateless HTTP mode (no session management)');
      
      app.post('/mcp', async (req: Request, res: Response) => {
        try {
          // Parse configuration from Smithery HTTP context
          const config = parseConfigFromRequest(req);
          
          // Set environment variables from config for this request
          if (config.msTenantId) process.env.MS_TENANT_ID = config.msTenantId;
          if (config.msClientId) process.env.MS_CLIENT_ID = config.msClientId;
          if (config.msClientSecret) process.env.MS_CLIENT_SECRET = config.msClientSecret;
          if (config.logLevel) process.env.LOG_LEVEL = config.logLevel;
          if (config.stateless !== undefined) process.env.STATELESS = config.stateless.toString();
          
          // In stateless mode, create a new instance of transport and server for each request
          const mcpServer = new M365CoreServer();
          const transport = new StreamableHTTPServerTransport({
            sessionIdGenerator: undefined, // No session ID in stateless mode
          });
          
          res.on('close', () => {
            console.log('Request closed');
            transport.close();
            mcpServer.server.close();
          });
          
          await mcpServer.server.connect(transport);
          await transport.handleRequest(req, res, req.body);
        } catch (error) {
          console.error('Error handling MCP request:', error);
          if (!res.headersSent) {
            res.status(500).json({
              jsonrpc: '2.0',
              error: {
                code: -32603,
                message: 'Internal server error',
              },
              id: null,
            });
          }
        }
      });

      // In stateless mode, GET and DELETE are not supported
      const unsupportedMethodHandler = (req: Request, res: Response) => {
        console.log(`Received ${req.method} MCP request in stateless mode`);
        res.status(405).json({
          jsonrpc: "2.0",
          error: {
            code: -32000,
            message: "Method not allowed in stateless mode."
          },
          id: null
        });
      };

      app.get('/mcp', unsupportedMethodHandler);
      app.delete('/mcp', unsupportedMethodHandler);
    } else {
      // Stateful mode with session management
      console.log('Running in stateful HTTP mode (with session management)');
      
      // Map to store transports by session ID
      const transports: { [sessionId: string]: StreamableHTTPServerTransport } = {};

      // Handle POST requests for client-to-server communication
      app.post('/mcp', async (req: Request, res: Response) => {
        try {
          // Parse configuration from Smithery HTTP context for every request
          const config = parseConfigFromRequest(req);
          
          // Set environment variables from config for this request
          if (config.msTenantId) process.env.MS_TENANT_ID = config.msTenantId;
          if (config.msClientId) process.env.MS_CLIENT_ID = config.msClientId;
          if (config.msClientSecret) process.env.MS_CLIENT_SECRET = config.msClientSecret;
          if (config.logLevel) process.env.LOG_LEVEL = config.logLevel;
          if (config.stateless !== undefined) process.env.STATELESS = config.stateless.toString();
          
          // Check for existing session ID
          const sessionId = req.headers['mcp-session-id'] as string | undefined;
          let transport: StreamableHTTPServerTransport;

          if (sessionId && transports[sessionId]) {
            // Reuse existing transport
            transport = transports[sessionId];
            console.log(`Using existing session: ${sessionId}`);
          } else if (!sessionId && isInitializeRequest(req.body)) {
            // New initialization request
            transport = new StreamableHTTPServerTransport({
              sessionIdGenerator: () => randomUUID(),
              onsessioninitialized: (newSessionId: string) => {
                // Store the transport by session ID
                if (newSessionId) {
                  transports[newSessionId] = transport;
                  console.log(`New session initialized: ${newSessionId}`);
                }
              }
            });

            // Clean up transport when closed
            transport.onclose = () => {
              if (transport.sessionId) {
                delete transports[transport.sessionId];
                console.log(`Session closed: ${transport.sessionId}`);
              }
            };

            // Connect to the MCP server
            await server.server.connect(transport);
          } else {
            // Invalid request
            res.status(400).json({
              jsonrpc: '2.0',
              error: {
                code: -32000,
                message: 'Bad Request: No valid session ID provided',
              },
              id: null,
            });
            return;
          }

          // Handle the request
          await transport.handleRequest(req, res, req.body);
        } catch (error) {
          console.error('Error handling MCP request:', error);
          if (!res.headersSent) {
            res.status(500).json({
              jsonrpc: '2.0',
              error: {
                code: -32603,
                message: 'Internal server error',
              },
              id: null,
            });
          }
        }
      });

      // Reusable handler for GET and DELETE requests
      const handleSessionRequest = async (req: Request, res: Response) => {
        const sessionId = req.headers['mcp-session-id'] as string | undefined;
        if (!sessionId || !transports[sessionId]) {
          res.status(400).json({
            jsonrpc: '2.0',
            error: {
              code: -32000,
              message: 'Invalid or missing session ID',
            },
            id: null,
          });
          return;
        }
        
        const transport = transports[sessionId];
        await transport.handleRequest(req, res);
      };

      // Handle GET requests for server-to-client notifications
      app.get('/mcp', handleSessionRequest);

      // Handle DELETE requests for session termination
      app.delete('/mcp', handleSessionRequest);
    }

    // Add SSE endpoint for real-time updates
    app.get('/sse', (req: Request, res: Response) => {
      // Set headers for SSE
      res.writeHead(200, {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Cache-Control'
      });

      // Add client to server's SSE clients
      server.addSSEClient(res);

      // Send initial connection event
      res.write(`data: ${JSON.stringify({
        type: 'connected',
        message: 'Connected to M365 Core MCP Server SSE stream',
        timestamp: new Date().toISOString()
      })}\n\n`);

      // Handle client disconnect
      req.on('close', () => {
        server.removeSSEClient(res);
      });

      req.on('aborted', () => {
        server.removeSSEClient(res);
      });
    });    // Health check endpoint with full capabilities report
    app.get('/health', (req: Request, res: Response) => {
      res.json({
        status: 'healthy',
        server: 'M365 Core MCP Server',
        version: '1.1.0', // Enhanced version with improved API capabilities
        capabilities: {
          tools: true,
          resources: true,
          prompts: true,
          sse: true,
          progressReporting: true,
          resourceSubscriptions: true,
          streamingResponses: true,
          lazyLoading: true
        },
        features: {
          'Microsoft Graph API': 'Full access to M365 services',
          'Azure AD Management': 'Users, groups, roles, devices, apps',
          'Exchange Online': 'Mailbox and transport settings',
          'SharePoint': 'Sites and lists management',
          'Security & Compliance': 'Audit logs, alerts, compliance',
          'Real-time Updates': 'SSE for live notifications',
          'Progress Tracking': 'Long-running operation status',
          'Resource Subscriptions': 'Live resource change notifications'
        },
        sseClients: server.sseClients?.size || 0,
        activeOperations: server.progressTrackers?.size || 0,
        timestamp: new Date().toISOString()
      });
    });    // Capabilities endpoint for MCP clients
    app.get('/capabilities', (req: Request, res: Response) => {
      res.json({
        name: 'm365-core-server',
        version: '1.1.0', // Enhanced version with improved API capabilities
        protocol: 'mcp',
        capabilities: {
          tools: {
            listChanged: true
          },
          resources: {
            subscribe: true,
            listChanged: true
          },
          prompts: {
            listChanged: true
          },
          experimental: {
            progressReporting: true,
            streamingResponses: true
          }
        },
        tools: [
          'manage_distribution_lists',
          'manage_security_groups', 
          'manage_m365_groups',
          'manage_exchange_settings',
          'manage_user_settings',
          'manage_offboarding',
          'manage_sharepoint_sites',
          'manage_sharepoint_lists',
          'manage_azure_ad_roles',
          'manage_azure_ad_apps',
          'manage_azure_ad_devices',
          'manage_service_principals',
          'dynamicendpoints_m365_assistant',
          'search_audit_log',
          'manage_alerts'
        ],
        resources: [
          'current_user',
          'organization_info',
          'security_score',
          'recent_audit_logs'
        ],
        prompts: [
          'create_security_group',
          'setup_compliance_policy',
          'analyze_security_score',
          'audit_user_activity'
        ]
      });
    });

    // Start the HTTP server
    app.listen(PORT, () => {
      console.log(`M365 Core MCP Server running on HTTP at port ${PORT}`);
      console.log(`Server URL: http://localhost:${PORT}/mcp`);
      console.log(`Mode: ${STATELESS ? 'Stateless' : 'Stateful with session management'}`);
    });
  } else {
    // Use stdio transport
    const transport = new StdioServerTransport();
    await server.server.connect(transport);
    console.log('M365 Core MCP Server running on stdio');
  }
}

// Start the server
startServer().catch(error => {
  console.error('Failed to start server:', error);
  process.exit(1);
});
