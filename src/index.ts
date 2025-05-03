#!/usr/bin/env node
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { isInitializeRequest } from '@modelcontextprotocol/sdk/types.js';
import express, { Request, Response } from 'express';
import { randomUUID } from 'node:crypto';
import { M365CoreServer } from './server.js';

// Environment validation
const PORT = process.env.PORT ? parseInt(process.env.PORT, 10) : 3000;
const LOG_LEVEL = process.env.LOG_LEVEL ?? 'info';
const USE_HTTP = process.env.USE_HTTP === 'true';
const STATELESS = process.env.STATELESS === 'true';

async function startServer() {
  const server = new M365CoreServer();

  if (USE_HTTP) {
    // Setup Express app for HTTP transport
    const app = express();
    app.use(express.json());

    if (STATELESS) {
      // Stateless mode - create a new instance for each request
      console.log('Running in stateless HTTP mode (no session management)');
      
      app.post('/mcp', async (req: Request, res: Response) => {
        try {
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
