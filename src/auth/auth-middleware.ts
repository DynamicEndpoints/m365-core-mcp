/**
 * Authentication Middleware for MCP HTTP Server
 * 
 * Implements MCP SDK authentication best practices:
 * - Bearer token extraction and validation
 * - AuthInfo propagation to tool handlers
 * - Support for both OAuth and API key authentication
 * 
 * @see https://modelcontextprotocol.io/docs/concepts/authentication
 */

import { Request, Response, NextFunction } from 'express';
import { getOAuthProvider, AuthInfo } from './oauth-provider.js';

// Extend Express Request to include authInfo
declare global {
  namespace Express {
    interface Request {
      authInfo?: AuthInfo;
      mcpSessionId?: string;
    }
  }
}

/**
 * Extract bearer token from Authorization header
 */
function extractBearerToken(authHeader: string | undefined): string | null {
  if (!authHeader) return null;
  
  const parts = authHeader.split(' ');
  if (parts.length !== 2 || parts[0].toLowerCase() !== 'bearer') {
    return null;
  }
  
  return parts[1];
}

/**
 * Authentication middleware for MCP HTTP endpoints
 * Validates tokens and attaches AuthInfo to request
 */
export function mcpAuthMiddleware(options: {
  required?: boolean;
  allowApiKey?: boolean;
} = {}) {
  const { required = false, allowApiKey = true } = options;

  return async (req: Request, res: Response, next: NextFunction) => {
    try {
      // Extract session ID if present
      req.mcpSessionId = req.headers['mcp-session-id'] as string | undefined;

      // Try Bearer token first
      const bearerToken = extractBearerToken(req.headers.authorization);
      
      if (bearerToken) {
        const provider = getOAuthProvider();
        const tokenInfo = await provider.verifyAccessToken(bearerToken);
        
        if (tokenInfo) {
          req.authInfo = {
            token: tokenInfo.token,
            clientId: tokenInfo.clientId,
            scopes: tokenInfo.scopes,
            userId: tokenInfo.userId,
            tenantId: tokenInfo.tenantId
          };
          return next();
        }
      }

      // Try API key from query params (Smithery pattern)
      if (allowApiKey) {
        const apiKey = req.query.api_key as string | undefined;
        if (apiKey) {
          // API key is typically the Smithery user's key - allow pass-through
          req.authInfo = {
            token: apiKey,
            clientId: 'smithery-api-key',
            scopes: ['*']
          };
          return next();
        }
      }

      // Check for Microsoft credentials in query params (Smithery config)
      const msTenantId = req.query.msTenantId as string;
      const msClientId = req.query.msClientId as string;
      const msClientSecret = req.query.msClientSecret as string;
      
      if (msTenantId && msClientId && msClientSecret) {
        // Set environment variables for this request context
        process.env.MS_TENANT_ID = msTenantId;
        process.env.MS_CLIENT_ID = msClientId;
        process.env.MS_CLIENT_SECRET = msClientSecret;
        
        req.authInfo = {
          token: 'client-credentials',
          clientId: msClientId,
          scopes: ['https://graph.microsoft.com/.default'],
          tenantId: msTenantId
        };
        return next();
      }

      // No authentication found
      if (required) {
        return res.status(401).json({
          jsonrpc: '2.0',
          error: {
            code: -32001,
            message: 'Authentication required',
            data: {
              supportedMethods: ['Bearer token', 'API key', 'Microsoft credentials']
            }
          },
          id: null
        });
      }

      // Continue without auth for optional endpoints
      next();
    } catch (error) {
      console.error('Auth middleware error:', error);
      next(error);
    }
  };
}

/**
 * Middleware to require authentication
 */
export const requireAuth = mcpAuthMiddleware({ required: true });

/**
 * Middleware for optional authentication
 */
export const optionalAuth = mcpAuthMiddleware({ required: false });

/**
 * Extract AuthInfo from request for passing to tool handlers
 */
export function getAuthInfoFromRequest(req: Request): AuthInfo | undefined {
  return req.authInfo;
}
