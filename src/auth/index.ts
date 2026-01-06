/**
 * Authentication Module Exports
 * 
 * Central export point for all authentication-related functionality
 */

export { 
  M365OAuthProvider, 
  getOAuthProvider, 
  resetOAuthProvider,
  DEFAULT_M365_SCOPES,
  DELEGATED_SCOPES 
} from './oauth-provider.js';

export type { 
  M365OAuthConfig, 
  TokenInfo, 
  AuthInfo 
} from './oauth-provider.js';

export { 
  mcpAuthMiddleware, 
  requireAuth, 
  optionalAuth,
  getAuthInfoFromRequest 
} from './auth-middleware.js';
