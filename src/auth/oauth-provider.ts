/**
 * MCP OAuth Provider Implementation
 * 
 * Implements OAuth 2.0 for MCP SDK integration per latest best practices:
 * - ProxyOAuthServerProvider pattern for Azure AD integration
 * - Token validation and refresh
 * - Smithery OAuthProvider compatibility
 * 
 * @see https://modelcontextprotocol.io/docs/concepts/authentication
 */

import type { OAuthClientInformation } from '@modelcontextprotocol/sdk/shared/auth.js';

// Types for OAuth flow
export interface M365OAuthConfig {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  redirectUri?: string;
  scopes?: string[];
}

export interface TokenInfo {
  token: string;
  clientId: string;
  scopes: string[];
  expiresAt?: number;
  userId?: string;
  tenantId?: string;
}

export interface AuthInfo {
  token: string;
  clientId: string;
  scopes: string[];
  userId?: string;
  tenantId?: string;
}

// Default scopes for Microsoft 365 operations
export const DEFAULT_M365_SCOPES = [
  'https://graph.microsoft.com/.default'
];

// User consent scopes for delegated access
export const DELEGATED_SCOPES = [
  'User.Read',
  'User.ReadWrite.All',
  'Group.ReadWrite.All',
  'Directory.ReadWrite.All',
  'Mail.ReadWrite',
  'Sites.ReadWrite.All',
  'Files.ReadWrite.All',
  'offline_access'
];

/**
 * Microsoft 365 OAuth Provider for MCP servers
 * Supports both client credentials (app-only) and delegated (user) flows
 */
export class M365OAuthProvider {
  private config: M365OAuthConfig;
  private tokenCache: Map<string, { token: string; expiresAt: number }> = new Map();
  private registeredClients: Map<string, OAuthClientInformation> = new Map();

  constructor(config?: Partial<M365OAuthConfig>) {
    this.config = {
      tenantId: config?.tenantId || process.env.MS_TENANT_ID || '',
      clientId: config?.clientId || process.env.MS_CLIENT_ID || '',
      clientSecret: config?.clientSecret || process.env.MS_CLIENT_SECRET || '',
      redirectUri: config?.redirectUri || process.env.MS_REDIRECT_URI || 'http://localhost:8080/oauth/callback',
      scopes: config?.scopes || DEFAULT_M365_SCOPES
    };
  }

  /**
   * Get OAuth 2.0 endpoints for Azure AD
   */
  get endpoints() {
    return {
      authorizationUrl: `https://login.microsoftonline.com/${this.config.tenantId}/oauth2/v2.0/authorize`,
      tokenUrl: `https://login.microsoftonline.com/${this.config.tenantId}/oauth2/v2.0/token`,
      revocationUrl: `https://login.microsoftonline.com/${this.config.tenantId}/oauth2/v2.0/logout`,
      userInfoUrl: 'https://graph.microsoft.com/v1.0/me'
    };
  }

  /**
   * Verify an access token and return token info
   * This is called by MCP SDK to validate incoming tokens
   */
  async verifyAccessToken(token: string): Promise<TokenInfo | null> {
    try {
      // Validate token by calling Microsoft Graph /me endpoint
      const response = await fetch(this.endpoints.userInfoUrl, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json'
        }
      });

      if (!response.ok) {
        console.error('Token validation failed:', response.status);
        return null;
      }

      const userData = await response.json();
      
      return {
        token,
        clientId: this.config.clientId,
        scopes: DELEGATED_SCOPES,
        userId: userData.id,
        tenantId: this.config.tenantId
      };
    } catch (error) {
      console.error('Token verification error:', error);
      return null;
    }
  }

  /**
   * Get registered OAuth client information
   */
  async getClient(clientId: string): Promise<OAuthClientInformation | null> {
    // Return pre-registered client or validate against known clients
    if (this.registeredClients.has(clientId)) {
      return this.registeredClients.get(clientId)!;
    }

    // For Smithery deployment, accept the configured client
    if (clientId === this.config.clientId) {
      return {
        client_id: this.config.clientId,
        redirect_uris: [this.config.redirectUri!],
        grant_types: ['authorization_code', 'refresh_token', 'client_credentials'],
        response_types: ['code'],
        scope: DELEGATED_SCOPES.join(' '),
        token_endpoint_auth_method: 'client_secret_post'
      } as OAuthClientInformation;
    }

    return null;
  }

  /**
   * Register a new OAuth client dynamically
   */
  registerClient(clientInfo: OAuthClientInformation): void {
    this.registeredClients.set(clientInfo.client_id, clientInfo);
  }

  /**
   * Generate authorization URL for user consent flow
   */
  getAuthorizationUrl(state: string, scopes?: string[]): string {
    const url = new URL(this.endpoints.authorizationUrl);
    url.searchParams.set('client_id', this.config.clientId);
    url.searchParams.set('response_type', 'code');
    url.searchParams.set('redirect_uri', this.config.redirectUri!);
    url.searchParams.set('scope', (scopes || DELEGATED_SCOPES).join(' '));
    url.searchParams.set('state', state);
    url.searchParams.set('response_mode', 'query');
    return url.toString();
  }

  /**
   * Exchange authorization code for tokens
   */
  async exchangeCode(code: string, codeVerifier?: string): Promise<{
    access_token: string;
    refresh_token?: string;
    expires_in: number;
    scope: string;
  }> {
    const params = new URLSearchParams();
    params.set('client_id', this.config.clientId);
    params.set('client_secret', this.config.clientSecret);
    params.set('code', code);
    params.set('redirect_uri', this.config.redirectUri!);
    params.set('grant_type', 'authorization_code');
    
    if (codeVerifier) {
      params.set('code_verifier', codeVerifier);
    }

    const response = await fetch(this.endpoints.tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params
    });

    if (!response.ok) {
      const error = await response.text();
      throw new Error(`Token exchange failed: ${error}`);
    }

    return response.json();
  }

  /**
   * Refresh an access token using refresh token
   */
  async refreshToken(refreshToken: string): Promise<{
    access_token: string;
    refresh_token?: string;
    expires_in: number;
  }> {
    const params = new URLSearchParams();
    params.set('client_id', this.config.clientId);
    params.set('client_secret', this.config.clientSecret);
    params.set('refresh_token', refreshToken);
    params.set('grant_type', 'refresh_token');

    const response = await fetch(this.endpoints.tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params
    });

    if (!response.ok) {
      const error = await response.text();
      throw new Error(`Token refresh failed: ${error}`);
    }

    return response.json();
  }

  /**
   * Get access token using client credentials (app-only)
   */
  async getClientCredentialsToken(scope: string = 'https://graph.microsoft.com/.default'): Promise<string> {
    const cacheKey = `client_credentials:${scope}`;
    const cached = this.tokenCache.get(cacheKey);
    const now = Date.now();

    // Return cached token if still valid (with 60s buffer)
    if (cached && cached.expiresAt > now + 60000) {
      return cached.token;
    }

    const params = new URLSearchParams();
    params.set('client_id', this.config.clientId);
    params.set('client_secret', this.config.clientSecret);
    params.set('scope', scope);
    params.set('grant_type', 'client_credentials');

    const response = await fetch(this.endpoints.tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params
    });

    if (!response.ok) {
      const error = await response.text();
      throw new Error(`Client credentials token acquisition failed: ${error}`);
    }

    const data = await response.json();
    
    // Cache the token
    this.tokenCache.set(cacheKey, {
      token: data.access_token,
      expiresAt: now + (data.expires_in * 1000)
    });

    return data.access_token;
  }

  /**
   * Validate configuration
   */
  isConfigured(): boolean {
    return !!(this.config.tenantId && this.config.clientId && this.config.clientSecret);
  }

  /**
   * Get configuration status
   */
  getConfigStatus(): { configured: boolean; missing: string[] } {
    const missing: string[] = [];
    if (!this.config.tenantId) missing.push('MS_TENANT_ID');
    if (!this.config.clientId) missing.push('MS_CLIENT_ID');
    if (!this.config.clientSecret) missing.push('MS_CLIENT_SECRET');
    
    return {
      configured: missing.length === 0,
      missing
    };
  }
}

// Singleton instance for server-wide use
let providerInstance: M365OAuthProvider | null = null;

export function getOAuthProvider(config?: Partial<M365OAuthConfig>): M365OAuthProvider {
  if (!providerInstance) {
    providerInstance = new M365OAuthProvider(config);
  }
  return providerInstance;
}

export function resetOAuthProvider(): void {
  providerInstance = null;
}
