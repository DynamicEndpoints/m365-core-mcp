import { Client } from '@microsoft/microsoft-graph-client';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { randomUUID } from 'crypto';

/**
 * Modern Microsoft Graph API Utilities
 * Provides enhanced error handling, retry logic, and performance optimizations
 * 
 * Enhanced to support multiple resource endpoints for Intune-specific operations
 */

export interface GraphApiOptions {
  method?: 'GET' | 'POST' | 'PATCH' | 'DELETE' | 'PUT';
  body?: any;
  select?: string[];
  filter?: string;
  expand?: string;
  top?: number;
  skip?: number;
  orderBy?: string;
  maxRetries?: number;
  headers?: Record<string, string>;
}

export interface GraphApiResponse<T = any> {
  data: T;
  requestId: string;
  duration: number;
}

/**
 * Enhanced Graph API client with modern best practices
 */
export class ModernGraphClient {
  private client: Client;
  private defaultHeaders: Record<string, string>;
  private resource: string;

  /**
   * Create a new ModernGraphClient instance
   * @param client - The Microsoft Graph client instance
   * @param resource - The resource endpoint (default: 'https://graph.microsoft.com')
   */
  constructor(client: Client, resource: string = 'https://graph.microsoft.com') {
    this.client = client;
    this.resource = resource;
    this.defaultHeaders = {
      'User-Agent': 'M365-Core-MCP/1.0',
      'Prefer': 'return=minimal', // Optimize response size
      'Content-Type': 'application/json'
    };
  }

  /**
   * Get the resource endpoint for this client
   */
  getResource(): string {
    return this.resource;
  }

  /**
   * Make an optimized Graph API call with automatic retry and error handling
   */
  async makeApiCall<T>(endpoint: string, options: GraphApiOptions = {}): Promise<GraphApiResponse<T>> {
    const {
      method = 'GET',
      body,
      select,
      filter,
      expand,
      top,
      skip,
      orderBy,
      maxRetries = 3,
      headers = {}
    } = options;

    const startTime = Date.now();
    const requestId = randomUUID();
    
    // Merge headers
    const requestHeaders = {
      ...this.defaultHeaders,
      'client-request-id': requestId,
      'x-ms-client-resource': this.resource,
      ...headers
    };

    // Log resource usage for debugging
    console.log(`📡 Request to ${endpoint} using resource: ${this.resource}`);

    let apiCall = this.client.api(endpoint);

    // Apply query parameters for optimization
    if (select && select.length > 0) {
      apiCall = apiCall.select(select.join(','));
    }
    if (filter) {
      apiCall = apiCall.filter(filter);
    }
    if (expand) {
      apiCall = apiCall.expand(expand);
    }
    if (top) {
      apiCall = apiCall.top(top);
    }
    if (skip) {
      apiCall = apiCall.skip(skip);
    }
    if (orderBy) {
      apiCall = apiCall.orderby(orderBy);
    }

    // Add headers
    Object.entries(requestHeaders).forEach(([key, value]) => {
      apiCall = apiCall.header(key, value);
    });

    // Implement retry logic with exponential backoff
    for (let attempt = 0; attempt < maxRetries; attempt++) {
      try {
        let result: T;
        
        switch (method) {
          case 'GET':
            result = await apiCall.get();
            break;
          case 'POST':
            result = await apiCall.post(body);
            break;
          case 'PATCH':
            result = await apiCall.patch(body);
            break;
          case 'PUT':
            result = await apiCall.put(body);
            break;
          case 'DELETE':
            result = await apiCall.delete();
            break;
          default:
            throw new Error(`Unsupported HTTP method: ${method}`);
        }

        const duration = Date.now() - startTime;
        
        // Log successful API call
        console.log(`✅ Graph API Success: ${method} ${endpoint}`, {
          requestId,
          duration: `${duration}ms`,
          attempt: attempt + 1,
          hasSelect: !!select,
          hasFilter: !!filter
        });

        return { data: result, requestId, duration };

      } catch (error: any) {
        const isLastAttempt = attempt === maxRetries - 1;
        const duration = Date.now() - startTime;
        
        // Handle throttling (429) with retry-after
        if (error.status === 429 && !isLastAttempt) {
          const retryAfter = parseInt(error.headers?.['retry-after'] || '1');
          const backoffDelay = Math.min(retryAfter * 1000, Math.pow(2, attempt) * 1000);
          
          console.warn(`⚠️ Graph API Throttled: ${method} ${endpoint}`, {
            requestId,
            retryAfter: `${backoffDelay}ms`,
            attempt: attempt + 1,
            duration: `${duration}ms`
          });
          
          await this.delay(backoffDelay);
          continue;
        }
        
        // Handle transient server errors (5xx)
        if (error.status >= 500 && error.status < 600 && !isLastAttempt) {
          const backoffDelay = Math.pow(2, attempt) * 1000;
          
          console.warn(`⚠️ Graph API Server Error: ${method} ${endpoint}`, {
            requestId,
            status: error.status,
            retryIn: `${backoffDelay}ms`,
            attempt: attempt + 1,
            duration: `${duration}ms`
          });
          
          await this.delay(backoffDelay);
          continue;
        }

        // Log final error
        console.error(`❌ Graph API Error: ${method} ${endpoint}`, {
          requestId,
          status: error.status,
          message: error.message,
          attempt: attempt + 1,
          duration: `${duration}ms`,
          responseHeaders: error.headers
        });
        
        // Transform error to MCP format
        throw this.transformError(error, method, endpoint, requestId);
      }
    }

    // This should never be reached
    throw new McpError(ErrorCode.InternalError, 'Maximum retry attempts exceeded');
  }

  /**
   * Get paginated results with automatic page handling
   */
  async *getPaginatedResults<T>(
    endpoint: string, 
    options: GraphApiOptions = {}
  ): AsyncGenerator<T[], void, unknown> {
    let nextUrl: string | undefined = endpoint;
    let pageCount = 0;
    const maxPages = 100; // Safety limit

    while (nextUrl && pageCount < maxPages) {
      try {
        const response: GraphApiResponse<{ value: T[]; '@odata.nextLink'?: string }> = await this.makeApiCall<{ value: T[]; '@odata.nextLink'?: string }>(
          nextUrl,
          { ...options, maxRetries: 2 } // Reduce retries for pagination
        );

        if (response.data.value && response.data.value.length > 0) {
          yield response.data.value;
        }

        nextUrl = response.data['@odata.nextLink'];
        pageCount++;

        // Log pagination progress
        if (nextUrl) {
          console.log(`📄 Paginated API: ${endpoint} - Page ${pageCount + 1} available`);
        }

      } catch (error) {
        console.error(`❌ Pagination error on page ${pageCount + 1}:`, error);
        break;
      }
    }

    if (pageCount >= maxPages) {
      console.warn(`⚠️ Pagination limit reached for ${endpoint} (${maxPages} pages)`);
    }
  }

  /**
   * Batch multiple API calls for efficiency
   */
  async batchRequests<T>(requests: Array<{
    id: string;
    method: string;
    url: string;
    body?: any;
    headers?: Record<string, string>;
  }>): Promise<Array<{ id: string; status: number; body: T }>> {
    const batchPayload = {
      requests: requests.map(req => ({
        id: req.id,
        method: req.method,
        url: req.url,
        body: req.body,
        headers: req.headers || {}
      }))
    };

    const response = await this.makeApiCall<{
      responses: Array<{ id: string; status: number; body: T }>
    }>('$batch', {
      method: 'POST',
      body: batchPayload
    });

    return response.data.responses;
  }

  /**
   * Create a delay for retry logic
   */
  private delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Transform Graph API errors to MCP errors
   */
  private transformError(error: any, method: string, endpoint: string, requestId: string): McpError {
    let errorCode: ErrorCode;
    let message: string;

    switch (error.status) {
      case 400:
        errorCode = ErrorCode.InvalidParams;
        message = `Bad Request: ${error.message}`;
        break;
      case 401:
        errorCode = ErrorCode.InvalidParams;
        message = `Authentication failed: ${error.message}`;
        break;
      case 403:
        errorCode = ErrorCode.InvalidParams;
        message = `Access denied: ${error.message}`;
        break;
      case 404:
        errorCode = ErrorCode.InvalidParams;
        message = `Resource not found: ${error.message}`;
        break;
      case 429:
        errorCode = ErrorCode.InternalError;
        message = `Rate limited: ${error.message}`;
        break;
      case 500:
      case 502:
      case 503:
      case 504:
        errorCode = ErrorCode.InternalError;
        message = `Server error: ${error.message}`;
        break;
      default:
        errorCode = ErrorCode.InternalError;
        message = `Graph API error: ${error.message}`;
    }

    return new McpError(
      errorCode,
      `${message} (${method} ${endpoint}, Request ID: ${requestId})`
    );
  }
}

/**
 * Common Graph API query patterns for optimization
 */
export const GraphQueries = {
  // User queries
  users: {
    basic: ['id', 'displayName', 'userPrincipalName', 'accountEnabled'],
    detailed: ['id', 'displayName', 'userPrincipalName', 'accountEnabled', 'mail', 'jobTitle', 'department', 'lastSignInDateTime'],
    security: ['id', 'displayName', 'userPrincipalName', 'accountEnabled', 'signInActivity', 'riskLevel']
  },

  // Device queries
  devices: {
    basic: ['id', 'displayName', 'operatingSystem', 'operatingSystemVersion'],
    compliance: ['id', 'deviceName', 'operatingSystem', 'complianceState', 'lastSyncDateTime', 'enrolledDateTime'],
    security: ['id', 'deviceName', 'operatingSystem', 'trustType', 'isCompliant', 'managementType']
  },

  // Security queries
  security: {
    alerts: ['id', 'displayName', 'severity', 'status', 'createdDateTime', 'classification'],
    incidents: ['id', 'displayName', 'status', 'severity', 'createdDateTime', 'lastUpdateDateTime'],
    compliance: ['id', 'displayName', 'state', 'policyType', 'platform']
  },

  // Group queries
  groups: {
    basic: ['id', 'displayName', 'groupTypes', 'securityEnabled'],
    membership: ['id', 'displayName', 'groupTypes', 'membershipRule', 'membershipRuleProcessingState'],
    teams: ['id', 'displayName', 'resourceProvisioningOptions', 'visibility']
  }
};

/**
 * Utility functions for Intune-specific operations
 */

/**
 * Check if an endpoint requires Intune-specific authentication
 */
export function isIntuneEndpoint(endpoint: string): boolean {
  const intunePatterns = [
    '/deviceManagement/',
    '/deviceAppManagement/',
    '/informationProtection/bitlocker/'
  ];
  
  return intunePatterns.some(pattern => endpoint.includes(pattern));
}

/**
 * Create an Intune-specific Graph client with the correct resource
 * Note: All Intune operations now use the standard Graph API endpoint
 */
export function createIntuneGraphClient(baseClient: Client): ModernGraphClient {
  return new ModernGraphClient(baseClient, 'https://graph.microsoft.com');
}

/**
 * Create a standard Graph client
 */
export function createStandardGraphClient(baseClient: Client): ModernGraphClient {
  return new ModernGraphClient(baseClient, 'https://graph.microsoft.com');
}

export default ModernGraphClient;
