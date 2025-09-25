import { Client } from '@microsoft/microsoft-graph-client';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { z } from 'zod';

// Advanced Graph API features: Batch operations, webhooks, delta queries
export class GraphAdvancedFeatures {
  private graphClient: Client;
  private getAccessToken: (scope: string) => Promise<string>;

  constructor(graphClient: Client, getAccessToken: (scope: string) => Promise<string>) {
    this.graphClient = graphClient;
    this.getAccessToken = getAccessToken;
  }

  // Batch Operations - Execute multiple Graph requests in a single call
  async executeBatch(requests: BatchRequest[]): Promise<BatchResponse> {
    if (requests.length === 0) {
      throw new McpError(ErrorCode.InvalidParams, 'At least one request is required for batch operation');
    }

    if (requests.length > 20) {
      throw new McpError(ErrorCode.InvalidParams, 'Maximum 20 requests allowed per batch');
    }

    const batchPayload = {
      requests: requests.map((req, index) => ({
        id: req.id || index.toString(),
        method: req.method.toUpperCase(),
        url: req.url,
        headers: req.headers || {},
        body: req.body
      }))
    };

    try {
      const response = await this.graphClient
        .api('/$batch')
        .post(batchPayload);

      return {
        responses: response.responses,
        executedAt: new Date().toISOString(),
        totalRequests: requests.length,
        successCount: response.responses.filter((r: any) => r.status >= 200 && r.status < 300).length,
        errorCount: response.responses.filter((r: any) => r.status >= 400).length
      };
    } catch (error) {
      throw new McpError(
        ErrorCode.InternalError,
        `Batch operation failed: ${error instanceof Error ? error.message : 'Unknown error'}`
      );
    }
  }

  // Delta Queries - Efficiently track changes to Graph resources
  async executeDeltaQuery(resource: string, deltaToken?: string): Promise<DeltaQueryResponse> {
    let apiPath = resource;
    
    // Add delta function to the path
    if (!apiPath.includes('/delta')) {
      apiPath = apiPath.endsWith('/') ? `${apiPath}delta` : `${apiPath}/delta`;
    }

    try {
      let request = this.graphClient.api(apiPath);

      // If we have a delta token, use it to get only changes since last query
      if (deltaToken) {
        request = request.query({ $deltatoken: deltaToken });
      }

      const response = await request.get();

      // Extract delta link and delta token from response
      const deltaLink = response['@odata.deltaLink'];
      const nextLink = response['@odata.nextLink'];
      
      let extractedDeltaToken = '';
      if (deltaLink) {
        const tokenMatch = deltaLink.match(/\$deltatoken=([^&]+)/);
        extractedDeltaToken = tokenMatch ? decodeURIComponent(tokenMatch[1]) : '';
      }

      return {
        value: response.value || [],
        deltaToken: extractedDeltaToken,
        deltaLink: deltaLink,
        nextLink: nextLink,
        hasMoreChanges: !!nextLink,
        changeCount: response.value ? response.value.length : 0,
        queriedAt: new Date().toISOString()
      };
    } catch (error) {
      throw new McpError(
        ErrorCode.InternalError,
        `Delta query failed: ${error instanceof Error ? error.message : 'Unknown error'}`
      );
    }
  }

  // Webhook Subscriptions - Set up real-time change notifications
  async createSubscription(subscription: WebhookSubscription): Promise<SubscriptionResponse> {
    const subscriptionPayload: any = {
      changeType: subscription.changeTypes.join(','),
      notificationUrl: subscription.notificationUrl,
      resource: subscription.resource,
      expirationDateTime: subscription.expirationDateTime || this.getDefaultExpiration(),
      clientState: subscription.clientState,
      latestSupportedTlsVersion: subscription.tlsVersion || 'v1_2'
    };

    // Add lifecycle notification URL if provided
    if (subscription.lifecycleNotificationUrl) {
      subscriptionPayload.lifecycleNotificationUrl = subscription.lifecycleNotificationUrl;
    }

    try {
      const response = await this.graphClient
        .api('/subscriptions')
        .post(subscriptionPayload);

      return {
        id: response.id,
        resource: response.resource,
        changeType: response.changeType,
        notificationUrl: response.notificationUrl,
        expirationDateTime: response.expirationDateTime,
        clientState: response.clientState,
        createdAt: new Date().toISOString()
      };
    } catch (error) {
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to create subscription: ${error instanceof Error ? error.message : 'Unknown error'}`
      );
    }
  }

  // Update existing subscription
  async updateSubscription(subscriptionId: string, updates: Partial<WebhookSubscription>): Promise<SubscriptionResponse> {
    const updatePayload: any = {};

    if (updates.expirationDateTime) {
      updatePayload.expirationDateTime = updates.expirationDateTime;
    }

    if (updates.notificationUrl) {
      updatePayload.notificationUrl = updates.notificationUrl;
    }

    try {
      const response = await this.graphClient
        .api(`/subscriptions/${subscriptionId}`)
        .patch(updatePayload);

      return {
        id: response.id,
        resource: response.resource,
        changeType: response.changeType,
        notificationUrl: response.notificationUrl,
        expirationDateTime: response.expirationDateTime,
        clientState: response.clientState,
        updatedAt: new Date().toISOString()
      };
    } catch (error) {
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to update subscription: ${error instanceof Error ? error.message : 'Unknown error'}`
      );
    }
  }

  // Delete subscription
  async deleteSubscription(subscriptionId: string): Promise<{ deleted: boolean; deletedAt: string }> {
    try {
      await this.graphClient
        .api(`/subscriptions/${subscriptionId}`)
        .delete();

      return {
        deleted: true,
        deletedAt: new Date().toISOString()
      };
    } catch (error) {
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to delete subscription: ${error instanceof Error ? error.message : 'Unknown error'}`
      );
    }
  }

  // List all subscriptions
  async listSubscriptions(): Promise<SubscriptionResponse[]> {
    try {
      const response = await this.graphClient
        .api('/subscriptions')
        .get();

      return response.value.map((sub: any) => ({
        id: sub.id,
        resource: sub.resource,
        changeType: sub.changeType,
        notificationUrl: sub.notificationUrl,
        expirationDateTime: sub.expirationDateTime,
        clientState: sub.clientState
      }));
    } catch (error) {
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to list subscriptions: ${error instanceof Error ? error.message : 'Unknown error'}`
      );
    }
  }

  // Advanced Search - Use Microsoft Search API
  async executeSearch(query: SearchQuery): Promise<SearchResponse> {
    const searchPayload = {
      requests: [{
        entityTypes: query.entityTypes,
        query: {
          queryString: query.queryString
        },
        from: query.from || 0,
        size: query.size || 25,
        fields: query.fields,
        sortProperties: query.sortProperties,
        aggregations: query.aggregations,
        queryAlterationOptions: query.queryAlterationOptions
      }]
    };

    try {
      const response = await this.graphClient
        .api('/search/query')
        .post(searchPayload);

      const searchResults = response.value[0];
      
      return {
        hits: searchResults.hitsContainers[0]?.hits || [],
        totalCount: searchResults.hitsContainers[0]?.total || 0,
        moreResultsAvailable: searchResults.hitsContainers[0]?.moreResultsAvailable || false,
        aggregations: searchResults.hitsContainers[0]?.aggregations || [],
        queryAlterationResponse: searchResults.queryAlterationResponse,
        searchedAt: new Date().toISOString()
      };
    } catch (error) {
      throw new McpError(
        ErrorCode.InternalError,
        `Search query failed: ${error instanceof Error ? error.message : 'Unknown error'}`
      );
    }
  }

  // Get default expiration time for subscriptions (maximum allowed)
  private getDefaultExpiration(): string {
    const now = new Date();
    // Most subscriptions have a maximum lifetime of 4230 minutes (about 3 days)
    now.setMinutes(now.getMinutes() + 4230);
    return now.toISOString();
  }

  // Validate webhook notification (for webhook endpoint implementation)
  validateWebhookNotification(notification: any, clientState?: string): boolean {
    if (!notification || !notification.value) {
      return false;
    }

    // Validate client state if provided
    if (clientState && notification.clientState !== clientState) {
      return false;
    }

    // Validate required fields
    const requiredFields = ['subscriptionId', 'changeType', 'resource'];
    for (const field of requiredFields) {
      if (!notification[field]) {
        return false;
      }
    }

    return true;
  }

  // Process webhook notification
  processWebhookNotification(notification: any): ProcessedNotification {
    return {
      subscriptionId: notification.subscriptionId,
      changeType: notification.changeType,
      resource: notification.resource,
      resourceData: notification.resourceData,
      subscriptionExpirationDateTime: notification.subscriptionExpirationDateTime,
      clientState: notification.clientState,
      tenantId: notification.tenantId,
      processedAt: new Date().toISOString()
    };
  }
}

// Type definitions for advanced features
export interface BatchRequest {
  id?: string;
  method: 'GET' | 'POST' | 'PATCH' | 'PUT' | 'DELETE';
  url: string;
  headers?: Record<string, string>;
  body?: any;
}

export interface BatchResponse {
  responses: Array<{
    id: string;
    status: number;
    headers: Record<string, string>;
    body: any;
  }>;
  executedAt: string;
  totalRequests: number;
  successCount: number;
  errorCount: number;
}

export interface DeltaQueryResponse {
  value: any[];
  deltaToken: string;
  deltaLink?: string;
  nextLink?: string;
  hasMoreChanges: boolean;
  changeCount: number;
  queriedAt: string;
}

export interface WebhookSubscription {
  resource: string;
  changeTypes: ('created' | 'updated' | 'deleted')[];
  notificationUrl: string;
  expirationDateTime?: string;
  clientState?: string;
  lifecycleNotificationUrl?: string;
  tlsVersion?: 'v1_0' | 'v1_1' | 'v1_2' | 'v1_3';
}

export interface SubscriptionResponse {
  id: string;
  resource: string;
  changeType: string;
  notificationUrl: string;
  expirationDateTime: string;
  clientState?: string;
  createdAt?: string;
  updatedAt?: string;
}

export interface SearchQuery {
  entityTypes: ('message' | 'event' | 'drive' | 'driveItem' | 'list' | 'listItem' | 'site' | 'person')[];
  queryString: string;
  from?: number;
  size?: number;
  fields?: string[];
  sortProperties?: Array<{
    name: string;
    isDescending?: boolean;
  }>;
  aggregations?: Array<{
    field: string;
    size?: number;
    bucketDefinition?: any;
  }>;
  queryAlterationOptions?: {
    enableSuggestion?: boolean;
    enableModification?: boolean;
  };
}

export interface SearchResponse {
  hits: Array<{
    hitId: string;
    rank: number;
    summary: string;
    resource: any;
  }>;
  totalCount: number;
  moreResultsAvailable: boolean;
  aggregations: any[];
  queryAlterationResponse?: any;
  searchedAt: string;
}

export interface ProcessedNotification {
  subscriptionId: string;
  changeType: string;
  resource: string;
  resourceData?: any;
  subscriptionExpirationDateTime?: string;
  clientState?: string;
  tenantId?: string;
  processedAt: string;
}

// Zod schemas for validation
export const batchRequestSchema = z.object({
  requests: z.array(z.object({
    id: z.string().optional(),
    method: z.enum(['GET', 'POST', 'PATCH', 'PUT', 'DELETE']),
    url: z.string(),
    headers: z.record(z.string()).optional(),
    body: z.any().optional()
  })).min(1).max(20)
});

export const deltaQuerySchema = z.object({
  resource: z.string().describe('Graph resource path (e.g., /users, /groups)'),
  deltaToken: z.string().optional().describe('Delta token from previous query')
});

export const webhookSubscriptionSchema = z.object({
  resource: z.string().describe('Graph resource to monitor'),
  changeTypes: z.array(z.enum(['created', 'updated', 'deleted'])).describe('Types of changes to monitor'),
  notificationUrl: z.string().url().describe('Webhook endpoint URL'),
  expirationDateTime: z.string().optional().describe('Subscription expiration (ISO 8601)'),
  clientState: z.string().optional().describe('Client state for validation'),
  lifecycleNotificationUrl: z.string().url().optional().describe('Lifecycle notification URL'),
  tlsVersion: z.enum(['v1_0', 'v1_1', 'v1_2', 'v1_3']).optional().describe('Minimum TLS version')
});

export const searchQuerySchema = z.object({
  entityTypes: z.array(z.enum(['message', 'event', 'drive', 'driveItem', 'list', 'listItem', 'site', 'person'])),
  queryString: z.string().describe('Search query string'),
  from: z.number().min(0).optional().describe('Starting index for results'),
  size: z.number().min(1).max(1000).optional().describe('Number of results to return'),
  fields: z.array(z.string()).optional().describe('Fields to include in results'),
  sortProperties: z.array(z.object({
    name: z.string(),
    isDescending: z.boolean().optional()
  })).optional().describe('Sort properties'),
  aggregations: z.array(z.object({
    field: z.string(),
    size: z.number().optional(),
    bucketDefinition: z.any().optional()
  })).optional().describe('Aggregation definitions'),
  queryAlterationOptions: z.object({
    enableSuggestion: z.boolean().optional(),
    enableModification: z.boolean().optional()
  }).optional().describe('Query alteration options')
});
