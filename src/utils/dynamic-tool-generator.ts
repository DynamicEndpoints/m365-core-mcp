import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { z } from 'zod';
import { wrapToolHandler } from '../utils.js';
import { GraphMetadataService, GraphEndpoint, GraphScopeManager } from './graph-metadata-service.js';

// Dynamic tool generator for Graph API endpoints
export class DynamicToolGenerator {
  private metadataService: GraphMetadataService;
  private server: McpServer;
  private graphClient: Client;
  private getAccessToken: (scope: string) => Promise<string>;
  private validateCredentials: () => void;

  constructor(
    server: McpServer,
    graphClient: Client,
    getAccessToken: (scope: string) => Promise<string>,
    validateCredentials: () => void
  ) {
    this.server = server;
    this.graphClient = graphClient;
    this.getAccessToken = getAccessToken;
    this.validateCredentials = validateCredentials;
    this.metadataService = new GraphMetadataService(graphClient);
  }

  // Generate and register all dynamic tools
  async generateAllTools(): Promise<void> {
    console.log('🔧 Generating dynamic Graph API tools...');
    
    try {
      // Generate tools for both v1.0 and beta endpoints
      const v1Endpoints = await this.metadataService.discoverEndpoints('v1.0');
      const betaEndpoints = await this.metadataService.discoverEndpoints('beta');
      
      const allEndpoints = [...v1Endpoints, ...betaEndpoints];
      console.log(`📊 Discovered ${allEndpoints.length} Graph API endpoints`);

      // Group endpoints by category for organized tool registration
      const endpointsByCategory = this.groupEndpointsByCategory(allEndpoints);
      
      for (const [category, endpoints] of Object.entries(endpointsByCategory)) {
        await this.generateCategoryTools(category, endpoints);
      }

      console.log('✅ Dynamic Graph API tools generated successfully');
    } catch (error) {
      console.error('❌ Failed to generate dynamic tools:', error);
    }
  }

  // Group endpoints by category
  private groupEndpointsByCategory(endpoints: GraphEndpoint[]): Record<string, GraphEndpoint[]> {
    const grouped: Record<string, GraphEndpoint[]> = {};
    
    for (const endpoint of endpoints) {
      if (!grouped[endpoint.category]) {
        grouped[endpoint.category] = [];
      }
      grouped[endpoint.category].push(endpoint);
    }
    
    return grouped;
  }

  // Generate tools for a specific category
  private async generateCategoryTools(category: string, endpoints: GraphEndpoint[]): Promise<void> {
    console.log(`🔨 Generating ${category} tools (${endpoints.length} endpoints)`);

    // Create a unified tool for each category that can handle multiple endpoints
    const toolName = `manage_${category}_resources`;
    const schema = this.generateUnifiedSchema(endpoints);

    this.server.tool(
      toolName,
      schema.shape,
      wrapToolHandler(async (args: any) => {
        this.validateCredentials();
        try {
          return await this.handleDynamicRequest(category, endpoints, args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing ${category} tool: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );

    // Also create individual tools for complex endpoints
    for (const endpoint of endpoints) {
      if (this.shouldCreateIndividualTool(endpoint)) {
        await this.createIndividualTool(endpoint);
      }
    }
  }

  // Generate unified schema for category tools
  private generateUnifiedSchema(endpoints: GraphEndpoint[]): z.ZodObject<any> {
    // Base schema that all category tools share
    const baseSchema = z.object({
      endpoint: z.enum(endpoints.map(e => e.path) as [string, ...string[]]).describe('Graph API endpoint to call'),
      action: z.enum(['get', 'post', 'patch', 'delete', 'list']).describe('HTTP action to perform'),
      version: z.enum(['v1.0', 'beta']).optional().default('v1.0').describe('Graph API version'),
      queryParams: z.record(z.string()).optional().describe('Query parameters'),
      body: z.record(z.any()).optional().describe('Request body for POST/PATCH operations'),
      fetchAll: z.boolean().optional().default(false).describe('Fetch all pages of results'),
      consistencyLevel: z.string().optional().describe('Consistency level for advanced queries'),
    });

    // Add category-specific parameters
    const categoryParams = this.getCategorySpecificParams(endpoints[0]?.category);
    
    return baseSchema.extend(categoryParams);
  }

  // Get category-specific parameters
  private getCategorySpecificParams(category: string): Record<string, z.ZodSchema> {
    const params: Record<string, z.ZodSchema> = {};

    switch (category) {
      case 'teams':
        return {
          teamId: z.string().optional().describe('Team ID for team-specific operations'),
          channelId: z.string().optional().describe('Channel ID for channel-specific operations'),
          messageId: z.string().optional().describe('Message ID for message-specific operations'),
          meetingId: z.string().optional().describe('Meeting ID for meeting-specific operations'),
          userId: z.string().optional().describe('User ID for user-specific operations'),
        };
        
      case 'productivity':
        return {
          notebookId: z.string().optional().describe('OneNote notebook ID'),
          sectionId: z.string().optional().describe('OneNote section ID'),
          pageId: z.string().optional().describe('OneNote page ID'),
          planId: z.string().optional().describe('Planner plan ID'),
          bucketId: z.string().optional().describe('Planner bucket ID'),
          taskId: z.string().optional().describe('Task ID'),
          listId: z.string().optional().describe('To Do list ID'),
          businessId: z.string().optional().describe('Booking business ID'),
          appointmentId: z.string().optional().describe('Booking appointment ID'),
        };
        
      case 'security':
        return {
          incidentId: z.string().optional().describe('Security incident ID'),
          alertId: z.string().optional().describe('Security alert ID'),
          severity: z.enum(['low', 'medium', 'high', 'critical']).optional().describe('Alert severity filter'),
          status: z.enum(['active', 'resolved', 'dismissed']).optional().describe('Incident status filter'),
        };
        
      case 'analytics':
        return {
          period: z.enum(['D7', 'D30', 'D90', 'D180']).optional().describe('Report period'),
          date: z.string().optional().describe('Specific date for report (YYYY-MM-DD)'),
          format: z.enum(['json', 'csv']).optional().default('json').describe('Report format'),
        };
        
      case 'power-platform':
        return {
          workspaceId: z.string().optional().describe('Power BI workspace ID'),
          datasetId: z.string().optional().describe('Power BI dataset ID'),
          reportId: z.string().optional().describe('Power BI report ID'),
          dashboardId: z.string().optional().describe('Power BI dashboard ID'),
        };
        
      case 'viva':
        return {
          insightType: z.enum(['trending', 'used', 'shared']).optional().describe('Viva Insights type'),
          timeRange: z.enum(['week', 'month', 'quarter']).optional().describe('Time range for insights'),
        };
        
      default:
        return {};
    }
  }

  // Handle dynamic requests
  private async handleDynamicRequest(
    category: string,
    endpoints: GraphEndpoint[],
    args: any
  ): Promise<{ content: { type: string; text: string }[] }> {
    const { endpoint, action, version = 'v1.0', queryParams = {}, body, fetchAll = false, consistencyLevel } = args;

    // Find the matching endpoint
    const targetEndpoint = endpoints.find(e => e.path === endpoint && e.version === version);
    if (!targetEndpoint) {
      throw new McpError(ErrorCode.InvalidParams, `Endpoint ${endpoint} not found for ${category} category`);
    }

    // Validate that the action is supported for this endpoint
    if (!targetEndpoint.methods.includes(action.toUpperCase())) {
      throw new McpError(
        ErrorCode.InvalidParams,
        `Action ${action} not supported for endpoint ${endpoint}. Supported actions: ${targetEndpoint.methods.join(', ')}`
      );
    }

    // Build the actual API path by replacing placeholders
    let apiPath = this.buildApiPath(targetEndpoint.path, args);

    // Prepare the request
    let request = this.graphClient.api(apiPath).version(version);

    // Add query parameters
    if (Object.keys(queryParams).length > 0) {
      request = request.query(queryParams);
    }

    // Add consistency level if provided
    if (consistencyLevel) {
      request = request.header('ConsistencyLevel', consistencyLevel);
    }

    // Execute the request based on action
    let result: any;
    const startTime = Date.now();

    try {
      switch (action.toLowerCase()) {
        case 'get':
        case 'list':
          if (fetchAll) {
            result = await this.fetchAllPages(request);
          } else {
            result = await request.get();
          }
          break;
          
        case 'post':
          result = await request.post(body || {});
          break;
          
        case 'patch':
          result = await request.patch(body || {});
          break;
          
        case 'delete':
          result = await request.delete();
          if (result === undefined || result === null) {
            result = { status: 'Success (No Content)', deletedAt: new Date().toISOString() };
          }
          break;
          
        default:
          throw new McpError(ErrorCode.InvalidParams, `Unsupported action: ${action}`);
      }

      const executionTime = Date.now() - startTime;
      
      // Format response
      let responseText = `Result for ${category} API - ${action.toUpperCase()} ${apiPath}:\n`;
      responseText += `Execution time: ${executionTime}ms\n`;
      responseText += `Version: ${version}\n`;
      
      if (fetchAll && result.totalCount !== undefined) {
        responseText += `Total items fetched: ${result.totalCount}\n`;
      }
      
      responseText += `\n${JSON.stringify(result, null, 2)}`;

      return {
        content: [{ type: 'text', text: responseText }]
      };

    } catch (error) {
      const executionTime = Date.now() - startTime;
      console.error(`Error in dynamic ${category} request:`, error);
      
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to execute ${category} request: ${error instanceof Error ? error.message : 'Unknown error'} (execution time: ${executionTime}ms)`
      );
    }
  }

  // Build API path by replacing placeholders with actual values
  private buildApiPath(pathTemplate: string, args: any): string {
    let path = pathTemplate;
    
    // Replace common placeholders
    const replacements: Record<string, string> = {
      '{team-id}': args.teamId || '',
      '{channel-id}': args.channelId || '',
      '{message-id}': args.messageId || '',
      '{meeting-id}': args.meetingId || '',
      '{user-id}': args.userId || 'me',
      '{notebook-id}': args.notebookId || '',
      '{section-id}': args.sectionId || '',
      '{page-id}': args.pageId || '',
      '{plan-id}': args.planId || '',
      '{bucket-id}': args.bucketId || '',
      '{task-id}': args.taskId || '',
      '{list-id}': args.listId || '',
      '{business-id}': args.businessId || '',
      '{appointment-id}': args.appointmentId || '',
      '{incident-id}': args.incidentId || '',
      '{alert-id}': args.alertId || '',
      '{workspace-id}': args.workspaceId || '',
      '{dataset-id}': args.datasetId || '',
      '{report-id}': args.reportId || '',
      '{dashboard-id}': args.dashboardId || '',
    };

    for (const [placeholder, value] of Object.entries(replacements)) {
      if (path.includes(placeholder)) {
        if (!value) {
          throw new McpError(ErrorCode.InvalidParams, `Required parameter missing for ${placeholder}`);
        }
        path = path.replace(placeholder, value);
      }
    }

    return path;
  }

  // Fetch all pages for paginated results
  private async fetchAllPages(request: any): Promise<any> {
    let allItems: any[] = [];
    let nextLink: string | null | undefined = null;
    
    // Get first page
    const firstPageResponse = await request.get();
    const odataContext = firstPageResponse['@odata.context'];
    
    if (firstPageResponse.value && Array.isArray(firstPageResponse.value)) {
      allItems = [...firstPageResponse.value];
    }
    
    nextLink = firstPageResponse['@odata.nextLink'];
    
    // Fetch subsequent pages
    while (nextLink) {
      const nextPageResponse = await this.graphClient.api(nextLink).get();
      
      if (nextPageResponse.value && Array.isArray(nextPageResponse.value)) {
        allItems = [...allItems, ...nextPageResponse.value];
      }
      
      nextLink = nextPageResponse['@odata.nextLink'];
    }
    
    return {
      '@odata.context': odataContext,
      value: allItems,
      totalCount: allItems.length,
      fetchedAt: new Date().toISOString()
    };
  }

  // Determine if an endpoint should have its own individual tool
  private shouldCreateIndividualTool(endpoint: GraphEndpoint): boolean {
    // Create individual tools for complex or frequently used endpoints
    const complexEndpoints = [
      '/teams/{team-id}/channels/{channel-id}/messages',
      '/me/onlineMeetings',
      '/planner/plans',
      '/security/incidents',
      '/reports/getTeamsUserActivityUserDetail'
    ];
    
    return complexEndpoints.some(pattern => endpoint.path.includes(pattern.replace(/\{[^}]+\}/g, '')));
  }

  // Create individual tool for complex endpoints
  private async createIndividualTool(endpoint: GraphEndpoint): Promise<void> {
    const toolName = this.generateToolName(endpoint);
    const schema = this.metadataService.generateSchema(endpoint) as z.ZodObject<any>;

    this.server.tool(
      toolName,
      schema.shape,
      wrapToolHandler(async (args: any) => {
        this.validateCredentials();
        try {
          return await this.handleIndividualEndpoint(endpoint, args);
        } catch (error) {
          if (error instanceof McpError) {
            throw error;
          }
          throw new McpError(
            ErrorCode.InternalError,
            `Error executing ${toolName}: ${error instanceof Error ? error.message : 'Unknown error'}`
          );
        }
      })
    );
  }

  // Generate tool name from endpoint
  private generateToolName(endpoint: GraphEndpoint): string {
    const pathParts = endpoint.path.split('/').filter(part => part && !part.startsWith('{'));
    const category = endpoint.category;
    const version = endpoint.version === 'beta' ? '_beta' : '';
    
    return `${category}_${pathParts.join('_')}${version}`.toLowerCase();
  }

  // Handle individual endpoint requests
  private async handleIndividualEndpoint(endpoint: GraphEndpoint, args: any): Promise<{ content: { type: string; text: string }[] }> {
    // This would be similar to handleDynamicRequest but more specialized
    // For now, delegate to the dynamic handler
    return await this.handleDynamicRequest(endpoint.category, [endpoint], {
      ...args,
      endpoint: endpoint.path,
      version: endpoint.version
    });
  }

  // Get tool statistics
  getToolStats(): { totalEndpoints: number; categoryCounts: Record<string, number> } {
    // This would return statistics about generated tools
    return {
      totalEndpoints: 0,
      categoryCounts: {}
    };
  }
}
