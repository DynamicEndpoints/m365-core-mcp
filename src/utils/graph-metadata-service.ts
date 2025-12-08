import { Client } from '@microsoft/microsoft-graph-client';
import { z } from 'zod';

// Graph API Metadata Service for dynamic tool generation
export class GraphMetadataService {
  private graphClient: Client;
  private metadataCache: Map<string, any> = new Map();
  private schemaCache: Map<string, z.ZodSchema> = new Map();

  constructor(graphClient: Client) {
    this.graphClient = graphClient;
  }

  // Get Graph API metadata for endpoint discovery
  async getMetadata(version: 'v1.0' | 'beta' = 'v1.0'): Promise<any> {
    const cacheKey = `metadata_${version}`;
    
    if (this.metadataCache.has(cacheKey)) {
      return this.metadataCache.get(cacheKey);
    }

    try {
      const metadata = await this.graphClient
        .api('/$metadata')
        .version(version)
        .get();
      
      this.metadataCache.set(cacheKey, metadata);
      return metadata;
    } catch (error) {
      console.error(`Failed to fetch Graph metadata for ${version}:`, error);
      return null;
    }
  }

  // Discover available endpoints from metadata
  async discoverEndpoints(version: 'v1.0' | 'beta' = 'v1.0'): Promise<GraphEndpoint[]> {
    const metadata = await this.getMetadata(version);
    if (!metadata) return [];

    const endpoints: GraphEndpoint[] = [];
    
    // Parse CSDL metadata to extract endpoints
    // This is a simplified implementation - real implementation would parse XML
    const commonEndpoints = this.getCommonEndpoints(version);
    endpoints.push(...commonEndpoints);

    return endpoints;
  }

  // Get common Graph endpoints with their schemas
  private getCommonEndpoints(version: 'v1.0' | 'beta'): GraphEndpoint[] {
    return [
      // Teams & Communication
      {
        path: '/teams',
        methods: ['GET', 'POST'],
        category: 'teams',
        description: 'Manage Microsoft Teams',
        requiredScopes: ['Team.ReadBasic.All', 'Team.Create'],
        version
      },
      {
        path: '/teams/{team-id}/channels',
        methods: ['GET', 'POST', 'PATCH', 'DELETE'],
        category: 'teams',
        description: 'Manage team channels',
        requiredScopes: ['Channel.ReadBasic.All', 'Channel.Create'],
        version
      },
      {
        path: '/teams/{team-id}/channels/{channel-id}/messages',
        methods: ['GET', 'POST'],
        category: 'teams',
        description: 'Manage channel messages',
        requiredScopes: ['ChannelMessage.Read.All', 'ChannelMessage.Send'],
        version
      },
      {
        path: '/me/chats',
        methods: ['GET'],
        category: 'teams',
        description: 'Get user chats',
        requiredScopes: ['Chat.Read', 'Chat.ReadWrite'],
        version
      },
      {
        path: '/me/onlineMeetings',
        methods: ['GET', 'POST'],
        category: 'teams',
        description: 'Manage online meetings',
        requiredScopes: ['OnlineMeetings.ReadWrite'],
        version
      },

      // OneNote
      {
        path: '/me/onenote/notebooks',
        methods: ['GET', 'POST'],
        category: 'productivity',
        description: 'Manage OneNote notebooks',
        requiredScopes: ['Notes.ReadWrite', 'Notes.Create'],
        version
      },
      {
        path: '/me/onenote/sections',
        methods: ['GET', 'POST'],
        category: 'productivity',
        description: 'Manage OneNote sections',
        requiredScopes: ['Notes.ReadWrite', 'Notes.Create'],
        version
      },
      {
        path: '/me/onenote/pages',
        methods: ['GET', 'POST', 'PATCH'],
        category: 'productivity',
        description: 'Manage OneNote pages',
        requiredScopes: ['Notes.ReadWrite', 'Notes.Create'],
        version
      },

      // Planner
      {
        path: '/planner/plans',
        methods: ['GET', 'POST', 'PATCH', 'DELETE'],
        category: 'productivity',
        description: 'Manage Planner plans',
        requiredScopes: ['Tasks.ReadWrite'],
        version
      },
      {
        path: '/planner/buckets',
        methods: ['GET', 'POST', 'PATCH', 'DELETE'],
        category: 'productivity',
        description: 'Manage Planner buckets',
        requiredScopes: ['Tasks.ReadWrite'],
        version
      },
      {
        path: '/planner/tasks',
        methods: ['GET', 'POST', 'PATCH', 'DELETE'],
        category: 'productivity',
        description: 'Manage Planner tasks',
        requiredScopes: ['Tasks.ReadWrite'],
        version
      },

      // To Do
      {
        path: '/me/todo/lists',
        methods: ['GET', 'POST', 'PATCH', 'DELETE'],
        category: 'productivity',
        description: 'Manage To Do lists',
        requiredScopes: ['Tasks.ReadWrite'],
        version
      },
      {
        path: '/me/todo/lists/{list-id}/tasks',
        methods: ['GET', 'POST', 'PATCH', 'DELETE'],
        category: 'productivity',
        description: 'Manage To Do tasks',
        requiredScopes: ['Tasks.ReadWrite'],
        version
      },

      // Power BI (if beta)
      ...(version === 'beta' ? [
        {
          path: '/me/insights/trending',
          methods: ['GET'],
          category: 'analytics' as const,
          description: 'Get trending documents',
          requiredScopes: ['Sites.Read.All'],
          version
        },
        {
          path: '/me/insights/used',
          methods: ['GET'],
          category: 'analytics' as const,
          description: 'Get used documents',
          requiredScopes: ['Sites.Read.All'],
          version
        },
        {
          path: '/me/insights/shared',
          methods: ['GET'],
          category: 'analytics' as const,
          description: 'Get shared documents',
          requiredScopes: ['Sites.Read.All'],
          version
        }
      ] : []),

      // Bookings
      {
        path: '/solutions/bookingBusinesses',
        methods: ['GET', 'POST'],
        category: 'productivity',
        description: 'Manage booking businesses',
        requiredScopes: ['BookingsAppointment.ReadWrite.All', 'Bookings.ReadWrite.All'],
        version
      },
      {
        path: '/solutions/bookingBusinesses/{business-id}/appointments',
        methods: ['GET', 'POST', 'PATCH', 'DELETE'],
        category: 'productivity',
        description: 'Manage booking appointments',
        requiredScopes: ['BookingsAppointment.ReadWrite.All'],
        version
      },

      // Advanced Security
      {
        path: '/security/incidents',
        methods: ['GET', 'PATCH'],
        category: 'security',
        description: 'Manage security incidents',
        requiredScopes: ['SecurityIncident.Read.All', 'SecurityIncident.ReadWrite.All'],
        version
      },
      {
        path: '/security/alerts_v2',
        methods: ['GET', 'PATCH'],
        category: 'security',
        description: 'Manage security alerts v2',
        requiredScopes: ['SecurityAlert.Read.All', 'SecurityAlert.ReadWrite.All'],
        version
      },
      {
        path: '/security/threatIntelligence/articles',
        methods: ['GET'],
        category: 'security',
        description: 'Get threat intelligence articles',
        requiredScopes: ['ThreatIntelligence.Read.All'],
        version
      },

      // Reports and Analytics
      {
        path: '/reports/getTeamsUserActivityUserDetail',
        methods: ['GET'],
        category: 'analytics',
        description: 'Get Teams user activity details',
        requiredScopes: ['Reports.Read.All'],
        version
      },
      {
        path: '/reports/getOffice365ActiveUserDetail',
        methods: ['GET'],
        category: 'analytics',
        description: 'Get Office 365 active user details',
        requiredScopes: ['Reports.Read.All'],
        version
      },
      {
        path: '/reports/getSharePointSiteUsageDetail',
        methods: ['GET'],
        category: 'analytics',
        description: 'Get SharePoint site usage details',
        requiredScopes: ['Reports.Read.All'],
        version
      }
    ];
  }

  // Generate Zod schema for endpoint parameters
  generateSchema(endpoint: GraphEndpoint): z.ZodSchema {
    const cacheKey = `${endpoint.path}_${endpoint.version}`;
    
    if (this.schemaCache.has(cacheKey)) {
      return this.schemaCache.get(cacheKey)!;
    }

    // Base schema for all endpoints
    const baseSchema = z.object({
      action: z.enum(['get', 'post', 'patch', 'delete', 'list']).describe('HTTP action to perform'),
      path: z.string().optional().describe('Override the default endpoint path'),
      queryParams: z.record(z.string(), z.string()).optional().describe('Query parameters'),
      body: z.record(z.string(), z.any()).optional().describe('Request body for POST/PATCH operations'),
      fetchAll: z.boolean().optional().default(false).describe('Fetch all pages of results'),
      consistencyLevel: z.string().optional().describe('Consistency level for advanced queries'),
    });

    // Add endpoint-specific parameters based on category
    let schema = baseSchema;
    
    switch (endpoint.category) {
      case 'teams':
        schema = baseSchema.extend({
          teamId: z.string().optional().describe('Team ID for team-specific operations'),
          channelId: z.string().optional().describe('Channel ID for channel-specific operations'),
          messageId: z.string().optional().describe('Message ID for message-specific operations'),
          meetingId: z.string().optional().describe('Meeting ID for meeting-specific operations'),
        });
        break;
        
      case 'productivity':
        schema = baseSchema.extend({
          notebookId: z.string().optional().describe('OneNote notebook ID'),
          sectionId: z.string().optional().describe('OneNote section ID'),
          pageId: z.string().optional().describe('OneNote page ID'),
          planId: z.string().optional().describe('Planner plan ID'),
          bucketId: z.string().optional().describe('Planner bucket ID'),
          taskId: z.string().optional().describe('Task ID'),
          listId: z.string().optional().describe('To Do list ID'),
          businessId: z.string().optional().describe('Booking business ID'),
          appointmentId: z.string().optional().describe('Booking appointment ID'),
        });
        break;
        
      case 'security':
        schema = baseSchema.extend({
          incidentId: z.string().optional().describe('Security incident ID'),
          alertId: z.string().optional().describe('Security alert ID'),
          severity: z.enum(['low', 'medium', 'high', 'critical']).optional().describe('Alert severity filter'),
          status: z.enum(['active', 'resolved', 'dismissed']).optional().describe('Incident status filter'),
        });
        break;
        
      case 'analytics':
        schema = baseSchema.extend({
          period: z.enum(['D7', 'D30', 'D90', 'D180']).optional().describe('Report period'),
          date: z.string().optional().describe('Specific date for report (YYYY-MM-DD)'),
          format: z.enum(['json', 'csv']).optional().default('json').describe('Report format'),
        });
        break;
    }

    this.schemaCache.set(cacheKey, schema);
    return schema;
  }

  // Get required scopes for an endpoint
  getRequiredScopes(endpoint: GraphEndpoint, action: string): string[] {
    const baseScopes = endpoint.requiredScopes || [];
    
    // Add action-specific scopes
    const actionScopes: Record<string, string[]> = {
      'get': [],
      'post': ['.Create', '.ReadWrite.All'],
      'patch': ['.ReadWrite.All'],
      'delete': ['.ReadWrite.All']
    };

    const additionalScopes = actionScopes[action.toLowerCase()] || [];
    return [...baseScopes, ...additionalScopes];
  }

  // Clear caches
  clearCache(): void {
    this.metadataCache.clear();
    this.schemaCache.clear();
  }
}

// Graph endpoint interface
export interface GraphEndpoint {
  path: string;
  methods: string[];
  category: 'teams' | 'productivity' | 'security' | 'analytics' | 'power-platform' | 'viva';
  description: string;
  requiredScopes: string[];
  version: 'v1.0' | 'beta';
  parameters?: Record<string, any>;
}

// Enhanced scope management
export class GraphScopeManager {
  private static readonly SCOPE_MAPPINGS: Record<string, string[]> = {
    // Teams & Communication
    'teams': [
      'Team.ReadBasic.All',
      'Team.Create',
      'Channel.ReadBasic.All',
      'Channel.Create',
      'ChannelMessage.Read.All',
      'ChannelMessage.Send',
      'Chat.Read',
      'Chat.ReadWrite',
      'OnlineMeetings.ReadWrite'
    ],
    
    // Productivity Apps
    'productivity': [
      'Notes.ReadWrite',
      'Notes.Create',
      'Tasks.ReadWrite',
      'BookingsAppointment.ReadWrite.All',
      'Bookings.ReadWrite.All'
    ],
    
    // Security & Compliance
    'security': [
      'SecurityIncident.Read.All',
      'SecurityIncident.ReadWrite.All',
      'SecurityAlert.Read.All',
      'SecurityAlert.ReadWrite.All',
      'ThreatIntelligence.Read.All'
    ],
    
    // Analytics & Reports
    'analytics': [
      'Reports.Read.All',
      'Sites.Read.All'
    ],
    
    // Power Platform
    'power-platform': [
      'https://analysis.windows.net/powerbi/api/.default'
    ],
    
    // Viva Suite
    'viva': [
      'People.Read.All',
      'Analytics.Read'
    ]
  };

  static getScopesForCategory(category: string): string[] {
    return this.SCOPE_MAPPINGS[category] || [];
  }

  static getAllScopes(): string[] {
    return Object.values(this.SCOPE_MAPPINGS).flat();
  }

  static getScopeForEndpoint(endpoint: GraphEndpoint): string[] {
    return endpoint.requiredScopes || this.getScopesForCategory(endpoint.category);
  }
}
