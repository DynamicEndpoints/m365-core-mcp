import { z } from 'zod';

// User Management Types
export interface UserManagementArgs {
  action: 'get' | 'update';
  userId: string;
  settings?: Record<string, unknown>;
}

// Offboarding Types
export interface OffboardingArgs {
  action: 'start' | 'check' | 'complete';
  userId: string;
  options?: {
    revokeAccess?: boolean;
    retainMailbox?: boolean;
    convertToShared?: boolean;
    backupData?: boolean;
  };
}

// Distribution List Types
export interface DistributionListArgs {
  action: 'get' | 'create' | 'update' | 'delete' | 'add_members' | 'remove_members';
  listId?: string;
  displayName?: string;
  emailAddress?: string;
  members?: string[];
  settings?: {
    hideFromGAL?: boolean;
    requireSenderAuthentication?: boolean;
    moderatedBy?: string[];
  };
}

// Security Group Types
export interface SecurityGroupArgs {
  action: 'get' | 'create' | 'update' | 'delete' | 'add_members' | 'remove_members';
  groupId?: string;
  displayName?: string;
  description?: string;
  members?: string[];
  settings?: {
    securityEnabled?: boolean;
    mailEnabled?: boolean;
  };
}

// M365 Group Types
export interface M365GroupArgs {
  action: 'get' | 'create' | 'update' | 'delete' | 'add_members' | 'remove_members';
  groupId?: string;
  displayName?: string;
  description?: string;
  owners?: string[];
  members?: string[];
  settings?: {
    visibility?: 'Private' | 'Public';
    allowExternalSenders?: boolean;
    autoSubscribeNewMembers?: boolean;
  };
}

// Exchange Settings Types
export interface ExchangeSettingsArgs {
  action: 'get' | 'update';
  settingType: 'mailbox' | 'transport' | 'organization' | 'retention';
  target?: string;
  settings?: {
    automateProcessing?: {
      autoForwardEnabled?: boolean;
      autoReplyEnabled?: boolean;
    };
    rules?: {
      name: string;
      conditions: Record<string, any>;
      actions: Record<string, any>;
    }[];
    sharingPolicy?: {
      enabled?: boolean;
      domains?: string[];
    };
    retentionTags?: {
      name: string;
      type: string;
      retentionDays: number;
    }[];
  };
}

// SharePoint Site Types
export interface SharePointSiteArgs {
  action: 'get' | 'create' | 'update' | 'delete' | 'add_users' | 'remove_users';
  siteId?: string;
  url?: string;
  title?: string;
  description?: string;
  template?: string;
  owners?: string[];
  members?: string[];
  settings?: {
    isPublic?: boolean;
    allowSharing?: boolean;
    storageQuota?: number;
  };
}

// SharePoint List Types
export interface SharePointListArgs {
  action: 'get' | 'create' | 'update' | 'delete' | 'add_items' | 'get_items';
  siteId: string;
  listId?: string;
  title?: string;
  description?: string;
  template?: string;
  columns?: {
    name: string;
    type: string;
    required?: boolean;
    defaultValue?: any;
  }[];
  items?: Record<string, any>[];
}

// Azure AD Role Types
export interface AzureAdRoleArgs {
  action: 'list_roles' | 'list_role_assignments' | 'assign_role' | 'remove_role_assignment';
  roleId?: string;
  principalId?: string;
  assignmentId?: string;
  filter?: string;
}

// Azure AD App Types
export interface AzureAdAppArgs {
  action: 'list_apps' | 'get_app' | 'update_app' | 'add_owner' | 'remove_owner';
  appId?: string;
  ownerId?: string;
  appDetails?: {
    displayName?: string;
    signInAudience?: string;
    [key: string]: any;
  };
  filter?: string;
}

// Azure AD Device Types
export interface AzureAdDeviceArgs {
  action: 'list_devices' | 'get_device' | 'enable_device' | 'disable_device' | 'delete_device';
  deviceId?: string;
  filter?: string;
}

// Azure AD Service Principal Types
export interface AzureAdSpArgs {
  action: 'list_sps' | 'get_sp' | 'add_owner' | 'remove_owner';
  spId?: string;
  ownerId?: string;
  filter?: string;
}

// Generic Microsoft API Call Types - Enhanced with performance and reliability features
export interface CallMicrosoftApiArgs {
  // Core API parameters (existing)
  apiType: 'graph' | 'azure';
  path: string;
  method: 'get' | 'post' | 'put' | 'patch' | 'delete';
  apiVersion?: string;
  subscriptionId?: string;
  queryParams?: Record<string, string>;
  body?: any;
  
  // Graph API specific parameters (existing + enhanced)
  graphApiVersion?: 'v1.0' | 'beta';
  fetchAll?: boolean;
  consistencyLevel?: string;
  
  // Enhanced performance and reliability features
  maxRetries?: number;
  retryDelay?: number;
  timeout?: number;
  
  // Customization features
  customHeaders?: Record<string, string>;
  responseFormat?: 'json' | 'raw' | 'minimal';
  
  // Graph API enhancement features
  selectFields?: string[];
  expandFields?: string[];
  batchSize?: number;
}

// Audit Log Types
export interface AuditLogArgs {
  filter?: string;
  top?: number;
}

// Alert Types
export interface AlertArgs {
  action: 'list_alerts' | 'get_alert';
  alertId?: string;
  filter?: string;
  top?: number;
}

// Intune Policy Creation Types
export interface CreateIntunePolicyArgs {
  platform: 'windows' | 'macos';
  policyType: 'Configuration' | 'Compliance' | 'Security' | 'Update' | 'AppProtection' | 'EndpointSecurity';
  displayName: string;
  description?: string;
  settings?: any;
  assignments?: any[];
}
