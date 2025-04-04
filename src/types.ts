export interface UserManagementArgs {
  action: 'get' | 'update';
  userPrincipalName: string;
  settings?: Record<string, unknown>;
}

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

export interface ExchangeSettingsArgs {
  action: 'get' | 'update';
  settingType: 'mailbox' | 'transport' | 'organization' | 'retention';
  target?: string; // User/Group ID for mailbox settings
  settings?: {
    // Mailbox settings
    automateProcessing?: {
      autoReplyEnabled?: boolean;
      autoForwardEnabled?: boolean;
    };
    // Transport settings
    rules?: {
      name: string;
      conditions: Record<string, unknown>;
      actions: Record<string, unknown>;
    }[];
    // Organization settings
    sharingPolicy?: {
      domains: string[];
      enabled: boolean;
    };
    // Retention settings
    retentionTags?: {
      name: string;
      type: string;
      retentionDays: number;
    }[];
  };
}

// --- Azure AD Types ---
export interface AzureAdRoleArgs {
  action: 'list_roles' | 'list_role_assignments' | 'assign_role' | 'remove_role_assignment';
  roleId?: string; // ID of the directoryRole
  principalId?: string; // ID of the user, group, or SP
  assignmentId?: string; // ID of the role assignment
  filter?: string; // OData filter
}

export interface AzureAdAppArgs {
  action: 'list_apps' | 'get_app' | 'update_app' | 'add_owner' | 'remove_owner';
  appId?: string; // Object ID of the application
  ownerId?: string; // Object ID of the user to add/remove as owner
  appDetails?: { // Details for update_app
    displayName?: string;
    signInAudience?: string;
    // Add other updatable properties as needed
  };
  filter?: string; // OData filter for list_apps
}

export interface AzureAdDeviceArgs {
  action: 'list_devices' | 'get_device' | 'enable_device' | 'disable_device' | 'delete_device';
  deviceId?: string; // Object ID of the device
  filter?: string; // OData filter for list_devices
}

export interface AzureAdSpArgs {
  action: 'list_sps' | 'get_sp' | 'add_owner' | 'remove_owner';
  spId?: string; // Object ID of the Service Principal
  ownerId?: string; // Object ID of the user to add/remove as owner
  filter?: string; // OData filter for list_sps
}

export interface CallMicrosoftApiArgs {
  apiType: 'graph' | 'azure';
  path: string;
  method: 'get' | 'post' | 'put' | 'patch' | 'delete';
  apiVersion?: string;
  subscriptionId?: string;
  queryParams?: Record<string, string>;
  body?: any;
}

// --- Security & Compliance Types ---
export interface AuditLogArgs {
  filter?: string;
  top?: number;
}

export interface AlertArgs {
  action: 'list_alerts' | 'get_alert';
  alertId?: string;
  filter?: string;
  top?: number;
}
