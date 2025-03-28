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
    securityEnabled: boolean;
    mailEnabled: boolean;
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
    visibility: 'Private' | 'Public';
    allowExternalSenders: boolean;
    autoSubscribeNewMembers: boolean;
  };
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
