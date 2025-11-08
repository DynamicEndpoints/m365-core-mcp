import { z } from 'zod';

// Microsoft Purview / Compliance Policy Types

// Data Loss Prevention (DLP) Policy Arguments
export interface DLPPolicyArgs {
  action: 'list' | 'get' | 'create' | 'update' | 'delete' | 'enable' | 'disable';
  policyId?: string;
  displayName?: string;
  description?: string;
  mode?: 'Test' | 'AuditAndNotify' | 'Enforce';
  priority?: number;
  locations?: {
    sharePointSites?: string[];
    exchangeEmail?: boolean;
    teamsChat?: boolean;
    oneDriveAccounts?: string[];
    deviceEndpoints?: boolean;
  };
  rules?: {
    name: string;
    conditions: {
      contentContainsSensitiveInfo?: any[];
      contentContainsLabels?: string[];
      documentIsUnsupported?: boolean;
      documentSizeOver?: number;
    };
    actions: {
      blockAccess?: boolean;
      encryptContent?: boolean;
      restrictAccess?: boolean;
      removeContent?: boolean;
    };
  }[];
}

// Retention Policy Arguments
export interface RetentionPolicyArgs {
  action: 'list' | 'get' | 'create' | 'update' | 'delete';
  policyId?: string;
  displayName?: string;
  description?: string;
  isEnabled?: boolean;
  retentionSettings: {
    retentionDuration: number; // in days
    retentionAction: 'Delete' | 'Keep' | 'KeepAndDelete';
    deletionType?: 'Immediately' | 'AfterRetentionPeriod';
  };
  locations?: {
    sharePointSites?: string[];
    exchangeEmail?: boolean;
    teamsChannels?: boolean;
    teamsChats?: boolean;
    oneDriveAccounts?: string[];
  };
}

// Sensitivity Label Arguments
export interface SensitivityLabelArgs {
  action: 'list' | 'get' | 'create' | 'update' | 'delete' | 'publish';
  labelId?: string;
  displayName?: string;
  description?: string;
  tooltip?: string;
  priority?: number;
  isEnabled?: boolean;
  settings?: {
    contentMarking?: {
      watermarkText?: string;
      headerText?: string;
      footerText?: string;
    };
    encryption?: {
      enabled: boolean;
      template?: string;
      doubleKeyEncryption?: boolean;
    };
    accessControl?: {
      permissions: {
        users: string[];
        rights: string[];
      }[];
    };
    autoLabeling?: {
      enabled: boolean;
      conditions: any[];
    };
  };
}

// Information Protection Policy Arguments
export interface InformationProtectionPolicyArgs {
  action: 'list' | 'get' | 'create' | 'update' | 'delete';
  policyId?: string;
  displayName?: string;
  description?: string;
  scope?: 'User' | 'Organization';
  settings?: {
    defaultLabelId?: string;
    requireJustification?: boolean;
    mandatoryLabelPolicy?: boolean;
    outlookDefaultLabel?: string;
    powerBIDefaultLabel?: string;
  };
}

// Conditional Access Policy Types

export interface ConditionalAccessPolicyArgs {
  action: 'list' | 'get' | 'create' | 'update' | 'delete' | 'enable' | 'disable';
  policyId?: string;
  displayName?: string;
  description?: string;
  state?: 'enabled' | 'disabled' | 'enabledForReportingButNotEnforced';
  conditions?: {
    users?: {
      includeUsers?: string[];
      excludeUsers?: string[];
      includeGroups?: string[];
      excludeGroups?: string[];
      includeRoles?: string[];
      excludeRoles?: string[];
    };
    applications?: {
      includeApplications?: string[];
      excludeApplications?: string[];
      includeUserActions?: string[];
    };
    locations?: {
      includeLocations?: string[];
      excludeLocations?: string[];
    };
    devices?: {
      includeDevices?: string[];
      excludeDevices?: string[];
      deviceFilter?: {
        mode: 'include' | 'exclude';
        rule: string;
      };
    };
    platforms?: {
      includePlatforms?: string[];
      excludePlatforms?: string[];
    };
    signInRisk?: {
      riskLevels: ('low' | 'medium' | 'high' | 'none')[];
    };
    userRisk?: {
      riskLevels: ('low' | 'medium' | 'high' | 'none')[];
    };
  };
  grantControls?: {
    operator: 'AND' | 'OR';
    builtInControls?: ('block' | 'mfa' | 'compliantDevice' | 'domainJoinedDevice' | 'approvedApplication' | 'compliantApplication')[];
    customAuthenticationFactors?: string[];
    termsOfUse?: string[];
  };
  sessionControls?: {
    applicationEnforcedRestrictions?: boolean;
    cloudAppSecurity?: {
      isEnabled: boolean;
      cloudAppSecurityType?: 'mcasConfigured' | 'monitorOnly' | 'blockDownloads';
    };
    signInFrequency?: {
      value: number;
      type: 'hours' | 'days';
    };
    persistentBrowser?: {
      mode: 'always' | 'never';
    };
  };
}

// Microsoft Defender for Office 365 Policy Types

export interface DefenderPolicyArgs {
  action: 'list' | 'get' | 'create' | 'update' | 'delete';
  policyType: 'safeAttachments' | 'safeLinks' | 'antiPhishing' | 'antiMalware' | 'antiSpam';
  policyId?: string;
  displayName?: string;
  description?: string;
  isEnabled?: boolean;
  settings?: {
    // Safe Attachments specific
    action?: 'Block' | 'Replace' | 'Allow' | 'DynamicDelivery';
    redirectToRecipients?: string[];
    actionOnError?: boolean;
    
    // Safe Links specific  
    scanUrls?: boolean;
    enableForInternalSenders?: boolean;
    trackClicks?: boolean;
    allowClickThrough?: boolean;
    
    // Anti-phishing specific
    enableMailboxIntelligence?: boolean;
    enableSpoofIntelligence?: boolean;
    enableUnauthenticatedSender?: boolean;
    enableViaTag?: boolean;
    
    // Anti-malware specific
    enableFileFilter?: boolean;
    fileTypes?: string[];
    zap?: boolean;
    
    // Anti-spam specific
    bulkThreshold?: number;
    quarantineRetentionPeriod?: number;
    enableEndUserSpamNotifications?: boolean;
  };
  appliedTo?: {
    recipientDomains?: string[];
    recipientGroups?: string[];
    recipients?: string[];
  };
}

// Microsoft Teams Policy Types

export interface TeamsPolicyArgs {
  action: 'list' | 'get' | 'create' | 'update' | 'delete' | 'assign';
  policyType: 'messaging' | 'meeting' | 'calling' | 'appSetup' | 'updateManagement';
  policyId?: string;
  displayName?: string;
  description?: string;
  settings?: {
    // Messaging policy settings
    allowOwnerDeleteMessage?: boolean;
    allowUserEditMessage?: boolean;
    allowUserDeleteMessage?: boolean;
    allowUserChat?: boolean;
    allowGiphy?: boolean;
    giphyRatingType?: 'Strict' | 'Moderate';
    allowMemes?: boolean;
    allowStickers?: boolean;
    allowUrlPreviews?: boolean;
    
    // Meeting policy settings
    allowMeetNow?: boolean;
    allowIPVideo?: boolean;
    allowAnonymousUsersToDialOut?: boolean;
    allowAnonymousUsersToStartMeeting?: boolean;
    allowPrivateMeetingScheduling?: boolean;
    allowChannelMeetingScheduling?: boolean;
    allowOutlookAddIn?: boolean;
    allowPowerPointSharing?: boolean;
    allowWhiteboard?: boolean;
    allowSharedNotes?: boolean;
    allowTranscription?: boolean;
    allowCloudRecording?: boolean;
    
    // Calling policy settings
    allowPrivateCalling?: boolean;
    allowVoicemail?: 'Enabled' | 'Disabled' | 'UserOverride';
    allowCallGroups?: boolean;
    allowDelegation?: boolean;
    allowCallForwardingToUser?: boolean;
    allowCallForwardingToPhone?: boolean;
    preventTollBypass?: boolean;
    
    // App setup policy settings
    allowUserPinning?: boolean;
    allowSideLoading?: boolean;
    pinnedApps?: {
      id: string;
      order: number;
    }[];
  };
  assignTo?: {
    users?: string[];
    groups?: string[];
  };
}

// Exchange Online Policy Types

export interface ExchangePolicyArgs {
  action: 'list' | 'get' | 'create' | 'update' | 'delete';
  policyType: 'addressBook' | 'outlookWebApp' | 'activeSyncMailbox' | 'retentionPolicy' | 'dlpPolicy';
  policyId?: string;
  displayName?: string;
  description?: string;
  isDefault?: boolean;
  settings?: {
    // Address Book Policy settings
    addressLists?: string[];
    globalAddressList?: string;
    offlineAddressBook?: string;
    roomList?: string;
    
    // Outlook Web App Policy settings
    activeSyncIntegrationEnabled?: boolean;
    allAddressListsEnabled?: boolean;
    calendarEnabled?: boolean;
    contactsEnabled?: boolean;
    journalEnabled?: boolean;
    junkEmailEnabled?: boolean;
    remindersAndNotificationsEnabled?: boolean;
    notesEnabled?: boolean;
    premiumClientEnabled?: boolean;
    searchFoldersEnabled?: boolean;
    signatureEnabled?: boolean;
    spellCheckerEnabled?: boolean;
    tasksEnabled?: boolean;
    umIntegrationEnabled?: boolean;
    changePasswordEnabled?: boolean;
    rulesEnabled?: boolean;
    publicFoldersEnabled?: boolean;
    smimeEnabled?: boolean;
    
    // ActiveSync Mailbox Policy settings
    devicePasswordEnabled?: boolean;
    alphanumericDevicePasswordRequired?: boolean;
    devicePasswordExpiration?: number;
    devicePasswordHistory?: number;
    maxDevicePasswordFailedAttempts?: number;
    maxInactivityTimeDeviceLock?: number;
    minDevicePasswordLength?: number;
    allowNonProvisionableDevices?: boolean;
    attachmentsEnabled?: boolean;
    maxAttachmentSize?: number;
    deviceEncryptionEnabled?: boolean;
    requireStorageCardEncryption?: boolean;
    passwordRecoveryEnabled?: boolean;
    requireDeviceEncryption?: boolean;
    allowCamera?: boolean;
    allowWiFi?: boolean;
    allowIrDA?: boolean;
    allowInternetSharing?: boolean;
    allowRemoteDesktop?: boolean;
    allowDesktopSync?: boolean;
    allowHTMLEmail?: boolean;
    allowTextMessaging?: boolean;
    allowPOPIMAPEmail?: boolean;
    allowBrowser?: boolean;
    allowConsumerEmail?: boolean;
    allowUnsignedApplications?: boolean;
    allowUnsignedInstallationPackages?: boolean;
  };
  appliedTo?: {
    users?: string[];
    groups?: string[];
  };
}

// SharePoint Governance Policy Types

export interface SharePointGovernancePolicyArgs {
  action: 'list' | 'get' | 'create' | 'update' | 'delete';
  policyType: 'sharingPolicy' | 'accessPolicy' | 'informationBarrier' | 'retentionLabel';
  policyId?: string;
  displayName?: string;
  description?: string;
  scope?: {
    sites?: string[];
    siteCollections?: string[];
    webApplications?: string[];
  };
  settings?: {
    // Sharing Policy settings
    sharingCapability?: 'Disabled' | 'ExternalUserSharingOnly' | 'ExternalUserAndGuestSharing' | 'ExistingExternalUserSharingOnly';
    requireAcceptanceForExternalUsers?: boolean;
    requireAnonymousLinksExpireInDays?: number;
    fileAnonymousLinkType?: 'None' | 'View' | 'Edit';
    folderAnonymousLinkType?: 'None' | 'View' | 'Edit';
    defaultSharingLinkType?: 'None' | 'Direct' | 'Internal' | 'AnonymousAccess';
    preventExternalUsersFromResharing?: boolean;
    
    // Access Policy settings
    conditionalAccessPolicy?: 'AllowFullAccess' | 'AllowLimitedAccess' | 'BlockAccess';
    limitedAccessFileType?: 'OfficeOnlineFilesOnly' | 'WebPreviewableFiles' | 'OtherFiles';
    allowDownload?: boolean;
    allowPrint?: boolean;
    allowCopy?: boolean;
    
    // Information Barrier settings
    informationBarrierMode?: 'Open' | 'Owner' | 'Members' | 'Explicit';
    
    // Retention Label settings
    retentionLabels?: {
      labelId: string;
      isDefault: boolean;
      autoApply?: boolean;
    }[];
  };
}

// Security and Compliance Alert Policy Types

export interface SecurityAlertPolicyArgs {
  action: 'list' | 'get' | 'create' | 'update' | 'delete' | 'enable' | 'disable';
  policyId?: string;
  displayName?: string;
  description?: string;
  category?: 'DataLossPrevention' | 'ThreatManagement' | 'DataGovernance' | 'AccessGovernance' | 'Others';
  severity?: 'Low' | 'Medium' | 'High' | 'Informational';
  isEnabled?: boolean;
  conditions?: {
    activityType?: string;
    objectType?: string;
    userType?: 'Admin' | 'Regular' | 'Guest' | 'System';
    locationFilter?: string[];
    timeRange?: {
      startTime: string;
      endTime: string;
    };
  };
  actions?: {
    notifyUsers?: string[];
    escalateToAdmin?: boolean;
    suppressRecurringAlerts?: boolean;
    threshold?: {
      value: number;
      timeWindow: number; // in minutes
    };
  };
}