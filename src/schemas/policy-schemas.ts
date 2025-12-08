import { z } from 'zod';

// Microsoft Purview / Compliance Policy Schemas

// DLP Policy Schema
export const dlpPolicyArgsSchema = z.object({
  action: z.enum(['list', 'get', 'create', 'update', 'delete', 'enable', 'disable']).describe('Action to perform on DLP policy'),
  policyId: z.string().optional().describe('DLP policy ID for specific operations'),
  displayName: z.string().optional().describe('Display name for the DLP policy'),
  description: z.string().optional().describe('Description of the DLP policy'),
  mode: z.enum(['Test', 'AuditAndNotify', 'Enforce']).optional().describe('DLP policy enforcement mode'),
  priority: z.number().optional().describe('Policy priority (higher number = higher priority)'),
  locations: z.object({
    sharePointSites: z.array(z.string()).optional().describe('SharePoint sites to include'),
    exchangeEmail: z.boolean().optional().describe('Include Exchange email'),
    teamsChat: z.boolean().optional().describe('Include Teams chat and channel messages'),
    oneDriveAccounts: z.array(z.string()).optional().describe('OneDrive accounts to include'),
    deviceEndpoints: z.boolean().optional().describe('Include device endpoints'),
  }).optional().describe('Locations where the policy applies'),
  rules: z.array(z.object({
    name: z.string().describe('Rule name'),
    conditions: z.object({
      contentContainsSensitiveInfo: z.array(z.any()).optional().describe('Sensitive information types'),
      contentContainsLabels: z.array(z.string()).optional().describe('Sensitivity labels'),
      documentIsUnsupported: z.boolean().optional().describe('Document is unsupported'),
      documentSizeOver: z.number().optional().describe('Document size threshold in MB'),
    }).describe('Rule conditions'),
    actions: z.object({
      blockAccess: z.boolean().optional().describe('Block access to content'),
      encryptContent: z.boolean().optional().describe('Encrypt content'),
      restrictAccess: z.boolean().optional().describe('Restrict access'),
      removeContent: z.boolean().optional().describe('Remove content'),
    }).describe('Actions to take when rule is triggered'),
  })).optional().describe('DLP policy rules'),
});

// Retention Policy Schema
export const retentionPolicyArgsSchema = z.object({
  action: z.enum(['list', 'get', 'create', 'update', 'delete']).describe('Action to perform on retention policy'),
  policyId: z.string().optional().describe('Retention policy ID for specific operations'),
  displayName: z.string().optional().describe('Display name for the retention policy'),
  description: z.string().optional().describe('Description of the retention policy'),
  isEnabled: z.boolean().optional().describe('Whether the policy is enabled'),
  retentionSettings: z.object({
    retentionDuration: z.number().describe('Retention duration in days'),
    retentionAction: z.enum(['Delete', 'Keep', 'KeepAndDelete']).describe('Action to take after retention period'),
    deletionType: z.enum(['Immediately', 'AfterRetentionPeriod']).optional().describe('When to delete content'),
  }).describe('Retention policy settings'),
  locations: z.object({
    sharePointSites: z.array(z.string()).optional().describe('SharePoint sites to include'),
    exchangeEmail: z.boolean().optional().describe('Include Exchange email'),
    teamsChannels: z.boolean().optional().describe('Include Teams channels'),
    teamsChats: z.boolean().optional().describe('Include Teams chats'),
    oneDriveAccounts: z.array(z.string()).optional().describe('OneDrive accounts to include'),
  }).optional().describe('Locations where the policy applies'),
});

// Sensitivity Label Schema
export const sensitivityLabelArgsSchema = z.object({
  action: z.enum(['list', 'get', 'create', 'update', 'delete', 'publish']).describe('Action to perform on sensitivity label'),
  labelId: z.string().optional().describe('Sensitivity label ID for specific operations'),
  displayName: z.string().optional().describe('Display name for the sensitivity label'),
  description: z.string().optional().describe('Description of the sensitivity label'),
  tooltip: z.string().optional().describe('Tooltip text for the label'),
  priority: z.number().optional().describe('Label priority (higher number = higher priority)'),
  isEnabled: z.boolean().optional().describe('Whether the label is enabled'),
  settings: z.object({
    contentMarking: z.object({
      watermarkText: z.string().optional().describe('Watermark text'),
      headerText: z.string().optional().describe('Header text'),
      footerText: z.string().optional().describe('Footer text'),
    }).optional().describe('Content marking settings'),
    encryption: z.object({
      enabled: z.boolean().describe('Enable encryption'),
      template: z.string().optional().describe('Encryption template'),
      doubleKeyEncryption: z.boolean().optional().describe('Enable double key encryption'),
    }).optional().describe('Encryption settings'),
    accessControl: z.object({
      permissions: z.array(z.object({
        users: z.array(z.string()).describe('Users with permissions'),
        rights: z.array(z.string()).describe('Rights granted'),
      })).describe('Access control permissions'),
    }).optional().describe('Access control settings'),
    autoLabeling: z.object({
      enabled: z.boolean().describe('Enable auto-labeling'),
      conditions: z.array(z.any()).describe('Auto-labeling conditions'),
    }).optional().describe('Auto-labeling settings'),
  }).optional().describe('Label settings'),
});

// Information Protection Policy Schema
export const informationProtectionPolicyArgsSchema = z.object({
  action: z.enum(['list', 'get', 'create', 'update', 'delete']).describe('Action to perform on information protection policy'),
  policyId: z.string().optional().describe('Information protection policy ID for specific operations'),
  displayName: z.string().optional().describe('Display name for the policy'),
  description: z.string().optional().describe('Description of the policy'),
  scope: z.enum(['User', 'Organization']).optional().describe('Policy scope'),
  settings: z.object({
    defaultLabelId: z.string().optional().describe('Default sensitivity label ID'),
    requireJustification: z.boolean().optional().describe('Require justification for label changes'),
    mandatoryLabelPolicy: z.boolean().optional().describe('Mandatory labeling policy'),
    outlookDefaultLabel: z.string().optional().describe('Default label for Outlook'),
    powerBIDefaultLabel: z.string().optional().describe('Default label for Power BI'),
  }).optional().describe('Policy settings'),
});

// Conditional Access Policy Schema
export const conditionalAccessPolicyArgsSchema = z.object({
  action: z.enum(['list', 'get', 'create', 'update', 'delete', 'enable', 'disable']).describe('Action to perform on Conditional Access policy'),
  policyId: z.string().optional().describe('Conditional Access policy ID for specific operations'),
  displayName: z.string().optional().describe('Display name for the policy'),
  description: z.string().optional().describe('Description of the policy'),
  state: z.enum(['enabled', 'disabled', 'enabledForReportingButNotEnforced']).optional().describe('Policy state'),
  conditions: z.object({
    users: z.object({
      includeUsers: z.array(z.string()).optional().describe('Users to include'),
      excludeUsers: z.array(z.string()).optional().describe('Users to exclude'),
      includeGroups: z.array(z.string()).optional().describe('Groups to include'),
      excludeGroups: z.array(z.string()).optional().describe('Groups to exclude'),
      includeRoles: z.array(z.string()).optional().describe('Roles to include'),
      excludeRoles: z.array(z.string()).optional().describe('Roles to exclude'),
    }).optional().describe('User conditions'),
    applications: z.object({
      includeApplications: z.array(z.string()).optional().describe('Applications to include'),
      excludeApplications: z.array(z.string()).optional().describe('Applications to exclude'),
      includeUserActions: z.array(z.string()).optional().describe('User actions to include'),
    }).optional().describe('Application conditions'),
    locations: z.object({
      includeLocations: z.array(z.string()).optional().describe('Locations to include'),
      excludeLocations: z.array(z.string()).optional().describe('Locations to exclude'),
    }).optional().describe('Location conditions'),
    devices: z.object({
      includeDevices: z.array(z.string()).optional().describe('Devices to include'),
      excludeDevices: z.array(z.string()).optional().describe('Devices to exclude'),
      deviceFilter: z.object({
        mode: z.enum(['include', 'exclude']).describe('Filter mode'),
        rule: z.string().describe('Filter rule'),
      }).optional().describe('Device filter'),
    }).optional().describe('Device conditions'),
    platforms: z.object({
      includePlatforms: z.array(z.string()).optional().describe('Platforms to include'),
      excludePlatforms: z.array(z.string()).optional().describe('Platforms to exclude'),
    }).optional().describe('Platform conditions'),
    signInRisk: z.object({
      riskLevels: z.array(z.enum(['low', 'medium', 'high', 'none'])).describe('Sign-in risk levels'),
    }).optional().describe('Sign-in risk conditions'),
    userRisk: z.object({
      riskLevels: z.array(z.enum(['low', 'medium', 'high', 'none'])).describe('User risk levels'),
    }).optional().describe('User risk conditions'),
  }).optional().describe('Policy conditions'),
  grantControls: z.object({
    operator: z.enum(['AND', 'OR']).describe('Grant controls operator'),
    builtInControls: z.array(z.enum(['block', 'mfa', 'compliantDevice', 'domainJoinedDevice', 'approvedApplication', 'compliantApplication'])).optional().describe('Built-in controls'),
    customAuthenticationFactors: z.array(z.string()).optional().describe('Custom authentication factors'),
    termsOfUse: z.array(z.string()).optional().describe('Terms of use'),
  }).optional().describe('Grant controls'),
  sessionControls: z.object({
    applicationEnforcedRestrictions: z.boolean().optional().describe('Application enforced restrictions'),
    cloudAppSecurity: z.object({
      isEnabled: z.boolean().describe('Enable cloud app security'),
      cloudAppSecurityType: z.enum(['mcasConfigured', 'monitorOnly', 'blockDownloads']).optional().describe('Cloud app security type'),
    }).optional().describe('Cloud app security controls'),
    signInFrequency: z.object({
      value: z.number().describe('Sign-in frequency value'),
      type: z.enum(['hours', 'days']).describe('Sign-in frequency type'),
    }).optional().describe('Sign-in frequency controls'),
    persistentBrowser: z.object({
      mode: z.enum(['always', 'never']).describe('Persistent browser mode'),
    }).optional().describe('Persistent browser controls'),
  }).optional().describe('Session controls'),
});

// Microsoft Defender for Office 365 Policy Schema
export const defenderPolicyArgsSchema = z.object({
  action: z.enum(['list', 'get', 'create', 'update', 'delete']).describe('Action to perform on Defender policy'),
  policyType: z.enum(['safeAttachments', 'safeLinks', 'antiPhishing', 'antiMalware', 'antiSpam']).describe('Type of Defender policy'),
  policyId: z.string().optional().describe('Defender policy ID for specific operations'),
  displayName: z.string().optional().describe('Display name for the policy'),
  description: z.string().optional().describe('Description of the policy'),
  isEnabled: z.boolean().optional().describe('Whether the policy is enabled'),
  settings: z.object({
    action: z.enum(['Block', 'Replace', 'Allow', 'DynamicDelivery']).optional().describe('Safe Attachments action'),
    redirectToRecipients: z.array(z.string()).optional().describe('Redirect recipients for Safe Attachments'),
    actionOnError: z.boolean().optional().describe('Action on error for Safe Attachments'),
    scanUrls: z.boolean().optional().describe('Scan URLs for Safe Links'),
    enableForInternalSenders: z.boolean().optional().describe('Enable Safe Links for internal senders'),
    trackClicks: z.boolean().optional().describe('Track clicks for Safe Links'),
    allowClickThrough: z.boolean().optional().describe('Allow click through for Safe Links'),
    enableMailboxIntelligence: z.boolean().optional().describe('Enable mailbox intelligence for anti-phishing'),
    enableSpoofIntelligence: z.boolean().optional().describe('Enable spoof intelligence'),
    enableUnauthenticatedSender: z.boolean().optional().describe('Enable unauthenticated sender indicators'),
    enableViaTag: z.boolean().optional().describe('Enable via tag'),
    enableFileFilter: z.boolean().optional().describe('Enable file filter for anti-malware'),
    fileTypes: z.array(z.string()).optional().describe('File types to filter'),
    zap: z.boolean().optional().describe('Enable Zero-hour Auto Purge'),
    bulkThreshold: z.number().optional().describe('Bulk email threshold'),
    quarantineRetentionPeriod: z.number().optional().describe('Quarantine retention period in days'),
    enableEndUserSpamNotifications: z.boolean().optional().describe('Enable end user spam notifications'),
  }).optional().describe('Policy settings'),
  appliedTo: z.object({
    recipientDomains: z.array(z.string()).optional().describe('Recipient domains'),
    recipientGroups: z.array(z.string()).optional().describe('Recipient groups'),
    recipients: z.array(z.string()).optional().describe('Individual recipients'),
  }).optional().describe('Policy application scope'),
});

// Microsoft Teams Policy Schema
export const teamsPolicyArgsSchema = z.object({
  action: z.enum(['list', 'get', 'create', 'update', 'delete', 'assign']).describe('Action to perform on Teams policy'),
  policyType: z.enum(['messaging', 'meeting', 'calling', 'appSetup', 'updateManagement']).describe('Type of Teams policy'),
  policyId: z.string().optional().describe('Teams policy ID for specific operations'),
  displayName: z.string().optional().describe('Display name for the policy'),
  description: z.string().optional().describe('Description of the policy'),
  settings: z.object({
    allowOwnerDeleteMessage: z.boolean().optional().describe('Allow owners to delete messages'),
    allowUserEditMessage: z.boolean().optional().describe('Allow users to edit messages'),
    allowUserDeleteMessage: z.boolean().optional().describe('Allow users to delete messages'),
    allowUserChat: z.boolean().optional().describe('Allow user chat'),
    allowGiphy: z.boolean().optional().describe('Allow Giphy'),
    giphyRatingType: z.enum(['Strict', 'Moderate']).optional().describe('Giphy rating type'),
    allowMemes: z.boolean().optional().describe('Allow memes'),
    allowStickers: z.boolean().optional().describe('Allow stickers'),
    allowUrlPreviews: z.boolean().optional().describe('Allow URL previews'),
    allowMeetNow: z.boolean().optional().describe('Allow Meet Now'),
    allowIPVideo: z.boolean().optional().describe('Allow IP video'),
    allowAnonymousUsersToDialOut: z.boolean().optional().describe('Allow anonymous users to dial out'),
    allowAnonymousUsersToStartMeeting: z.boolean().optional().describe('Allow anonymous users to start meetings'),
    allowPrivateMeetingScheduling: z.boolean().optional().describe('Allow private meeting scheduling'),
    allowChannelMeetingScheduling: z.boolean().optional().describe('Allow channel meeting scheduling'),
    allowOutlookAddIn: z.boolean().optional().describe('Allow Outlook add-in'),
    allowPowerPointSharing: z.boolean().optional().describe('Allow PowerPoint sharing'),
    allowWhiteboard: z.boolean().optional().describe('Allow whiteboard'),
    allowSharedNotes: z.boolean().optional().describe('Allow shared notes'),
    allowTranscription: z.boolean().optional().describe('Allow transcription'),
    allowCloudRecording: z.boolean().optional().describe('Allow cloud recording'),
    allowPrivateCalling: z.boolean().optional().describe('Allow private calling'),
    allowVoicemail: z.enum(['Enabled', 'Disabled', 'UserOverride']).optional().describe('Voicemail setting'),
    allowCallGroups: z.boolean().optional().describe('Allow call groups'),
    allowDelegation: z.boolean().optional().describe('Allow delegation'),
    allowCallForwardingToUser: z.boolean().optional().describe('Allow call forwarding to user'),
    allowCallForwardingToPhone: z.boolean().optional().describe('Allow call forwarding to phone'),
    preventTollBypass: z.boolean().optional().describe('Prevent toll bypass'),
    allowUserPinning: z.boolean().optional().describe('Allow user pinning of apps'),
    allowSideLoading: z.boolean().optional().describe('Allow side loading of apps'),
    pinnedApps: z.array(z.object({
      id: z.string().describe('App ID'),
      order: z.number().describe('App order'),
    })).optional().describe('Pinned apps'),
  }).optional().describe('Policy settings'),
  assignTo: z.object({
    users: z.array(z.string()).optional().describe('Users to assign policy to'),
    groups: z.array(z.string()).optional().describe('Groups to assign policy to'),
  }).optional().describe('Policy assignment'),
});

// Exchange Online Policy Schema
export const exchangePolicyArgsSchema = z.object({
  action: z.enum(['list', 'get', 'create', 'update', 'delete']).describe('Action to perform on Exchange policy'),
  policyType: z.enum(['addressBook', 'outlookWebApp', 'activeSyncMailbox', 'retentionPolicy', 'dlpPolicy']).describe('Type of Exchange policy'),
  policyId: z.string().optional().describe('Exchange policy ID for specific operations'),
  displayName: z.string().optional().describe('Display name for the policy'),
  description: z.string().optional().describe('Description of the policy'),
  isDefault: z.boolean().optional().describe('Whether this is the default policy'),
  settings: z.object({
    addressLists: z.array(z.string()).optional().describe('Address lists'),
    globalAddressList: z.string().optional().describe('Global address list'),
    offlineAddressBook: z.string().optional().describe('Offline address book'),
    roomList: z.string().optional().describe('Room list'),
    activeSyncIntegrationEnabled: z.boolean().optional().describe('ActiveSync integration enabled'),
    allAddressListsEnabled: z.boolean().optional().describe('All address lists enabled'),
    calendarEnabled: z.boolean().optional().describe('Calendar enabled'),
    contactsEnabled: z.boolean().optional().describe('Contacts enabled'),
    journalEnabled: z.boolean().optional().describe('Journal enabled'),
    junkEmailEnabled: z.boolean().optional().describe('Junk email enabled'),
    remindersAndNotificationsEnabled: z.boolean().optional().describe('Reminders and notifications enabled'),
    notesEnabled: z.boolean().optional().describe('Notes enabled'),
    premiumClientEnabled: z.boolean().optional().describe('Premium client enabled'),
    searchFoldersEnabled: z.boolean().optional().describe('Search folders enabled'),
    signatureEnabled: z.boolean().optional().describe('Signature enabled'),
    spellCheckerEnabled: z.boolean().optional().describe('Spell checker enabled'),
    tasksEnabled: z.boolean().optional().describe('Tasks enabled'),
    umIntegrationEnabled: z.boolean().optional().describe('UM integration enabled'),
    changePasswordEnabled: z.boolean().optional().describe('Change password enabled'),
    rulesEnabled: z.boolean().optional().describe('Rules enabled'),
    publicFoldersEnabled: z.boolean().optional().describe('Public folders enabled'),
    smimeEnabled: z.boolean().optional().describe('S/MIME enabled'),
    devicePasswordEnabled: z.boolean().optional().describe('Device password enabled'),
    alphanumericDevicePasswordRequired: z.boolean().optional().describe('Alphanumeric device password required'),
    devicePasswordExpiration: z.number().optional().describe('Device password expiration in days'),
    devicePasswordHistory: z.number().optional().describe('Device password history'),
    maxDevicePasswordFailedAttempts: z.number().optional().describe('Max device password failed attempts'),
    maxInactivityTimeDeviceLock: z.number().optional().describe('Max inactivity time before device lock in minutes'),
    minDevicePasswordLength: z.number().optional().describe('Minimum device password length'),
    allowNonProvisionableDevices: z.boolean().optional().describe('Allow non-provisionable devices'),
    attachmentsEnabled: z.boolean().optional().describe('Attachments enabled'),
    maxAttachmentSize: z.number().optional().describe('Max attachment size in MB'),
    deviceEncryptionEnabled: z.boolean().optional().describe('Device encryption enabled'),
    requireStorageCardEncryption: z.boolean().optional().describe('Require storage card encryption'),
    passwordRecoveryEnabled: z.boolean().optional().describe('Password recovery enabled'),
    requireDeviceEncryption: z.boolean().optional().describe('Require device encryption'),
    allowCamera: z.boolean().optional().describe('Allow camera'),
    allowWiFi: z.boolean().optional().describe('Allow WiFi'),
    allowIrDA: z.boolean().optional().describe('Allow IrDA'),
    allowInternetSharing: z.boolean().optional().describe('Allow internet sharing'),
    allowRemoteDesktop: z.boolean().optional().describe('Allow remote desktop'),
    allowDesktopSync: z.boolean().optional().describe('Allow desktop sync'),
    allowHTMLEmail: z.boolean().optional().describe('Allow HTML email'),
    allowTextMessaging: z.boolean().optional().describe('Allow text messaging'),
    allowPOPIMAPEmail: z.boolean().optional().describe('Allow POP/IMAP email'),
    allowBrowser: z.boolean().optional().describe('Allow browser'),
    allowConsumerEmail: z.boolean().optional().describe('Allow consumer email'),
    allowUnsignedApplications: z.boolean().optional().describe('Allow unsigned applications'),
    allowUnsignedInstallationPackages: z.boolean().optional().describe('Allow unsigned installation packages'),
  }).optional().describe('Policy settings'),
  appliedTo: z.object({
    users: z.array(z.string()).optional().describe('Users the policy applies to'),
    groups: z.array(z.string()).optional().describe('Groups the policy applies to'),
  }).optional().describe('Policy application scope'),
});

// SharePoint Governance Policy Schema
export const sharePointGovernancePolicyArgsSchema = z.object({
  action: z.enum(['list', 'get', 'create', 'update', 'delete']).describe('Action to perform on SharePoint governance policy'),
  policyType: z.enum(['sharingPolicy', 'accessPolicy', 'informationBarrier', 'retentionLabel']).describe('Type of SharePoint governance policy'),
  policyId: z.string().optional().describe('SharePoint governance policy ID for specific operations'),
  displayName: z.string().optional().describe('Display name for the policy'),
  description: z.string().optional().describe('Description of the policy'),
  scope: z.object({
    sites: z.array(z.string()).optional().describe('Sites the policy applies to'),
    siteCollections: z.array(z.string()).optional().describe('Site collections the policy applies to'),
    webApplications: z.array(z.string()).optional().describe('Web applications the policy applies to'),
  }).optional().describe('Policy scope'),
  settings: z.object({
    sharingCapability: z.enum(['Disabled', 'ExternalUserSharingOnly', 'ExternalUserAndGuestSharing', 'ExistingExternalUserSharingOnly']).optional().describe('Sharing capability'),
    requireAcceptanceForExternalUsers: z.boolean().optional().describe('Require acceptance for external users'),
    requireAnonymousLinksExpireInDays: z.number().optional().describe('Anonymous links expiration in days'),
    fileAnonymousLinkType: z.enum(['None', 'View', 'Edit']).optional().describe('File anonymous link type'),
    folderAnonymousLinkType: z.enum(['None', 'View', 'Edit']).optional().describe('Folder anonymous link type'),
    defaultSharingLinkType: z.enum(['None', 'Direct', 'Internal', 'AnonymousAccess']).optional().describe('Default sharing link type'),
    preventExternalUsersFromResharing: z.boolean().optional().describe('Prevent external users from resharing'),
    conditionalAccessPolicy: z.enum(['AllowFullAccess', 'AllowLimitedAccess', 'BlockAccess']).optional().describe('Conditional access policy'),
    limitedAccessFileType: z.enum(['OfficeOnlineFilesOnly', 'WebPreviewableFiles', 'OtherFiles']).optional().describe('Limited access file type'),
    allowDownload: z.boolean().optional().describe('Allow download'),
    allowPrint: z.boolean().optional().describe('Allow print'),
    allowCopy: z.boolean().optional().describe('Allow copy'),
    informationBarrierMode: z.enum(['Open', 'Owner', 'Members', 'Explicit']).optional().describe('Information barrier mode'),
    retentionLabels: z.array(z.object({
      labelId: z.string().describe('Retention label ID'),
      isDefault: z.boolean().describe('Is default label'),
      autoApply: z.boolean().optional().describe('Auto-apply label'),
    })).optional().describe('Retention labels'),
  }).optional().describe('Policy settings'),
});

// Security and Compliance Alert Policy Schema
export const securityAlertPolicyArgsSchema = z.object({
  action: z.enum(['list', 'get', 'create', 'update', 'delete', 'enable', 'disable']).describe('Action to perform on security alert policy'),
  policyId: z.string().optional().describe('Security alert policy ID for specific operations'),
  displayName: z.string().optional().describe('Display name for the policy'),
  description: z.string().optional().describe('Description of the policy'),
  category: z.enum(['DataLossPrevention', 'ThreatManagement', 'DataGovernance', 'AccessGovernance', 'Others']).optional().describe('Alert category'),
  severity: z.enum(['Low', 'Medium', 'High', 'Informational']).optional().describe('Alert severity'),
  isEnabled: z.boolean().optional().describe('Whether the policy is enabled'),
  conditions: z.object({
    activityType: z.string().optional().describe('Activity type to monitor'),
    objectType: z.string().optional().describe('Object type to monitor'),
    userType: z.enum(['Admin', 'Regular', 'Guest', 'System']).optional().describe('User type to monitor'),
    locationFilter: z.array(z.string()).optional().describe('Location filters'),
    timeRange: z.object({
      startTime: z.string().describe('Start time'),
      endTime: z.string().describe('End time'),
    }).optional().describe('Time range for alerts'),
  }).optional().describe('Alert conditions'),
  actions: z.object({
    notifyUsers: z.array(z.string()).optional().describe('Users to notify'),
    escalateToAdmin: z.boolean().optional().describe('Escalate to admin'),
    suppressRecurringAlerts: z.boolean().optional().describe('Suppress recurring alerts'),
    threshold: z.object({
      value: z.number().describe('Threshold value'),
      timeWindow: z.number().describe('Time window in minutes'),
    }).optional().describe('Alert threshold'),
  }).optional().describe('Alert actions'),
});