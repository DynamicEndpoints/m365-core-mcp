// DLP Policy Types
export interface DLPPolicyArgs {
  action: 'list' | 'get' | 'create' | 'update' | 'delete' | 'test';
  policyId?: string;
  name?: string;
  description?: string;
  locations?: ('Exchange' | 'SharePoint' | 'OneDrive' | 'Teams' | 'Endpoint')[];
  rules?: DLPRule[];
  settings?: {
    mode?: 'Test' | 'TestWithNotifications' | 'Enforce';
    priority?: number;
    enabled?: boolean;
  };
}

// DLP Rule Types
export interface DLPRule {
  name: string;
  conditions: DLPCondition[];
  actions: DLPAction[];
  enabled?: boolean;
  priority?: number;
}

export interface DLPCondition {
  type: 'ContentContains' | 'SensitiveInfoType' | 'DocumentProperty' | 'MessageProperty';
  value: string;
  operator?: 'Equals' | 'Contains' | 'StartsWith' | 'EndsWith' | 'RegexMatch';
  caseSensitive?: boolean;
}

export interface DLPAction {
  type: 'Block' | 'BlockWithOverride' | 'Notify' | 'Audit' | 'Quarantine';
  settings?: {
    notificationMessage?: string;
    blockMessage?: string;
    allowOverride?: boolean;
    overrideJustificationRequired?: boolean;
  };
}

// DLP Incident Types
export interface DLPIncidentArgs {
  action: 'list' | 'get' | 'resolve' | 'escalate';
  incidentId?: string;
  dateRange?: { 
    startDate: string; 
    endDate: string; 
  };
  severity?: 'Low' | 'Medium' | 'High' | 'Critical';
  status?: 'Active' | 'Resolved' | 'InProgress' | 'Dismissed';
  policyId?: string;
}

export interface DLPIncident {
  id: string;
  title: string;
  description: string;
  severity: 'Low' | 'Medium' | 'High' | 'Critical';
  status: 'Active' | 'Resolved' | 'InProgress' | 'Dismissed';
  createdDateTime: string;
  lastModifiedDateTime: string;
  policyId: string;
  policyName: string;
  affectedResources: DLPAffectedResource[];
  detectionDetails: DLPDetectionDetails;
}

export interface DLPAffectedResource {
  resourceType: 'Email' | 'Document' | 'SharePointItem' | 'OneDriveItem' | 'TeamsMessage';
  resourceId: string;
  resourceName: string;
  location: string;
  owner?: string;
}

export interface DLPDetectionDetails {
  sensitiveInfoTypes: string[];
  matchedContent: string[];
  confidenceLevel: number;
  matchCount: number;
  ruleMatched: string;
}

// DLP Sensitivity Label Types
export interface DLPSensitivityLabelArgs {
  action: 'list' | 'get' | 'create' | 'update' | 'delete' | 'apply' | 'remove';
  labelId?: string;
  name?: string;
  description?: string;
  settings?: {
    color?: string;
    sensitivity?: number;
    protectionSettings?: LabelProtectionSettings;
    markingSettings?: LabelMarkingSettings;
    autoLabelingSettings?: AutoLabelingSettings;
  };
  targetResource?: {
    resourceType: 'Email' | 'Document' | 'Site' | 'Container';
    resourceId: string;
  };
}

export interface LabelProtectionSettings {
  encryption?: {
    enabled: boolean;
    template?: string;
    permissions?: LabelPermission[];
  };
  contentMarking?: {
    watermark?: boolean;
    header?: boolean;
    footer?: boolean;
  };
  accessControl?: {
    expirationDate?: string;
    offlineAccess?: boolean;
    forwardingRestrictions?: boolean;
  };
}

export interface LabelMarkingSettings {
  watermark?: {
    text: string;
    fontSize: number;
    fontColor: string;
    layout: 'Horizontal' | 'Diagonal';
  };
  header?: {
    text: string;
    fontSize: number;
    fontColor: string;
    alignment: 'Left' | 'Center' | 'Right';
  };
  footer?: {
    text: string;
    fontSize: number;
    fontColor: string;
    alignment: 'Left' | 'Center' | 'Right';
  };
}

export interface AutoLabelingSettings {
  enabled: boolean;
  conditions: DLPCondition[];
  confidenceThreshold: number;
  simulationMode: boolean;
}

export interface LabelPermission {
  identity: string;
  permissions: ('View' | 'Edit' | 'Print' | 'Copy' | 'Export' | 'Reply' | 'ReplyAll' | 'Forward')[];
  expirationDate?: string;
}

// DLP Report Types
export interface DLPReportArgs {
  reportType: 'PolicyMatches' | 'IncidentSummary' | 'SensitiveDataExposure' | 'ComplianceStatus';
  dateRange: {
    startDate: string;
    endDate: string;
  };
  filters?: {
    policyIds?: string[];
    severity?: ('Low' | 'Medium' | 'High' | 'Critical')[];
    locations?: ('Exchange' | 'SharePoint' | 'OneDrive' | 'Teams' | 'Endpoint')[];
    users?: string[];
  };
  format: 'json' | 'csv' | 'excel';
  includeDetails?: boolean;
}

// DLP Configuration Types
export interface DLPConfigurationArgs {
  action: 'get' | 'update';
  configType: 'GlobalSettings' | 'LocationSettings' | 'NotificationSettings';
  settings?: {
    globalSettings?: {
      dlpEnabled: boolean;
      auditingEnabled: boolean;
      defaultAction: 'Allow' | 'Block' | 'Notify';
    };
    locationSettings?: {
      exchange?: LocationConfig;
      sharePoint?: LocationConfig;
      oneDrive?: LocationConfig;
      teams?: LocationConfig;
      endpoint?: LocationConfig;
    };
    notificationSettings?: {
      adminNotifications: boolean;
      userNotifications: boolean;
      emailTemplates: NotificationTemplate[];
    };
  };
}

export interface LocationConfig {
  enabled: boolean;
  includedLocations: string[];
  excludedLocations: string[];
  advancedSettings?: Record<string, any>;
}

export interface NotificationTemplate {
  id: string;
  name: string;
  subject: string;
  body: string;
  language: string;
  isDefault: boolean;
}
