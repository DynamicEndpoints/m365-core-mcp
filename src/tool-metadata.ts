/**
 * Tool metadata for MCP Server Quality
 * Provides descriptions and annotations for all tools
 */

export interface ToolMetadata {
  description: string;
  annotations?: {
    readOnlyHint?: boolean;
    destructiveHint?: boolean;
    idempotentHint?: boolean;
  };
}

export const toolMetadata: Record<string, ToolMetadata> = {
  // Distribution and Group Management
  manage_distribution_lists: {
    description: "Manage Exchange distribution lists including creation, updates, member management, and settings configuration.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  manage_security_groups: {
    description: "Manage Azure AD security groups for access control, including group creation, membership, and security settings.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  manage_m365_groups: {
    description: "Manage Microsoft 365 groups for team collaboration with shared resources like mailbox, calendar, and files.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },

  // SharePoint Management
  manage_sharepoint_sites: {
    description: "Manage SharePoint sites including creation, configuration, permissions, and site collection administration.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  manage_sharepoint_lists: {
    description: "Manage SharePoint lists and libraries including schema definition, items, views, and permissions.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },

  // User Management
  manage_user_settings: {
    description: "Manage user account settings including profile information, mailbox settings, licenses, and authentication methods.",
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true }
  },
  manage_offboarding: {
    description: "Automate user offboarding processes including account disablement, license removal, data backup, and access revocation.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },

  // Exchange Management
  manage_exchange_settings: {
    description: "Manage Exchange Online settings including mailbox configuration, transport rules, and organization policies.",
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true }
  },

  // Azure AD Management
  manage_azure_ad_roles: {
    description: "Manage Azure AD administrative roles including role assignments, custom roles, and privilege escalation controls.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  manage_azure_ad_apps: {
    description: "Manage Azure AD application registrations including app permissions, credentials, and OAuth configurations.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  manage_azure_ad_devices: {
    description: "Manage devices registered in Azure AD including device compliance, BitLocker keys, and device actions.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  manage_service_principals: {
    description: "Manage service principals for application access including permissions, credentials, and enterprise applications.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },

  // API Access
  call_microsoft_api: {
    description: "Make direct calls to any Microsoft Graph or Azure Resource Management API endpoint with full control over HTTP methods and parameters.",
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false }
  },

  // Audit and Security
  search_audit_log: {
    description: "Search and analyze Azure AD unified audit logs for security events, user activities, and compliance monitoring.",
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true }
  },
  manage_alerts: {
    description: "Manage security alerts from Microsoft Defender and other security products including investigation and remediation.",
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true }
  },

  // DLP and Information Protection
  manage_dlp_policies: {
    description: "Manage Data Loss Prevention policies to protect sensitive data across Exchange, SharePoint, OneDrive, and Teams.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  manage_dlp_incidents: {
    description: "Investigate and manage DLP policy violations and incidents including user notifications and remediation actions.",
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false }
  },
  manage_sensitivity_labels: {
    description: "Manage sensitivity labels for information protection including encryption, content marking, and classification policies.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },

  // Intune macOS Management
  manage_intune_macos_devices: {
    description: "Manage macOS devices in Intune including enrollment, compliance policies, device actions, and inventory management.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  manage_intune_macos_policies: {
    description: "Manage macOS configuration profiles and compliance policies for device security and management settings.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  manage_intune_macos_apps: {
    description: "Manage macOS application deployment including app assignments, updates, and installation requirements.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  manage_intune_macos_compliance: {
    description: "Assess macOS device compliance status and generate reports on policy adherence and security posture.",
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true }
  },

  // Intune Windows Management
  manage_intune_windows_devices: {
    description: "Manage Windows devices in Intune including enrollment, autopilot deployment, device actions, and health monitoring.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  manage_intune_windows_policies: {
    description: "Manage Windows configuration profiles and compliance policies including security baselines and update rings.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  manage_intune_windows_apps: {
    description: "Manage Windows application deployment including Win32 apps, Microsoft Store apps, and Office 365 assignments.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  manage_intune_windows_compliance: {
    description: "Assess Windows device compliance status including BitLocker encryption, antivirus status, and security configurations.",
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true }
  },

  // Compliance Framework Management
  manage_compliance_frameworks: {
    description: "Manage compliance frameworks and standards including HIPAA, GDPR, SOX, PCI-DSS, ISO 27001, and NIST configurations.",
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true }
  },
  manage_compliance_assessments: {
    description: "Conduct compliance assessments and generate detailed reports on regulatory adherence and security controls.",
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true }
  },
  manage_compliance_monitoring: {
    description: "Monitor ongoing compliance status with real-time alerts for policy violations and regulatory changes.",
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true }
  },
  manage_evidence_collection: {
    description: "Collect and preserve compliance evidence including audit logs, configuration snapshots, and attestation records.",
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true }
  },
  manage_gap_analysis: {
    description: "Perform gap analysis to identify compliance deficiencies and generate remediation recommendations.",
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true }
  },
  generate_audit_reports: {
    description: "Generate comprehensive audit reports for compliance frameworks with evidence documentation and findings.",
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true }
  },
  manage_cis_compliance: {
    description: "Manage CIS (Center for Internet Security) benchmark compliance including assessment and remediation tracking.",
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true }
  },

  // Advanced Graph API Features
  execute_graph_batch: {
    description: "Execute multiple Microsoft Graph API requests in a single batch operation for improved performance and efficiency.",
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false }
  },
  execute_delta_query: {
    description: "Track incremental changes to Microsoft Graph resources using delta queries for efficient synchronization.",
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true }
  },
  manage_graph_subscriptions: {
    description: "Manage webhook subscriptions for real-time change notifications from Microsoft Graph resources.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  execute_graph_search: {
    description: "Execute advanced search queries across Microsoft 365 content including emails, files, messages, and calendar events.",
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true }
  },

  // Policy Management
  manage_retention_policies: {
    description: "Manage retention policies for content across Exchange, SharePoint, OneDrive, and Teams with lifecycle rules.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  manage_conditional_access_policies: {
    description: "Manage Azure AD conditional access policies for zero-trust security including MFA, device compliance, and location-based controls.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  manage_information_protection_policies: {
    description: "Manage Azure Information Protection policies for data classification, encryption, and rights management.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  manage_defender_policies: {
    description: "Manage Microsoft Defender for Office 365 policies including Safe Attachments, Safe Links, anti-phishing, and anti-malware.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  manage_teams_policies: {
    description: "Manage Microsoft Teams policies for messaging, meetings, calling, apps, and live events across the organization.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  manage_exchange_policies: {
    description: "Manage Exchange Online policies including mail flow rules, mobile device access, and organization-wide settings.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  manage_sharepoint_governance_policies: {
    description: "Manage SharePoint governance policies including sharing controls, access restrictions, and site lifecycle management.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },
  manage_security_alert_policies: {
    description: "Manage security alert policies for monitoring threats, suspicious activities, and compliance violations across Microsoft 365.",
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false }
  },

  // Document Generation
  generate_powerpoint_presentation: {
    description: "Create professional PowerPoint presentations with custom slides, charts, tables, and themes from Microsoft 365 data.",
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false }
  },
  generate_word_document: {
    description: "Create professional Word documents with formatted sections, tables, charts, and table of contents from analysis data.",
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false }
  },
  generate_html_report: {
    description: "Create interactive HTML reports and dashboards with responsive design, charts, and filtering capabilities.",
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false }
  },
  generate_professional_report: {
    description: "Generate comprehensive professional reports in multiple formats (PowerPoint, Word, HTML, PDF) from Microsoft 365 data.",
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false }
  },
  oauth_authorize: {
    description: "Manage OAuth 2.0 authorization for user-delegated access to OneDrive and SharePoint files with secure token handling.",
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true }
  }
};

/**
 * Get metadata for a tool by name
 */
export function getToolMetadata(toolName: string): ToolMetadata | undefined {
  return toolMetadata[toolName];
}

/**
 * Get all tool names with metadata
 */
export function getAllToolNames(): string[] {
  return Object.keys(toolMetadata);
}
