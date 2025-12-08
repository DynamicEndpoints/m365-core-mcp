/**
 * Tool metadata for MCP Server Quality
 * Provides descriptions and annotations for all tools
 * Updated for MCP SDK v1.24.1 with openWorldHint and title support
 */

export interface ToolMetadata {
  description: string;
  title: string;
  annotations?: {
    title?: string;
    readOnlyHint?: boolean;
    destructiveHint?: boolean;
    idempotentHint?: boolean;
    openWorldHint?: boolean;
  };
}

export const toolMetadata: Record<string, ToolMetadata> = {
  // Distribution and Group Management
  manage_distribution_lists: {
    description: "Manage Exchange distribution lists including creation, updates, member management, and settings configuration.",
    title: "Distribution List Manager",
    annotations: { title: "Distribution List Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  manage_security_groups: {
    description: "Manage Azure AD security groups for access control, including group creation, membership, and security settings.",
    title: "Security Group Manager",
    annotations: { title: "Security Group Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  manage_m365_groups: {
    description: "Manage Microsoft 365 groups for team collaboration with shared resources like mailbox, calendar, and files.",
    title: "M365 Group Manager",
    annotations: { title: "M365 Group Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },

  // SharePoint Management
  manage_sharepoint_sites: {
    description: "Manage SharePoint sites including creation, configuration, permissions, and site collection administration.",
    title: "SharePoint Site Manager",
    annotations: { title: "SharePoint Site Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  manage_sharepoint_lists: {
    description: "Manage SharePoint lists and libraries including schema definition, items, views, and permissions.",
    title: "SharePoint List Manager",
    annotations: { title: "SharePoint List Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },

  // User Management
  manage_user_settings: {
    description: "Manage user account settings including profile information, mailbox settings, licenses, and authentication methods.",
    title: "User Settings Manager",
    annotations: { title: "User Settings Manager", readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: true }
  },
  manage_offboarding: {
    description: "Automate user offboarding processes including account disablement, license removal, data backup, and access revocation.",
    title: "User Offboarding Manager",
    annotations: { title: "User Offboarding Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },

  // Exchange Management
  manage_exchange_settings: {
    description: "Manage Exchange Online settings including mailbox configuration, transport rules, and organization policies.",
    title: "Exchange Settings Manager",
    annotations: { title: "Exchange Settings Manager", readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: true }
  },

  // Azure AD Management
  manage_azure_ad_roles: {
    description: "Manage Azure AD administrative roles including role assignments, custom roles, and privilege escalation controls.",
    title: "Azure AD Role Manager",
    annotations: { title: "Azure AD Role Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  manage_azure_ad_apps: {
    description: "Manage Azure AD application registrations including app permissions, credentials, and OAuth configurations.",
    title: "Azure AD App Manager",
    annotations: { title: "Azure AD App Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  manage_azure_ad_devices: {
    description: "Manage devices registered in Azure AD including device compliance, BitLocker keys, and device actions.",
    title: "Azure AD Device Manager",
    annotations: { title: "Azure AD Device Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  manage_service_principals: {
    description: "Manage service principals for application access including permissions, credentials, and enterprise applications.",
    title: "Service Principal Manager",
    annotations: { title: "Service Principal Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },

  // API Access
  call_microsoft_api: {
    description: "Make direct calls to any Microsoft Graph or Azure Resource Management API endpoint with full control over HTTP methods and parameters.",
    title: "Microsoft API Caller",
    annotations: { title: "Microsoft API Caller", readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true }
  },

  // Audit and Security
  search_audit_log: {
    description: "Search and analyze Azure AD unified audit logs for security events, user activities, and compliance monitoring.",
    title: "Audit Log Search",
    annotations: { title: "Audit Log Search", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true }
  },
  manage_alerts: {
    description: "Manage security alerts from Microsoft Defender and other security products including investigation and remediation.",
    title: "Security Alert Manager",
    annotations: { title: "Security Alert Manager", readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: true }
  },

  // DLP and Information Protection
  manage_dlp_policies: {
    description: "Manage Data Loss Prevention policies to protect sensitive data across Exchange, SharePoint, OneDrive, and Teams.",
    title: "DLP Policy Manager",
    annotations: { title: "DLP Policy Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  manage_dlp_incidents: {
    description: "Investigate and manage DLP policy violations and incidents including user notifications and remediation actions.",
    title: "DLP Incident Manager",
    annotations: { title: "DLP Incident Manager", readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true }
  },
  manage_sensitivity_labels: {
    description: "Manage sensitivity labels for information protection including encryption, content marking, and classification policies.",
    title: "Sensitivity Label Manager",
    annotations: { title: "Sensitivity Label Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },

  // Intune macOS Management
  manage_intune_macos_devices: {
    description: "Manage macOS devices in Intune including enrollment, compliance policies, device actions, and inventory management.",
    title: "Intune macOS Device Manager",
    annotations: { title: "Intune macOS Device Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  manage_intune_macos_policies: {
    description: "Manage macOS configuration profiles and compliance policies for device security and management settings.",
    title: "Intune macOS Policy Manager",
    annotations: { title: "Intune macOS Policy Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  manage_intune_macos_apps: {
    description: "Manage macOS application deployment including app assignments, updates, and installation requirements.",
    title: "Intune macOS App Manager",
    annotations: { title: "Intune macOS App Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  manage_intune_macos_compliance: {
    description: "Assess macOS device compliance status and generate reports on policy adherence and security posture.",
    title: "Intune macOS Compliance Checker",
    annotations: { title: "Intune macOS Compliance Checker", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true }
  },

  // Intune Windows Management
  manage_intune_windows_devices: {
    description: "Manage Windows devices in Intune including enrollment, autopilot deployment, device actions, and health monitoring.",
    title: "Intune Windows Device Manager",
    annotations: { title: "Intune Windows Device Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  manage_intune_windows_policies: {
    description: "Manage Windows configuration profiles and compliance policies including security baselines and update rings.",
    title: "Intune Windows Policy Manager",
    annotations: { title: "Intune Windows Policy Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  manage_intune_windows_apps: {
    description: "Manage Windows application deployment including Win32 apps, Microsoft Store apps, and Office 365 assignments.",
    title: "Intune Windows App Manager",
    annotations: { title: "Intune Windows App Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  manage_intune_windows_compliance: {
    description: "Assess Windows device compliance status including BitLocker encryption, antivirus status, and security configurations.",
    title: "Intune Windows Compliance Checker",
    annotations: { title: "Intune Windows Compliance Checker", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true }
  },

  // Compliance Framework Management
  manage_compliance_frameworks: {
    description: "Manage compliance frameworks and standards including HIPAA, GDPR, SOX, PCI-DSS, ISO 27001, and NIST configurations.",
    title: "Compliance Framework Manager",
    annotations: { title: "Compliance Framework Manager", readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: true }
  },
  manage_compliance_assessments: {
    description: "Conduct compliance assessments and generate detailed reports on regulatory adherence and security controls.",
    title: "Compliance Assessment Tool",
    annotations: { title: "Compliance Assessment Tool", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true }
  },
  manage_compliance_monitoring: {
    description: "Monitor ongoing compliance status with real-time alerts for policy violations and regulatory changes.",
    title: "Compliance Monitor",
    annotations: { title: "Compliance Monitor", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true }
  },
  manage_evidence_collection: {
    description: "Collect and preserve compliance evidence including audit logs, configuration snapshots, and attestation records.",
    title: "Evidence Collection Tool",
    annotations: { title: "Evidence Collection Tool", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true }
  },
  manage_gap_analysis: {
    description: "Perform gap analysis to identify compliance deficiencies and generate remediation recommendations.",
    title: "Compliance Gap Analyzer",
    annotations: { title: "Compliance Gap Analyzer", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true }
  },
  generate_audit_reports: {
    description: "Generate comprehensive audit reports for compliance frameworks with evidence documentation and findings.",
    title: "Audit Report Generator",
    annotations: { title: "Audit Report Generator", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true }
  },
  manage_cis_compliance: {
    description: "Manage CIS (Center for Internet Security) benchmark compliance including assessment and remediation tracking.",
    title: "CIS Compliance Manager",
    annotations: { title: "CIS Compliance Manager", readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: true }
  },

  // Advanced Graph API Features
  execute_graph_batch: {
    description: "Execute multiple Microsoft Graph API requests in a single batch operation for improved performance and efficiency.",
    title: "Graph Batch Executor",
    annotations: { title: "Graph Batch Executor", readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true }
  },
  execute_delta_query: {
    description: "Track incremental changes to Microsoft Graph resources using delta queries for efficient synchronization.",
    title: "Graph Delta Query",
    annotations: { title: "Graph Delta Query", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true }
  },
  manage_graph_subscriptions: {
    description: "Manage webhook subscriptions for real-time change notifications from Microsoft Graph resources.",
    title: "Graph Subscription Manager",
    annotations: { title: "Graph Subscription Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  execute_graph_search: {
    description: "Execute advanced search queries across Microsoft 365 content including emails, files, messages, and calendar events.",
    title: "Graph Search",
    annotations: { title: "Graph Search", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true }
  },

  // Policy Management
  manage_retention_policies: {
    description: "Manage retention policies for content across Exchange, SharePoint, OneDrive, and Teams with lifecycle rules.",
    title: "Retention Policy Manager",
    annotations: { title: "Retention Policy Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  manage_conditional_access_policies: {
    description: "Manage Azure AD conditional access policies for zero-trust security including MFA, device compliance, and location-based controls.",
    title: "Conditional Access Policy Manager",
    annotations: { title: "Conditional Access Policy Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  manage_information_protection_policies: {
    description: "Manage Azure Information Protection policies for data classification, encryption, and rights management.",
    title: "Information Protection Policy Manager",
    annotations: { title: "Information Protection Policy Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  manage_defender_policies: {
    description: "Manage Microsoft Defender for Office 365 policies including Safe Attachments, Safe Links, anti-phishing, and anti-malware.",
    title: "Defender Policy Manager",
    annotations: { title: "Defender Policy Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  manage_teams_policies: {
    description: "Manage Microsoft Teams policies for messaging, meetings, calling, apps, and live events across the organization.",
    title: "Teams Policy Manager",
    annotations: { title: "Teams Policy Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  manage_exchange_policies: {
    description: "Manage Exchange Online policies including mail flow rules, mobile device access, and organization-wide settings.",
    title: "Exchange Policy Manager",
    annotations: { title: "Exchange Policy Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  manage_sharepoint_governance_policies: {
    description: "Manage SharePoint governance policies including sharing controls, access restrictions, and site lifecycle management.",
    title: "SharePoint Governance Manager",
    annotations: { title: "SharePoint Governance Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },
  manage_security_alert_policies: {
    description: "Manage security alert policies for monitoring threats, suspicious activities, and compliance violations across Microsoft 365.",
    title: "Security Alert Policy Manager",
    annotations: { title: "Security Alert Policy Manager", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true }
  },

  // Document Generation
  generate_powerpoint_presentation: {
    description: "Create professional PowerPoint presentations with custom slides, charts, tables, and themes from Microsoft 365 data.",
    title: "PowerPoint Generator",
    annotations: { title: "PowerPoint Generator", readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true }
  },
  generate_word_document: {
    description: "Create professional Word documents with formatted sections, tables, charts, and table of contents from analysis data.",
    title: "Word Document Generator",
    annotations: { title: "Word Document Generator", readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true }
  },
  generate_html_report: {
    description: "Create interactive HTML reports and dashboards with responsive design, charts, and filtering capabilities.",
    title: "HTML Report Generator",
    annotations: { title: "HTML Report Generator", readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true }
  },
  generate_professional_report: {
    description: "Generate comprehensive professional reports in multiple formats (PowerPoint, Word, HTML, PDF) from Microsoft 365 data.",
    title: "Professional Report Generator",
    annotations: { title: "Professional Report Generator", readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true }
  },
  oauth_authorize: {
    description: "Manage OAuth 2.0 authorization for user-delegated access to OneDrive and SharePoint files with secure token handling.",
    title: "OAuth Authorization Manager",
    annotations: { title: "OAuth Authorization Manager", readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: true }
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
