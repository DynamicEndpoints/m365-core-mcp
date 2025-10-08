import { Client } from '@microsoft/microsoft-graph-client';

/**
 * MCP Prompts for Microsoft 365
 * These prompts provide LLMs with contextual guidance for common M365 tasks and analyses
 */

export interface PromptHandler {
  name: string;
  description: string;
  arguments?: Array<{
    name: string;
    description: string;
    required: boolean;
  }>;
  handler: (graphClient: Client, args?: Record<string, string>) => Promise<string>;
}

export const m365Prompts: PromptHandler[] = [
  // ===== SECURITY ASSESSMENT PROMPT =====
  {
    name: 'security_assessment',
    description: 'Comprehensive security posture analysis with recommendations for Microsoft 365 environment',
    arguments: [
      {
        name: 'focus_area',
        description: 'Specific area to focus on: identity, devices, data, compliance, or all (default: all)',
        required: false
      },
      {
        name: 'severity_threshold',
        description: 'Minimum severity to include: low, medium, high, critical (default: medium)',
        required: false
      }
    ],
    handler: async (graphClient, args) => {
      const focusArea = args?.focus_area || 'all';
      const threshold = args?.severity_threshold || 'medium';

      return `# Microsoft 365 Security Assessment

## Objective
Conduct a comprehensive security assessment of the Microsoft 365 environment, focusing on ${focusArea === 'all' ? 'all security areas' : focusArea}.

## Assessment Areas

### 1. Identity Protection
- Review user accounts and authentication methods
- Analyze conditional access policies and their effectiveness
- Identify risky users and sign-ins
- Review MFA adoption and coverage
- Check for privileged accounts without proper protection

**Recommended Tools:**
- \`call_microsoft_api\` with endpoint: /identityProtection/riskyUsers
- \`call_microsoft_api\` with endpoint: /identity/conditionalAccess/policies
- \`manage_azure_ad_roles\` to review privileged access
- \`search_audit_log\` to review authentication events

**Recommended Resources:**
- m365://security/identity-protection
- m365://security/conditional-access
- m365://users/privileged

### 2. Threat Detection & Response
- Review active security alerts and incidents (severity: ${threshold}+)
- Analyze threat patterns and attack vectors
- Check Microsoft Secure Score and recommendations
- Review security baseline compliance

**Recommended Tools:**
- \`manage_alerts\` to list and analyze security alerts
- \`call_microsoft_api\` with endpoint: /security/incidents
- \`call_microsoft_api\` with endpoint: /security/secureScores

**Recommended Resources:**
- m365://security/alerts
- m365://security/incidents
- m365://security/threat-assessment

### 3. Device Security
- Review device compliance status across platforms
- Analyze Intune policy effectiveness
- Identify unmanaged or non-compliant devices
- Check for devices with outdated OS or security updates

**Recommended Tools:**
- \`manage_intune_windows_devices\` for Windows device analysis
- \`manage_intune_macos_devices\` for macOS device analysis
- \`manage_intune_windows_compliance\` for compliance checking
- \`manage_azure_ad_devices\` for Azure AD registered devices

**Recommended Resources:**
- m365://devices/compliance
- m365://devices/intune-policies
- m365://devices/overview

### 4. Data Protection
- Review DLP policies and their coverage
- Analyze sensitivity label usage and adoption
- Check for data exfiltration patterns
- Review external sharing configurations

**Recommended Tools:**
- \`manage_dlp_policies\` to review data loss prevention
- \`manage_sensitivity_labels\` to check information protection
- \`manage_dlp_incidents\` to review violations
- \`call_microsoft_api\` with endpoint: /sites for SharePoint sharing

**Recommended Resources:**
- m365://compliance/policies
- m365://compliance/sensitivity-labels
- m365://sharepoint/permissions

### 5. Access Control
- Review group memberships and role assignments
- Analyze privileged access patterns
- Check for excessive permissions
- Review guest user access

**Recommended Tools:**
- \`manage_security_groups\` to review group memberships
- \`manage_azure_ad_roles\` to analyze role assignments
- \`call_microsoft_api\` with endpoint: /users?$filter=userType eq 'Guest'
- \`search_audit_log\` to review access changes

**Recommended Resources:**
- m365://users/privileged
- m365://groups/all
- m365://governance/access-reviews

## Analysis Framework

1. **Discovery**: Gather current security configurations and status
2. **Assessment**: Compare against security best practices and benchmarks
3. **Risk Evaluation**: Identify and prioritize security gaps
4. **Remediation**: Provide actionable recommendations with specific tools
5. **Monitoring**: Suggest ongoing monitoring strategies

## Output Format

Please provide:
1. **Executive Summary**: High-level findings and risk rating
2. **Detailed Findings**: Security gaps organized by severity
3. **Recommendations**: Prioritized remediation steps with specific MCP tools to use
4. **Quick Wins**: Immediate actions that can improve security posture
5. **Long-term Strategy**: Strategic recommendations for sustained security

## Compliance Context

Consider these frameworks in your assessment:
- CIS Microsoft 365 Foundations Benchmark
- NIST Cybersecurity Framework
- ISO 27001 controls
- SOC 2 requirements
- Industry-specific regulations (HIPAA, GDPR, etc.)

---

**Note**: Use the severity threshold of "${threshold}" to filter findings. Focus on ${focusArea === 'all' ? 'comprehensive coverage across all areas' : `in-depth analysis of ${focusArea}`}.`;
    }
  },

  // ===== COMPLIANCE REVIEW PROMPT =====
  {
    name: 'compliance_review',
    description: 'Framework-specific compliance gap analysis for SOC2, ISO27001, NIST, GDPR, HIPAA, or CIS',
    arguments: [
      {
        name: 'framework',
        description: 'Compliance framework: soc2, iso27001, nist, gdpr, hipaa, cis (required)',
        required: true
      },
      {
        name: 'scope',
        description: 'Specific controls or domains to focus on (optional)',
        required: false
      }
    ],
    handler: async (graphClient, args) => {
      const framework = args?.framework?.toLowerCase() || 'soc2';
      const scope = args?.scope || 'all controls';

      const frameworkDetails: Record<string, any> = {
        soc2: {
          name: 'SOC 2 Type II',
          categories: ['Security', 'Availability', 'Processing Integrity', 'Confidentiality', 'Privacy'],
          keyControls: 'CC6.1, CC6.6, CC7.2, CC8.1'
        },
        iso27001: {
          name: 'ISO/IEC 27001:2022',
          categories: ['A.5 Organizational', 'A.6 People', 'A.7 Physical', 'A.8 Technological'],
          keyControls: 'A.8.2, A.8.3, A.8.16, A.8.18'
        },
        nist: {
          name: 'NIST Cybersecurity Framework',
          categories: ['Identify', 'Protect', 'Detect', 'Respond', 'Recover'],
          keyControls: 'PR.AC, PR.DS, DE.CM, RS.AN'
        },
        gdpr: {
          name: 'GDPR',
          categories: ['Lawfulness', 'Purpose Limitation', 'Data Minimization', 'Accuracy', 'Storage Limitation', 'Security'],
          keyControls: 'Article 32 (Security), Article 33 (Breach Notification)'
        },
        hipaa: {
          name: 'HIPAA Security Rule',
          categories: ['Administrative', 'Physical', 'Technical Safeguards'],
          keyControls: '164.308, 164.310, 164.312'
        },
        cis: {
          name: 'CIS Controls v8',
          categories: ['Asset Management', 'Data Protection', 'Identity Management', 'Secure Configuration'],
          keyControls: 'CIS 1, CIS 3, CIS 5, CIS 6'
        }
      };

      const detail = frameworkDetails[framework] || frameworkDetails.soc2;

      return `# ${detail.name} Compliance Review

## Objective
Conduct a comprehensive compliance assessment against ${detail.name}, focusing on ${scope}.

## Framework Overview
**Categories**: ${detail.categories.join(', ')}
**Key Controls**: ${detail.keyControls}

## Assessment Methodology

### 1. Control Inventory
Review implemented controls in Microsoft 365:
- Identity and Access Management controls
- Data Protection and Encryption controls  
- Security Monitoring and Logging controls
- Incident Response capabilities
- Business Continuity controls

### 2. Evidence Collection
Gather evidence for compliance verification:

**Recommended Tools:**
- \`manage_compliance_frameworks\` - Configure and assess ${framework} framework
- \`manage_compliance_assessments\` - Run automated assessments
- \`manage_evidence_collection\` - Collect compliance evidence
- \`generate_audit_reports\` - Generate compliance reports
- \`manage_gap_analysis\` - Identify gaps and remediation needs

**Recommended Resources:**
- m365://compliance/policies
- m365://compliance/audit-summary
- m365://security/conditional-access
- m365://devices/compliance

### 3. Control Mapping

#### Identity & Access Controls
**${framework.toUpperCase()} Requirements:**
${framework === 'soc2' ? '- CC6.1: Logical and Physical Access Controls\n- CC6.2: Prior to Issuing Credentials\n- CC6.3: Removal of Access' : ''}
${framework === 'iso27001' ? '- A.8.2: Privileged Access Rights\n- A.8.3: Information Access Restriction\n- A.8.5: Secure Authentication' : ''}
${framework === 'nist' ? '- PR.AC-1: Identity Management\n- PR.AC-4: Access Permissions\n- PR.AC-7: Authenticator Management' : ''}
${framework === 'cis' ? '- CIS 5: Account Management\n- CIS 6: Access Control Management' : ''}

**Assessment Actions:**
- \`manage_azure_ad_roles\` - Review privileged access assignments
- \`call_microsoft_api\` with /identity/conditionalAccess/policies
- \`manage_user_settings\` - Verify user account configurations
- \`search_audit_log\` - Review access control changes

#### Data Protection Controls  
**${framework.toUpperCase()} Requirements:**
${framework === 'soc2' ? '- CC6.6: Encryption of Confidential Information\n- CC6.7: Data Classification' : ''}
${framework === 'iso27001' ? '- A.8.11: Data Masking\n- A.8.24: Cryptographic Controls' : ''}
${framework === 'gdpr' ? '- Article 32: Security of Processing\n- Article 25: Data Protection by Design' : ''}
${framework === 'cis' ? '- CIS 3: Data Protection\n- CIS 13: Network Monitoring' : ''}

**Assessment Actions:**
- \`manage_dlp_policies\` - Review data loss prevention policies
- \`manage_sensitivity_labels\` - Check information protection labels
- \`manage_dlp_incidents\` - Review data protection violations
- \`call_microsoft_api\` with /security/informationProtection

#### Security Monitoring Controls
**${framework.toUpperCase()} Requirements:**
${framework === 'soc2' ? '- CC7.2: System Monitoring\n- CC7.3: Evaluation of Security Events' : ''}
${framework === 'iso27001' ? '- A.8.15: Logging\n- A.8.16: Monitoring Activities' : ''}
${framework === 'nist' ? '- DE.CM-1: Network Monitoring\n- DE.CM-7: Unauthorized Activity Detection' : ''}

**Assessment Actions:**
- \`manage_alerts\` - Review security alert configuration
- \`search_audit_log\` - Verify audit logging completeness
- \`call_microsoft_api\` with /security/alerts_v2
- \`manage_compliance_monitoring\` - Set up continuous monitoring

#### Device Management Controls
**${framework.toUpperCase()} Requirements:**
${framework === 'soc2' ? '- CC6.8: Prevention of Unauthorized Software' : ''}
${framework === 'iso27001' ? '- A.8.1: User Endpoint Devices\n- A.8.9: Configuration Management' : ''}
${framework === 'cis' ? '- CIS 4: Secure Configuration\n- CIS 10: Malware Defenses' : ''}

**Assessment Actions:**
- \`manage_intune_windows_compliance\` - Check Windows device compliance
- \`manage_intune_macos_compliance\` - Check macOS device compliance  
- \`call_microsoft_api\` with /deviceManagement/managedDevices
- \`manage_azure_ad_devices\` - Review device registrations

### 4. Gap Analysis
**Recommended Tool:**
- \`manage_gap_analysis\` with framework: "${framework}"

This will:
1. Compare current state against ${detail.name} requirements
2. Identify control gaps and weaknesses
3. Prioritize remediation based on risk
4. Generate remediation roadmap

### 5. Remediation Planning

For each identified gap:
1. **Control Gap**: Describe the missing or insufficient control
2. **Risk Impact**: Assess the compliance and security risk
3. **Remediation Steps**: Provide specific actions using MCP tools
4. **Timeline**: Suggest implementation timeline
5. **Validation**: Define how to verify remediation

## Compliance Report Generation

**Generate Comprehensive Report:**
\`\`\`
generate_audit_reports {
  framework: "${framework}",
  reportType: "comprehensive",
  includeEvidence: true,
  format: "pdf"
}
\`\`\`

## Continuous Compliance

Set up ongoing monitoring:
- \`manage_compliance_monitoring\` - Configure automated compliance checks
- \`manage_evidence_collection\` - Automate evidence gathering
- \`manage_compliance_assessments\` - Schedule periodic assessments

## Output Requirements

Please provide:
1. **Compliance Dashboard**: Current compliance status by control category
2. **Gap Analysis**: Detailed findings for non-compliant controls
3. **Evidence Inventory**: Available evidence for each control
4. **Remediation Plan**: Prioritized action items with specific tools
5. **Timeline**: Realistic roadmap to achieve compliance
6. **Audit Readiness**: Assessment of preparedness for external audit

---

**Framework**: ${detail.name}
**Scope**: ${scope}
**Assessment Date**: ${new Date().toISOString().split('T')[0]}`;
    }
  },

  // ===== USER ACCESS REVIEW PROMPT =====
  {
    name: 'user_access_review',
    description: 'Individual or organization-wide access rights analysis and recommendations',
    arguments: [
      {
        name: 'user_id',
        description: 'Specific user to review (UPN or Object ID), or "all" for org-wide review (default: all)',
        required: false
      },
      {
        name: 'focus',
        description: 'Focus area: roles, groups, applications, external, or all (default: all)',
        required: false
      }
    ],
    handler: async (graphClient, args) => {
      const userId = args?.user_id || 'all';
      const focus = args?.focus || 'all';

      return `# User Access Rights Review

## Objective
${userId === 'all' ? 'Conduct organization-wide access review to identify excessive permissions and security risks' : `Review access rights for user: ${userId}`}

## Scope
Focus Area: ${focus === 'all' ? 'All access types' : focus}

## Review Areas

### 1. Administrative Roles
${userId === 'all' ? 'Review all users with administrative privileges' : 'Review administrative role assignments for the user'}

**Assessment Actions:**
- \`manage_azure_ad_roles\` action: "list_assignments" - Get all role assignments
- \`call_microsoft_api\` with /roleManagement/directory/roleAssignments${userId !== 'all' ? `?$filter=principalId eq '${userId}'` : ''}
- \`search_audit_log\` - Review role assignment changes (last 90 days)

**Key Questions:**
- Are privileged roles assigned appropriately?
- Is the principle of least privilege followed?
- Are there dormant admin accounts?
- Is Privileged Identity Management (PIM) in use?

### 2. Group Memberships
${userId === 'all' ? 'Analyze group membership patterns across the organization' : 'Review all group memberships for the user'}

**Assessment Actions:**
- \`manage_security_groups\` action: "list" - Get all security groups
- \`manage_m365_groups\` action: "list" - Get all M365 groups
- \`manage_distribution_lists\` action: "list" - Get distribution lists
${userId !== 'all' ? `- \`call_microsoft_api\` with /users/${userId}/memberOf` : ''}

**Key Questions:**
- Are users members of appropriate groups only?
- Are there orphaned or unused groups?
- Do security groups align with business roles?
- Are there excessive nested group memberships?

### 3. Application Access
Review enterprise application and OAuth consent grants

**Assessment Actions:**
- \`manage_service_principals\` action: "list" - List service principals
- \`call_microsoft_api\` with /oauth2PermissionGrants${userId !== 'all' ? `?$filter=principalId eq '${userId}'` : ''}
- \`manage_azure_ad_apps\` action: "list" - List application registrations

**Recommended Resources:**
- m365://applications/registrations
- m365://applications/service-principals
- m365://applications/consent

**Key Questions:**
- Are application permissions appropriate?
- Are there risky OAuth consents?
- Are unused applications identified?
- Is consent governance in place?

### 4. External Access
${userId === 'all' ? 'Review all guest user access and external sharing' : 'Review external collaboration for the user'}

**Assessment Actions:**
- \`call_microsoft_api\` with /users?$filter=userType eq 'Guest'
${userId !== 'all' ? `- \`call_microsoft_api\` with /users/${userId}/invitedBy` : ''}
- \`call_microsoft_api\` with /shares for shared content
- Review SharePoint external sharing: \`manage_sharepoint_sites\`

**Recommended Resources:**
- m365://sharepoint/permissions
- m365://groups/teams
- m365://users/directory?$filter=userType eq 'Guest'

**Key Questions:**
- Are guest accounts reviewed regularly?
- Is external sharing controlled appropriately?
- Are Teams external access policies enforced?
- Is there a guest access lifecycle process?

### 5. Data Access
Review access to sensitive data and SharePoint sites

**Assessment Actions:**
- \`manage_sharepoint_sites\` - Review site memberships
- \`call_microsoft_api\` with /sites for site collections
- \`manage_sensitivity_labels\` - Check labeled content access
${userId !== 'all' ? `- Review OneDrive sharing for user: /users/${userId}/drive/sharedWithMe` : ''}

**Recommended Resources:**
- m365://sharepoint/sites
- m365://sharepoint/permissions
- m365://compliance/sensitivity-labels

### 6. Device Access
Review device-based access and Intune enrollments

**Assessment Actions:**
- \`manage_azure_ad_devices\` action: "list" - Get user devices
${userId !== 'all' ? `- \`call_microsoft_api\` with /users/${userId}/ownedDevices` : ''}
- \`manage_intune_windows_devices\` - Review managed Windows devices
- \`manage_intune_macos_devices\` - Review managed macOS devices

**Recommended Resources:**
- m365://devices/overview
- m365://devices/compliance

## Risk-Based Analysis

### High-Risk Indicators
- Multiple administrative roles assigned
- Access to sensitive data without business justification
- Dormant accounts with elevated privileges
- External users with internal permissions
- Devices non-compliant accessing corporate resources
- Excessive OAuth application permissions

### Medium-Risk Indicators
- Group memberships exceeding role requirements
- Shared accounts or shared credentials
- Long-term guest access without review
- Application access without MFA
- Legacy authentication protocols in use

### Low-Risk Indicators
- Standard user access within job function
- Properly governed guest access
- Appropriate application permissions
- Compliant devices with conditional access

## Access Review Workflow

1. **Discovery**: Gather all access assignments ${userId === 'all' ? 'organization-wide' : 'for the user'}
2. **Categorization**: Group access by type and risk level
3. **Validation**: Verify business justification for access
4. **Risk Scoring**: Calculate access risk based on multiple factors
5. **Remediation**: Remove or adjust inappropriate access
6. **Documentation**: Record decisions and maintain audit trail

## Remediation Actions

For excessive or unnecessary access:

### Remove Role Assignments
\`\`\`
manage_azure_ad_roles {
  action: "remove_assignment",
  roleId: "<role-id>",
  principalId: "<user-id>"
}
\`\`\`

### Remove Group Memberships
\`\`\`
manage_security_groups {
  action: "remove_member",
  groupId: "<group-id>",
  memberId: "<user-id>"
}
\`\`\`

### Revoke Application Access
\`\`\`
call_microsoft_api {
  apiType: "graph",
  method: "delete",
  path: "/oauth2PermissionGrants/<grant-id>"
}
\`\`\`

### Remove External User
\`\`\`
call_microsoft_api {
  apiType: "graph",
  method: "delete",
  path: "/users/<guest-user-id>"
}
\`\`\`

## Governance Recommendations

1. **Access Review Campaigns**: Implement periodic access reviews
   - Use \`call_microsoft_api\` with /identityGovernance/accessReviews
   
2. **Entitlement Management**: Automate access lifecycle
   - Review: m365://governance/entitlement
   
3. **Conditional Access**: Enforce context-based access
   - Review: m365://security/conditional-access
   
4. **Just-in-Time Access**: Implement PIM for privileged roles
   - Configure through Azure AD Privileged Identity Management

## Output Requirements

Provide:
1. **Access Inventory**: Complete list of permissions ${userId === 'all' ? 'by user' : 'for the user'}
2. **Risk Assessment**: Categorize access by risk level
3. **Violations**: Identify policy violations or excessive permissions
4. **Recommendations**: Specific remediation actions with MCP tool commands
5. **Governance Plan**: Long-term access governance strategy

---

**Review Type**: ${userId === 'all' ? 'Organization-wide' : 'Individual user'}
**Focus**: ${focus}
**Date**: ${new Date().toISOString().split('T')[0]}`;
    }
  },

  // ===== DEVICE COMPLIANCE ANALYSIS PROMPT =====
  {
    name: 'device_compliance_analysis',
    description: 'Intune device management and compliance assessment with remediation guidance',
    arguments: [
      {
        name: 'platform',
        description: 'Platform to analyze: windows, macos, ios, android, or all (default: all)',
        required: false
      },
      {
        name: 'compliance_state',
        description: 'Filter by state: compliant, non-compliant, in-grace-period, or all (default: all)',
        required: false
      }
    ],
    handler: async (graphClient, args) => {
      const platform = args?.platform || 'all';
      const complianceState = args?.compliance_state || 'all';

      return `# Intune Device Compliance Analysis

## Objective
Assess device compliance and management posture for ${platform === 'all' ? 'all platforms' : platform} devices${complianceState !== 'all' ? ` (Status: ${complianceState})` : ''}.

## Analysis Scope

### Platforms Covered:
${platform === 'all' ? '- Windows\n- macOS\n- iOS\n- Android' : `- ${platform.charAt(0).toUpperCase() + platform.slice(1)}`}

### Compliance States:
${complianceState === 'all' ? '- Compliant devices\n- Non-compliant devices\n- Devices in grace period\n- Unknown state' : `- ${complianceState.charAt(0).toUpperCase() + complianceState.slice(1)} only`}

## Assessment Areas

### 1. Device Inventory
Get comprehensive device inventory:

**Windows Devices:**
\`\`\`
manage_intune_windows_devices {
  action: "list",
  filters: ${complianceState !== 'all' ? `{ complianceState: "${complianceState}" }` : '{}'}
}
\`\`\`

**macOS Devices:**
\`\`\`
manage_intune_macos_devices {
  action: "list",
  filters: ${complianceState !== 'all' ? `{ complianceState: "${complianceState}" }` : '{}'}
}
\`\`\`

**All Platforms:**
- \`call_microsoft_api\` with /deviceManagement/managedDevices${complianceState !== 'all' ? `?$filter=complianceState eq '${complianceState}'` : ''}

**Recommended Resources:**
- m365://devices/overview
- m365://devices/compliance

**Key Metrics:**
- Total devices by platform
- Compliance rate percentage
- Average last sync time
- Enrollment status
- OS version distribution

### 2. Compliance Policy Analysis
Review compliance policies and their effectiveness:

**Windows Compliance Policies:**
\`\`\`
manage_intune_windows_compliance {
  action: "list_policies"
}
\`\`\`

**macOS Compliance Policies:**
\`\`\`
manage_intune_macos_compliance {
  action: "list_policies"
}
\`\`\`

**All Policies:**
- \`call_microsoft_api\` with /deviceManagement/deviceCompliancePolicies

**Recommended Resources:**
- m365://devices/intune-policies

**Assessment Questions:**
- Are policies aligned with security baselines?
- Do policies cover all required security controls?
- Are compliance timeframes appropriate?
- Is conditional access integrated?

### 3. Configuration Profile Assessment
Analyze configuration profiles and deployment status:

**Windows Configuration Profiles:**
\`\`\`
manage_intune_windows_policies {
  action: "list_policies"
}
\`\`\`

**macOS Configuration Profiles:**
\`\`\`
manage_intune_macos_policies {
  action: "list_policies"
}
\`\`\`

**Profile Categories:**
- Security baselines
- Wi-Fi and VPN
- Certificate profiles
- Email configurations
- Browser settings
- Endpoint protection

### 4. Application Management
Review managed application deployment:

**Windows Applications:**
\`\`\`
manage_intune_windows_apps {
  action: "list"
}
\`\`\`

**macOS Applications:**
\`\`\`
manage_intune_macos_apps {
  action: "list"
}
\`\`\`

**Recommended Resources:**
- m365://devices/apps

**Assessment Points:**
- Application deployment success rate
- Required apps vs. available apps
- App protection policies
- Application update compliance

### 5. Security Baseline Compliance
Check compliance with security baselines:

**Windows Security Baselines:**
- Microsoft Security Baseline for Windows
- Microsoft Edge Security Baseline
- Microsoft Defender for Endpoint baseline

**macOS Security:**
- FileVault encryption status
- Firewall configuration
- Gatekeeper settings
- System integrity protection

**Assessment Tools:**
\`\`\`
manage_intune_windows_compliance {
  action: "check_compliance",
  deviceId: "<device-id>"
}
\`\`\`

### 6. Non-Compliance Analysis
Deep dive into non-compliant devices:

**Get Non-Compliant Devices:**
- \`call_microsoft_api\` with /deviceManagement/managedDevices?$filter=complianceState eq 'noncompliant'

**For Each Non-Compliant Device:**
1. Identify specific compliance violations
2. Review device configuration
3. Check last sync time
4. Review user assignment
5. Analyze remediation history

**Common Non-Compliance Reasons:**
- Outdated OS version
- Missing security updates
- BitLocker/FileVault not enabled
- Firewall disabled
- Antivirus not running or outdated
- Password policy violations
- Jailbroken/rooted devices

## Risk Assessment

### Critical Issues (Immediate Action Required)
- Devices with no sync in 30+ days
- Non-compliant devices accessing corporate resources
- Devices with disabled security features
- Unencrypted devices with corporate data
- Devices with known vulnerabilities

### High Priority Issues
- Devices in grace period nearing expiration
- Outdated OS versions (N-2 or older)
- Missing critical security updates
- Incomplete configuration profile deployment
- Devices without conditional access compliance

### Medium Priority Issues
- Devices with delayed compliance reporting
- Partial configuration profile deployment
- Non-critical application deployment failures
- Devices with pending restarts

## Remediation Strategies

### For Non-Compliant Devices

**Immediate Actions:**
1. Sync device to get latest status:
\`\`\`
manage_intune_windows_devices {
  action: "sync",
  deviceId: "<device-id>"
}
\`\`\`

2. Send device notification:
\`\`\`
manage_intune_windows_devices {
  action: "send_notification",
  deviceId: "<device-id>",
  message: "Your device is non-compliant. Please contact IT."
}
\`\`\`

3. For persistent non-compliance - Remote actions:
\`\`\`
manage_intune_windows_devices {
  action: "remote_lock",
  deviceId: "<device-id>"
}
\`\`\`

**Policy Remediation:**
- Review and adjust compliance policy settings
- Extend grace periods if appropriate
- Update conditional access policies
- Deploy remediation configuration profiles

### For Configuration Issues

**Reapply Policies:**
\`\`\`
manage_intune_windows_policies {
  action: "assign",
  policyId: "<policy-id>",
  assignments: [/* groups */]
}
\`\`\`

**Deploy Required Apps:**
\`\`\`
manage_intune_windows_apps {
  action: "assign",
  appId: "<app-id>",
  groupId: "<group-id>",
  intent: "required"
}
\`\`\`

### Automation Recommendations

1. **Automated Compliance Monitoring:**
   - Set up regular compliance scans
   - Configure automated notifications
   - Implement grace period workflows

2. **Self-Service Remediation:**
   - Company Portal messaging
   - Automated remediation scripts
   - User notification campaigns

3. **Conditional Access Integration:**
   - Block non-compliant devices from resources
   - Require compliant device for specific apps
   - Implement device-based conditional access

## Compliance Dashboard

Create a compliance dashboard showing:
- Overall compliance percentage by platform
- Trend analysis (compliance over time)
- Top non-compliance reasons
- Devices requiring immediate attention
- Policy effectiveness metrics
- User impact analysis

## Reporting Requirements

Generate comprehensive report with:

1. **Executive Summary**
   - Total devices managed
   - Overall compliance rate
   - Critical risks
   - Top remediation priorities

2. **Detailed Analysis**
   - Platform-specific compliance rates
   - Policy effectiveness metrics
   - Common non-compliance patterns
   - User and device trends

3. **Action Items**
   - Immediate remediation steps
   - Policy adjustments needed
   - Resource requirements
   - Timeline for compliance

4. **Recommendations**
   - Policy improvements
   - Automation opportunities
   - User training needs
   - Long-term strategy

---

**Platform**: ${platform}
**Compliance State**: ${complianceState}
**Assessment Date**: ${new Date().toISOString().split('T')[0]}`;
    }
  },

  // ===== COLLABORATION GOVERNANCE PROMPT =====
  {
    name: 'collaboration_governance',
    description: 'Microsoft Teams and SharePoint governance analysis with policy recommendations',
    arguments: [
      {
        name: 'focus',
        description: 'Focus area: teams, sharepoint, external-access, lifecycle, or all (default: all)',
        required: false
      }
    ],
    handler: async (graphClient, args) => {
      const focus = args?.focus || 'all';

      return `# Collaboration Governance Review

## Objective
Analyze collaboration governance for Microsoft Teams and SharePoint, focusing on ${focus === 'all' ? 'comprehensive governance across all areas' : focus}.

## Governance Areas

### 1. Teams Governance
Review Microsoft Teams creation, usage, and lifecycle:

**Assessment Actions:**
- \`call_microsoft_api\` with /groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')
- \`manage_m365_groups\` to review Teams settings
- \`call_microsoft_api\` with /teams for team details

**Recommended Resources:**
- m365://groups/teams
- m365://collaboration/teams-activity

**Key Governance Questions:**
- Who can create new teams?
- Are naming policies enforced?
- Is guest access controlled?
- Are unused teams archived?
- Is team expiration configured?

### 2. SharePoint Governance
Analyze SharePoint site creation and management:

**Assessment Actions:**
- \`manage_sharepoint_sites\` to list all sites
- \`call_microsoft_api\` with /sites
- Review external sharing settings

**Recommended Resources:**
- m365://sharepoint/sites
- m365://sharepoint/permissions

**Key Governance Questions:**
- Is site creation restricted?
- Are sensitivity labels applied?
- Is external sharing appropriate?
- Are site templates standardized?
- Is retention configured?

### 3. External Access Control
Review guest access and external sharing:

**Assessment Actions:**
- \`call_microsoft_api\` with /users?$filter=userType eq 'Guest'
- Review sharing links and permissions
- Check external access policies

**Key Questions:**
- Are guest users reviewed regularly?
- Is external sharing audited?
- Are sharing links time-limited?
- Is guest access conditional access protected?

### 4. Data Governance
Review data protection and compliance:

**Assessment Actions:**
- \`manage_sensitivity_labels\` - Check label application
- \`manage_dlp_policies\` - Review data protection policies
- \`call_microsoft_api\` with /compliance

**Recommended Resources:**
- m365://compliance/policies
- m365://compliance/sensitivity-labels

### 5. Lifecycle Management
Assess lifecycle policies and automation:

**Assessment Actions:**
- Review group expiration policies
- Check inactive teams and sites
- Analyze archival processes

**Recommendations:**
- Implement Teams lifecycle policies
- Configure site expiration
- Automate archival workflows
- Regular access reviews

## Output Requirements

Provide:
1. **Governance Dashboard**: Current state of collaboration governance
2. **Risk Assessment**: Identify governance gaps and risks
3. **Policy Recommendations**: Specific governance policies to implement
4. **Implementation Plan**: Step-by-step governance improvement roadmap
5. **Best Practices**: Industry-standard governance recommendations

---

**Focus Area**: ${focus}
**Assessment Date**: ${new Date().toISOString().split('T')[0]}`;
    }
  }
];

/**
 * Get prompt by name
 */
export function getPromptByName(name: string): PromptHandler | undefined {
  return m365Prompts.find(p => p.name === name);
}

/**
 * List all available prompts
 */
export function listPrompts(): Array<{ name: string; description: string; arguments?: Array<{ name: string; description: string; required: boolean }> }> {
  return m365Prompts.map(p => ({
    name: p.name,
    description: p.description,
    arguments: p.arguments
  }));
}
