# Document Generation - Quick Reference Guide

## 🚀 Quick Start

Generate professional reports, presentations, and documents from Microsoft 365 data.

## 🎯 Available Tools

### 1. generate_powerpoint_presentation
Create PowerPoint presentations with charts and custom layouts

### 2. generate_word_document
Generate Word documents with tables, charts, and professional formatting

### 3. generate_html_report
Create interactive HTML dashboards with themes

### 4. generate_professional_report
Generate comprehensive multi-format reports from Graph data

### 5. oauth_authorize
Manage OAuth 2.0 authorization for user-delegated file access

---

## 📋 JSON Examples

### PowerPoint Presentation

#### Create Presentation
```json
{
  "action": "create",
  "driveId": "b!abc123def456...",
  "folderPath": "Reports/Q4",
  "fileName": "security_analysis.pptx",
  "title": "Q4 Security Analysis",
  "template": "corporate",
  "slides": [
    {
      "layout": "title",
      "title": "Security Overview",
      "subtitle": "Q4 2024 Report"
    },
    {
      "layout": "chart",
      "title": "Alert Trends",
      "content": "Monthly alert statistics",
      "chart": {
        "type": "line",
        "data": {
          "labels": ["Oct", "Nov", "Dec"],
          "datasets": [{
            "label": "Critical Alerts",
            "data": [12, 8, 15],
            "borderColor": "#dc3545"
          }]
        }
      }
    },
    {
      "layout": "two-column",
      "title": "Key Findings",
      "leftColumn": {
        "title": "Strengths",
        "items": ["95% patch compliance", "Zero data breaches"]
      },
      "rightColumn": {
        "title": "Areas for Improvement",
        "items": ["MFA adoption at 65%", "3 legacy systems"]
      }
    }
  ]
}
```

#### List Presentations
```json
{
  "action": "list",
  "driveId": "b!abc123def456...",
  "folderPath": "Reports"
}
```

#### Export to PDF
```json
{
  "action": "export",
  "driveId": "b!abc123def456...",
  "itemId": "01ABC123DEF456...",
  "format": "pdf"
}
```

**Slide Layouts:**
- `title`: Title slide
- `title-content`: Title with bullet points
- `two-column`: Two-column layout
- `comparison`: Side-by-side comparison
- `chart`: Chart visualization
- `image`: Image with caption
- `blank`: Blank slide

---

### Word Document

#### Create Document
```json
{
  "action": "create",
  "driveId": "b!abc123def456...",
  "folderPath": "Reports",
  "fileName": "compliance_report.docx",
  "title": "Compliance Audit Report",
  "template": "business",
  "sections": [
    {
      "type": "heading1",
      "content": "Executive Summary"
    },
    {
      "type": "paragraph",
      "content": "This report presents the findings of our Q4 compliance audit.",
      "alignment": "left"
    },
    {
      "type": "heading2",
      "content": "Compliance Results"
    },
    {
      "type": "table",
      "content": "Policy Compliance Status",
      "data": {
        "headers": ["Policy", "Status", "Score", "Notes"],
        "rows": [
          ["Password Policy", "Compliant", "95%", "Excellent coverage"],
          ["MFA Policy", "Non-Compliant", "65%", "Needs improvement"],
          ["Device Encryption", "Compliant", "100%", "Full deployment"]
        ]
      }
    },
    {
      "type": "list",
      "items": [
        "Implement MFA awareness campaign",
        "Review legacy system access",
        "Update password complexity requirements"
      ],
      "ordered": true
    }
  ],
  "header": "Confidential - Internal Use Only",
  "footer": "Page {PAGE} of {NUMPAGES}",
  "includeTableOfContents": true
}
```

#### Append to Document
```json
{
  "action": "append",
  "driveId": "b!abc123def456...",
  "itemId": "01ABC123DEF456...",
  "sections": [
    {
      "type": "pageBreak"
    },
    {
      "type": "heading1",
      "content": "Appendix A: Detailed Findings"
    },
    {
      "type": "paragraph",
      "content": "Additional compliance data..."
    }
  ]
}
```

**Section Types:**
- `heading1`, `heading2`, `heading3`: Headings
- `paragraph`: Text paragraphs
- `table`: Data tables
- `list`: Bullet/numbered lists
- `image`: Images with captions
- `pageBreak`: Page breaks

**Templates:**
- `business`: Professional corporate
- `academic`: Academic style (APA/MLA)
- `report`: Executive report
- `memo`: Internal memo

---

### HTML Report

#### Create Dashboard
```json
{
  "action": "create",
  "driveId": "b!abc123def456...",
  "folderPath": "Dashboards",
  "fileName": "user_activity.html",
  "title": "User Activity Dashboard",
  "theme": "dashboard",
  "description": "Real-time user engagement metrics",
  "author": "IT Security Team",
  "sections": [
    {
      "type": "heading",
      "content": "User Activity Overview",
      "level": 1
    },
    {
      "type": "grid",
      "columns": 4,
      "items": [
        {
          "type": "card",
          "title": "Active Users",
          "content": "1,245",
          "className": "bg-primary text-white"
        },
        {
          "type": "card",
          "title": "New This Week",
          "content": "87",
          "className": "bg-success text-white"
        },
        {
          "type": "card",
          "title": "Login Success",
          "content": "98.5%",
          "className": "bg-info text-white"
        },
        {
          "type": "card",
          "title": "Alerts",
          "content": "3",
          "className": "bg-warning"
        }
      ]
    },
    {
      "type": "chart",
      "chart": {
        "type": "bar",
        "data": {
          "labels": ["Week 1", "Week 2", "Week 3", "Week 4"],
          "datasets": [{
            "label": "Active Users",
            "data": [1150, 1189, 1220, 1245],
            "backgroundColor": "#0078d4"
          }]
        }
      }
    },
    {
      "type": "table",
      "data": {
        "headers": ["User", "Last Login", "Department", "Status"],
        "rows": [
          ["John Doe", "2024-01-15 09:30", "IT", "Active"],
          ["Jane Smith", "2024-01-15 08:45", "HR", "Active"],
          ["Bob Johnson", "2024-01-14 16:20", "Sales", "Active"]
        ]
      }
    },
    {
      "type": "alert",
      "content": "3 users require password reset",
      "alertType": "warning"
    }
  ]
}
```

**Themes:**
- `modern`: Corporate blue, professional
- `classic`: Traditional, serif fonts
- `minimal`: Clean, simple design
- `dashboard`: Dark mode, gradients

**Section Types:**
- `heading`: H1-H6 headings
- `paragraph`: Text blocks
- `table`: Data tables
- `chart`: Chart.js visualizations
- `list`: Bullet/numbered lists
- `grid`: Responsive columns
- `card`: Content cards
- `alert`: Notifications (info/warning/success/danger)

---

### Professional Report

#### Security Analysis Report
```json
{
  "reportType": "security-analysis",
  "title": "Q4 2024 Security Posture Assessment",
  "description": "Comprehensive security analysis for Q4 2024",
  "dataQueries": [
    {
      "source": "alerts",
      "filter": "severity eq 'high' or severity eq 'critical'",
      "transformation": "count",
      "groupBy": "category"
    },
    {
      "source": "devices",
      "filter": "complianceState eq 'noncompliant'",
      "transformation": "count",
      "groupBy": "operatingSystem"
    },
    {
      "source": "audit-logs",
      "filter": "activityDisplayName eq 'Sign-in activity'",
      "transformation": "trend",
      "groupBy": "result"
    }
  ],
  "formats": ["pptx", "docx", "html", "pdf"],
  "driveId": "b!abc123def456...",
  "folderPath": "Reports/Q4",
  "includeExecutiveSummary": true,
  "includeRecommendations": true
}
```

#### Compliance Audit Report
```json
{
  "reportType": "compliance-audit",
  "title": "Annual Compliance Audit 2024",
  "description": "Comprehensive compliance verification",
  "dataQueries": [
    {
      "source": "policies",
      "transformation": "count",
      "groupBy": "complianceStatus"
    },
    {
      "source": "compliance",
      "transformation": "aggregate",
      "aggregation": "avg",
      "field": "complianceScore"
    },
    {
      "source": "devices",
      "filter": "complianceState eq 'noncompliant'",
      "transformation": "count"
    }
  ],
  "formats": ["pptx", "docx", "html"],
  "driveId": "b!abc123def456...",
  "folderPath": "Compliance/Annual",
  "includeExecutiveSummary": true,
  "includeRecommendations": true
}
```

#### User Activity Report
```json
{
  "reportType": "user-activity",
  "title": "User Engagement Analysis - Q4",
  "dataQueries": [
    {
      "source": "users",
      "filter": "accountEnabled eq true",
      "transformation": "count"
    },
    {
      "source": "audit-logs",
      "filter": "activityDisplayName eq 'Sign-in activity'",
      "transformation": "trend",
      "groupBy": "createdDateTime"
    }
  ],
  "formats": ["html", "pptx"],
  "driveId": "b!abc123def456...",
  "folderPath": "Reports/UserActivity",
  "includeExecutiveSummary": false,
  "includeRecommendations": true
}
```

**Report Types:**
- `security-analysis`: Security posture assessment
- `compliance-audit`: Compliance verification
- `user-activity`: User behavior analysis
- `device-health`: Device inventory and health
- `custom`: Custom data queries

**Data Sources:**
- `users`: User accounts
- `devices`: Device inventory
- `groups`: Group memberships
- `audit-logs`: Audit trail
- `alerts`: Security alerts
- `policies`: Applied policies
- `compliance`: Compliance status

**Transformations:**
- `count`: Count records
- `group-by`: Group and count
- `aggregate`: Sum/avg/min/max
- `trend`: Time-series analysis

---

### OAuth Authorization

#### Step 1: Get Authorization URL
```json
{
  "action": "get-auth-url",
  "scopes": [
    "Files.ReadWrite",
    "Sites.ReadWrite.All",
    "User.Read",
    "offline_access"
  ],
  "state": "random-csrf-token-12345"
}
```

**Response:**
```json
{
  "sessionId": "550e8400-e29b-41d4-a716-446655440000",
  "authorizationUrl": "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize?client_id=...&response_type=code&redirect_uri=...&scope=Files.ReadWrite%20Sites.ReadWrite.All&state=random-csrf-token-12345"
}
```

#### Step 2: Exchange Code for Token
```json
{
  "action": "exchange-code",
  "sessionId": "550e8400-e29b-41d4-a716-446655440000",
  "authorizationCode": "code-from-redirect",
  "state": "random-csrf-token-12345"
}
```

**Response:**
```json
{
  "accessToken": "eyJ0eXAiOiJKV1QiLCJhbGc...",
  "refreshToken": "0.AXcA...",
  "expiresIn": 3600,
  "tokenType": "Bearer",
  "scope": "Files.ReadWrite Sites.ReadWrite.All User.Read"
}
```

#### Step 3: Refresh Token
```json
{
  "action": "refresh-token",
  "sessionId": "550e8400-e29b-41d4-a716-446655440000",
  "refreshToken": "0.AXcA..."
}
```

#### Step 4: Revoke Token
```json
{
  "action": "revoke",
  "sessionId": "550e8400-e29b-41d4-a716-446655440000",
  "token": "eyJ0eXAiOiJKV1QiLCJhbGc..."
}
```

**Available Scopes:**
- `Files.ReadWrite`: Read/write user files
- `Files.ReadWrite.All`: Read/write all files
- `Sites.ReadWrite.All`: SharePoint site access
- `User.Read`: User profile
- `offline_access`: Refresh tokens

---

## 🎨 Chart Types

All document types support Chart.js visualizations:

### Line Chart
```json
{
  "type": "line",
  "data": {
    "labels": ["Jan", "Feb", "Mar", "Apr"],
    "datasets": [{
      "label": "Active Users",
      "data": [450, 520, 485, 510],
      "borderColor": "#0078d4"
    }]
  }
}
```

### Bar Chart
```json
{
  "type": "bar",
  "data": {
    "labels": ["Q1", "Q2", "Q3", "Q4"],
    "datasets": [{
      "label": "Revenue",
      "data": [65, 78, 82, 91],
      "backgroundColor": "#107c10"
    }]
  }
}
```

### Pie/Doughnut Chart
```json
{
  "type": "doughnut",
  "data": {
    "labels": ["Compliant", "Non-Compliant", "In Progress"],
    "datasets": [{
      "data": [75, 15, 10],
      "backgroundColor": ["#28a745", "#dc3545", "#ffc107"]
    }]
  }
}
```

### Multi-Dataset Chart
```json
{
  "type": "line",
  "data": {
    "labels": ["Week 1", "Week 2", "Week 3", "Week 4"],
    "datasets": [
      {
        "label": "Active Users",
        "data": [450, 520, 485, 510],
        "borderColor": "#0078d4"
      },
      {
        "label": "New Users",
        "data": [45, 52, 38, 61],
        "borderColor": "#00bcf2"
      }
    ]
  }
}
```

---

## 📊 Common Use Cases

### 1. Executive Security Briefing
```json
{
  "action": "create",
  "driveId": "b!...",
  "folderPath": "Executive/Security",
  "fileName": "security_briefing.pptx",
  "title": "Security Status Briefing",
  "template": "executive",
  "slides": [
    {"layout": "title", "title": "Security Overview"},
    {"layout": "chart", "title": "Threat Landscape"},
    {"layout": "comparison", "title": "Risk Assessment"}
  ]
}
```

### 2. Compliance Documentation
```json
{
  "action": "create",
  "driveId": "b!...",
  "folderPath": "Compliance",
  "fileName": "compliance_report.docx",
  "title": "Compliance Report",
  "template": "business",
  "sections": [
    {"type": "heading1", "content": "Compliance Status"},
    {"type": "table", "data": {...}},
    {"type": "heading2", "content": "Recommendations"}
  ]
}
```

### 3. Real-Time Dashboard
```json
{
  "action": "create",
  "driveId": "b!...",
  "folderPath": "Dashboards",
  "fileName": "security_dashboard.html",
  "title": "Security Dashboard",
  "theme": "dashboard",
  "sections": [
    {"type": "grid", "columns": 4, "items": [...]},
    {"type": "chart", "chart": {...}}
  ]
}
```

### 4. Multi-Format Analysis
```json
{
  "reportType": "security-analysis",
  "title": "Comprehensive Security Analysis",
  "dataQueries": [...],
  "formats": ["pptx", "docx", "html", "pdf"],
  "driveId": "b!...",
  "folderPath": "Reports",
  "includeExecutiveSummary": true
}
```

---

## 🔐 Required Permissions

### Application Permissions (Client Credentials)
- `Sites.ReadWrite.All`
- `Files.ReadWrite.All`
- `User.Read.All`
- `DeviceManagementManagedDevices.Read.All`
- `AuditLog.Read.All`
- `SecurityEvents.Read.All`

### Delegated Permissions (OAuth)
- `Files.ReadWrite`
- `Sites.ReadWrite.All`
- `User.Read`
- `offline_access`

---

## ⚡ Quick Tips

1. **Use Professional Reports** for comprehensive multi-format output
2. **HTML Reports** best for interactive dashboards
3. **PowerPoint** ideal for executive presentations
4. **Word Documents** for detailed analysis and documentation
5. **OAuth** required for user-specific file operations
6. Always specify `driveId` - get from Microsoft Graph: `/v1.0/me/drive`
7. Use `folderPath` format: `"Reports/Q4"` (no leading/trailing slashes)
8. File names should include extension: `.pptx`, `.docx`, `.html`
9. Chart.js types: `line`, `bar`, `pie`, `doughnut`, `radar`, `polarArea`
10. For large reports, use pagination in data queries

---

## 🛠️ Environment Setup

```bash
# Required environment variables
MS_TENANT_ID=your-tenant-id
MS_CLIENT_ID=your-client-id
MS_CLIENT_SECRET=your-client-secret

# Optional (for OAuth)
MS_REDIRECT_URI=http://localhost:3000/auth/callback
```

---

## 📚 Additional Resources

- **Complete Documentation**: See DOCUMENT_GENERATION_COMPLETE.md
- **API Reference**: Microsoft Graph Files API
- **Chart.js Docs**: https://www.chartjs.org/docs/
- **Bootstrap Docs**: https://getbootstrap.com/docs/5.3/

---

## 🎯 Quick Test

```bash
# Build project
npm run build

# Start server
npm start

# Test with MCP Inspector
npx @modelcontextprotocol/inspector node build/index.js
```

---

**Need Help?** Check DOCUMENT_GENERATION_COMPLETE.md for comprehensive examples and troubleshooting.
