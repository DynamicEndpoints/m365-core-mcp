/**
 * Document Generation Zod Schemas
 * Validation schemas for PowerPoint, Word, and HTML generation tools
 */

import { z } from 'zod';

// Chart Data Schema (shared)
export const chartDataSchema = z.object({
  type: z.enum(['bar', 'line', 'pie', 'column', 'area']).describe('Type of chart to create'),
  title: z.string().describe('Chart title'),
  categories: z.array(z.string()).describe('Category labels for the chart'),
  series: z.array(z.object({
    name: z.string().describe('Series name'),
    values: z.array(z.number()).describe('Data values')
  })).describe('Data series for the chart')
});

// Table Data Schema (shared)
export const tableDataSchema = z.object({
  headers: z.array(z.string()).describe('Table column headers'),
  rows: z.array(z.array(z.string())).describe('Table row data'),
  style: z.enum(['plain', 'striped', 'bordered']).optional().describe('Table styling')
});

// PowerPoint Schemas
export const powerPointSlideSchema = z.object({
  type: z.enum(['title', 'title-content', 'two-column', 'comparison', 'chart', 'image', 'blank'])
    .describe('Slide layout type'),
  title: z.string().optional().describe('Slide title'),
  subtitle: z.string().optional().describe('Slide subtitle (for title slides)'),
  content: z.array(z.string()).optional().describe('Bullet points or content items'),
  leftContent: z.string().optional().describe('Left column content (for two-column layout)'),
  rightContent: z.string().optional().describe('Right column content (for two-column layout)'),
  imageUrl: z.string().optional().describe('Image URL for image slides'),
  chartData: chartDataSchema.optional().describe('Chart data for chart slides'),
  notes: z.string().optional().describe('Speaker notes for the slide')
});

export const powerPointTemplateSchema = z.object({
  theme: z.enum(['light', 'dark', 'colorful', 'professional']).optional().describe('Presentation theme'),
  masterSlideLayout: z.string().optional().describe('Master slide layout name'),
  companyLogo: z.string().optional().describe('URL to company logo'),
  companyName: z.string().optional().describe('Company name for branding'),
  footer: z.string().optional().describe('Footer text for all slides')
});

export const powerPointPresentationArgsSchema = z.object({
  action: z.enum(['create', 'get', 'list', 'export'])
    .describe('Action to perform: create new presentation, get existing, list all, or export to format'),
  fileName: z.string().optional()
    .describe('Name for the new presentation file (for create action)'),
  driveId: z.string().optional()
    .describe('OneDrive/SharePoint drive ID where file should be created (default: user\'s OneDrive)'),
  folderId: z.string().optional()
    .describe('Folder ID within the drive (default: root)'),
  template: powerPointTemplateSchema.optional()
    .describe('Template configuration for presentation styling'),
  slides: z.array(powerPointSlideSchema).optional()
    .describe('Array of slide definitions to create'),
  fileId: z.string().optional()
    .describe('File ID for get/export actions'),
  format: z.enum(['pptx', 'pdf', 'odp']).optional()
    .describe('Export format (for export action)'),
  filter: z.string().optional()
    .describe('OData filter for list action'),
  top: z.number().optional()
    .describe('Number of results to return (for list action)')
});

// Word Document Schemas
export const wordSectionSchema = z.object({
  type: z.enum(['heading1', 'heading2', 'heading3', 'paragraph', 'table', 'list', 'image', 'pageBreak'])
    .describe('Content type for this section'),
  content: z.string().optional().describe('Text content for headings/paragraphs'),
  items: z.array(z.string()).optional().describe('List items (for list type)'),
  tableData: tableDataSchema.optional().describe('Table data (for table type)'),
  imageUrl: z.string().optional().describe('Image URL (for image type)'),
  style: z.object({
    bold: z.boolean().optional().describe('Apply bold formatting'),
    italic: z.boolean().optional().describe('Apply italic formatting'),
    underline: z.boolean().optional().describe('Apply underline formatting'),
    color: z.string().optional().describe('Text color (hex code)'),
    fontSize: z.number().optional().describe('Font size in points'),
    alignment: z.enum(['left', 'center', 'right', 'justify']).optional().describe('Text alignment')
  }).optional().describe('Text styling options')
});

export const wordTemplateSchema = z.object({
  style: z.enum(['business', 'academic', 'report', 'memo']).optional()
    .describe('Document template style'),
  header: z.string().optional().describe('Header text'),
  footer: z.string().optional().describe('Footer text'),
  pageNumbers: z.boolean().optional().describe('Include page numbers'),
  tableOfContents: z.boolean().optional().describe('Generate table of contents'),
  companyLogo: z.string().optional().describe('URL to company logo for header'),
  watermark: z.string().optional().describe('Watermark text')
});

export const wordDocumentArgsSchema = z.object({
  action: z.enum(['create', 'get', 'list', 'export', 'append'])
    .describe('Action: create new document, get existing, list all, export to format, or append content'),
  fileName: z.string().optional()
    .describe('Name for the new document file (for create action)'),
  driveId: z.string().optional()
    .describe('OneDrive/SharePoint drive ID (default: user\'s OneDrive)'),
  folderId: z.string().optional()
    .describe('Folder ID within the drive (default: root)'),
  template: wordTemplateSchema.optional()
    .describe('Template configuration for document styling'),
  sections: z.array(wordSectionSchema).optional()
    .describe('Array of content sections to create'),
  fileId: z.string().optional()
    .describe('File ID for get/export/append actions'),
  format: z.enum(['docx', 'pdf', 'html', 'txt']).optional()
    .describe('Export format (for export action)'),
  content: z.string().optional()
    .describe('Content to append (for append action)'),
  filter: z.string().optional()
    .describe('OData filter for list action'),
  top: z.number().optional()
    .describe('Number of results to return (for list action)')
});

// HTML Report Schemas
export const htmlSectionSchema = z.object({
  type: z.enum(['heading', 'paragraph', 'table', 'chart', 'list', 'grid', 'card', 'alert'])
    .describe('HTML section type'),
  heading: z.string().optional().describe('Section heading text'),
  content: z.string().optional().describe('Section content (HTML allowed)'),
  level: z.number().min(1).max(6).optional().describe('Heading level 1-6 (for heading type)'),
  items: z.array(z.string()).optional().describe('List items (for list type)'),
  tableData: tableDataSchema.optional().describe('Table data (for table type)'),
  chartData: chartDataSchema.optional().describe('Chart data (for chart type)'),
  gridColumns: z.number().optional().describe('Number of columns (for grid type)'),
  cardData: z.array(z.object({
    title: z.string().describe('Card title'),
    value: z.union([z.string(), z.number()]).describe('Card value'),
    icon: z.string().optional().describe('Card icon name'),
    color: z.string().optional().describe('Card color (hex code)')
  })).optional().describe('Card data (for card type)'),
  alertType: z.enum(['info', 'success', 'warning', 'danger']).optional()
    .describe('Alert type (for alert type)')
});

export const htmlTemplateSchema = z.object({
  theme: z.enum(['modern', 'classic', 'minimal', 'dashboard']).optional()
    .describe('HTML report theme'),
  title: z.string().describe('Report title'),
  description: z.string().optional().describe('Report description'),
  author: z.string().optional().describe('Report author'),
  companyName: z.string().optional().describe('Company name for branding'),
  companyLogo: z.string().optional().describe('URL to company logo'),
  customCSS: z.string().optional().describe('Custom CSS to inject'),
  includeBootstrap: z.boolean().optional().describe('Include Bootstrap CSS framework'),
  includeChartJS: z.boolean().optional().describe('Include Chart.js for interactive charts')
});

export const htmlReportArgsSchema = z.object({
  action: z.enum(['create', 'get', 'list'])
    .describe('Action: create new HTML report, get existing, or list all'),
  fileName: z.string().optional()
    .describe('Name for the HTML file (for create action)'),
  driveId: z.string().optional()
    .describe('OneDrive/SharePoint drive ID (default: user\'s OneDrive)'),
  folderId: z.string().optional()
    .describe('Folder ID within the drive (default: root)'),
  template: htmlTemplateSchema.optional()
    .describe('Template configuration for HTML report styling'),
  sections: z.array(htmlSectionSchema).optional()
    .describe('Array of HTML sections to create'),
  includeCharts: z.boolean().optional()
    .describe('Enable interactive charts with Chart.js'),
  fileId: z.string().optional()
    .describe('File ID for get action'),
  filter: z.string().optional()
    .describe('OData filter for list action'),
  top: z.number().optional()
    .describe('Number of results to return (for list action)')
});

// Professional Report Schema (combines multiple formats)
export const dataQuerySchema = z.object({
  source: z.enum(['users', 'devices', 'groups', 'audit-logs', 'alerts', 'policies', 'compliance'])
    .describe('Data source to query'),
  endpoint: z.string().describe('Microsoft Graph API endpoint to call'),
  filter: z.string().optional().describe('OData filter query'),
  select: z.array(z.string()).optional().describe('Fields to select'),
  transform: z.enum(['count', 'group-by', 'aggregate', 'trend']).optional()
    .describe('Data transformation to apply'),
  label: z.string().describe('Label for this data in the report')
});

export const professionalReportArgsSchema = z.object({
  reportType: z.enum(['security-analysis', 'compliance-audit', 'user-activity', 'device-health', 'custom'])
    .describe('Type of professional report to generate'),
  title: z.string().describe('Report title'),
  description: z.string().optional().describe('Report description'),
  dataQueries: z.array(dataQuerySchema).optional()
    .describe('Data queries to execute and include in report'),
  includeCharts: z.boolean().optional()
    .describe('Include visual charts in the report'),
  includeTables: z.boolean().optional()
    .describe('Include data tables in the report'),
  includeSummary: z.boolean().optional()
    .describe('Include executive summary'),
  outputFormats: z.array(z.enum(['pptx', 'docx', 'html', 'pdf']))
    .describe('Output formats to generate (can select multiple)'),
  driveId: z.string().optional()
    .describe('OneDrive/SharePoint drive ID for saving reports'),
  folderId: z.string().optional()
    .describe('Folder ID within the drive'),
  fileNamePrefix: z.string().optional()
    .describe('Prefix for generated file names'),
  template: z.object({
    theme: z.string().optional(),
    companyName: z.string().optional(),
    companyLogo: z.string().optional(),
    primaryColor: z.string().optional().describe('Primary brand color (hex)'),
    secondaryColor: z.string().optional().describe('Secondary brand color (hex)')
  }).optional().describe('Report branding and styling')
});

// OAuth Authorization Schema
export const oauthAuthorizationArgsSchema = z.object({
  action: z.enum(['get-auth-url', 'exchange-code', 'refresh-token', 'revoke'])
    .describe('OAuth action: get authorization URL, exchange code for token, refresh token, or revoke access'),
  scopes: z.array(z.string()).optional()
    .describe('OAuth scopes to request (e.g., Files.ReadWrite, Sites.ReadWrite.All)'),
  state: z.string().optional()
    .describe('State parameter for CSRF protection'),
  code: z.string().optional()
    .describe('Authorization code to exchange for access token'),
  refreshToken: z.string().optional()
    .describe('Refresh token to exchange for new access token')
});
