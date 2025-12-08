/**
 * Document Generation Type Definitions
 * Types for PowerPoint, Word, and HTML document creation
 */

// PowerPoint Presentation Types
export interface PowerPointPresentationArgs {
  action: 'create' | 'get' | 'list' | 'export';
  
  // For create
  fileName?: string;
  driveId?: string;
  folderId?: string;
  template?: PowerPointTemplate;
  slides?: PowerPointSlide[];
  
  // For get/export
  fileId?: string;
  format?: 'pptx' | 'pdf' | 'odp';
  
  // For list
  filter?: string;
  top?: number;
}

export interface PowerPointTemplate {
  theme?: 'light' | 'dark' | 'colorful' | 'professional';
  masterSlideLayout?: string;
  companyLogo?: string;
  companyName?: string;
  footer?: string;
}

export interface PowerPointSlide {
  type: 'title' | 'title-content' | 'two-column' | 'comparison' | 'chart' | 'image' | 'blank';
  title?: string;
  subtitle?: string;
  content?: string[];
  leftContent?: string;
  rightContent?: string;
  imageUrl?: string;
  chartData?: ChartData;
  notes?: string;
}

export interface ChartData {
  type: 'bar' | 'line' | 'pie' | 'column' | 'area';
  title: string;
  categories: string[];
  series: {
    name: string;
    values: number[];
  }[];
}

// Word Document Types
export interface WordDocumentArgs {
  action: 'create' | 'get' | 'list' | 'export' | 'append';
  
  // For create
  fileName?: string;
  driveId?: string;
  folderId?: string;
  template?: WordTemplate;
  sections?: WordSection[];
  
  // For get/export
  fileId?: string;
  format?: 'docx' | 'pdf' | 'html' | 'txt';
  
  // For append
  content?: string;
  
  // For list
  filter?: string;
  top?: number;
}

export interface WordTemplate {
  style?: 'business' | 'academic' | 'report' | 'memo';
  header?: string;
  footer?: string;
  pageNumbers?: boolean;
  tableOfContents?: boolean;
  companyLogo?: string;
  watermark?: string;
}

export interface WordSection {
  type: 'heading1' | 'heading2' | 'heading3' | 'paragraph' | 'table' | 'list' | 'image' | 'pageBreak';
  content?: string;
  items?: string[];
  tableData?: TableData;
  imageUrl?: string;
  style?: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string;
    fontSize?: number;
    alignment?: 'left' | 'center' | 'right' | 'justify';
  };
}

export interface TableData {
  headers: string[];
  rows: string[][];
  style?: 'plain' | 'striped' | 'bordered';
}

// HTML Report Types
export interface HTMLReportArgs {
  action: 'create' | 'get' | 'list';
  
  // For create
  fileName?: string;
  driveId?: string;
  folderId?: string;
  template?: HTMLTemplate;
  sections?: HTMLSection[];
  includeCharts?: boolean;
  
  // For get
  fileId?: string;
  
  // For list
  filter?: string;
  top?: number;
}

export interface HTMLTemplate {
  theme?: 'modern' | 'classic' | 'minimal' | 'dashboard';
  title: string;
  description?: string;
  author?: string;
  companyName?: string;
  companyLogo?: string;
  customCSS?: string;
  includeBootstrap?: boolean;
  includeChartJS?: boolean;
}

export interface HTMLSection {
  type: 'heading' | 'paragraph' | 'table' | 'chart' | 'list' | 'grid' | 'card' | 'alert';
  heading?: string;
  content?: string;
  level?: number; // 1-6 heading level
  items?: string[];
  tableData?: TableData;
  chartData?: ChartData;
  gridColumns?: number;
  cardData?: {
    title: string;
    value: string | number;
    icon?: string;
    color?: string;
  }[];
  alertType?: 'info' | 'success' | 'warning' | 'danger';
}

// Professional Report Types (combines all document types)
export interface ProfessionalReportArgs {
  reportType: 'security-analysis' | 'compliance-audit' | 'user-activity' | 'device-health' | 'custom';
  title: string;
  description?: string;
  
  // Data sources
  dataQueries?: DataQuery[];
  includeCharts?: boolean;
  includeTables?: boolean;
  includeSummary?: boolean;
  
  // Output formats
  outputFormats: ('pptx' | 'docx' | 'html' | 'pdf')[];
  
  // Storage location
  driveId?: string;
  folderId?: string;
  fileNamePrefix?: string;
  
  // Styling
  template?: {
    theme?: string;
    companyName?: string;
    companyLogo?: string;
    primaryColor?: string;
    secondaryColor?: string;
  };
}

export interface DataQuery {
  source: 'users' | 'devices' | 'groups' | 'audit-logs' | 'alerts' | 'policies' | 'compliance';
  endpoint: string;
  filter?: string;
  select?: string[];
  transform?: 'count' | 'group-by' | 'aggregate' | 'trend';
  label: string;
}

// OAuth Types
export interface OAuthConfig {
  clientId: string;
  clientSecret: string;
  tenantId: string;
  redirectUri: string;
  scopes: string[];
}

export interface OAuthTokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  token_type: string;
  scope: string;
}

export interface OAuthAuthorizationArgs {
  action: 'get-auth-url' | 'exchange-code' | 'refresh-token' | 'revoke';
  
  // For get-auth-url
  scopes?: string[];
  state?: string;
  
  // For exchange-code
  code?: string;
  
  // For refresh-token
  refreshToken?: string;
}
