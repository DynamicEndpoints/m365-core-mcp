// Compliance Framework Management Types
export interface ComplianceFrameworkArgs {
  action: 'list' | 'configure' | 'status' | 'assess' | 'activate' | 'deactivate';
  framework: 'hitrust' | 'iso27001' | 'soc2' | 'cis';
  scope?: string[];
  settings?: Record<string, unknown>;
}

export interface ComplianceControl {
  id: string;
  name: string;
  description: string;
  category: string;
  subcategory?: string;
  controlFamily: string;
  implementationStatus: 'implemented' | 'partiallyImplemented' | 'notImplemented' | 'notApplicable';
  testingStatus: 'passed' | 'failed' | 'notTested' | 'inProgress';
  riskLevel: 'low' | 'medium' | 'high' | 'critical';
  owner: string;
  lastAssessed?: string;
  nextAssessment?: string;
  evidence?: ComplianceEvidence[];
  remediation?: RemediationPlan;
  mappings?: ControlMapping[];
}

export interface AssessmentSettings {
  automaticTesting: boolean;
  testingFrequency: 'daily' | 'weekly' | 'monthly' | 'quarterly' | 'annually';
  evidenceCollection: boolean;
  riskAssessment: boolean;
  complianceThreshold: number; // Percentage
}

export interface ComplianceEvidence {
  id: string;
  type: 'document' | 'screenshot' | 'configuration' | 'log' | 'policy' | 'procedure';
  name: string;
  description: string;
  collectedDate: string;
  source: string;
  automated: boolean;
  filePath?: string;
  url?: string;
  metadata?: Record<string, any>;
}

export interface RemediationPlan {
  status: 'planned' | 'inProgress' | 'completed' | 'deferred';
  priority: 'low' | 'medium' | 'high' | 'critical';
  assignedTo: string;
  dueDate: string;
  estimatedEffort: string;
  description: string;
  tasks: RemediationTask[];
}

export interface RemediationTask {
  id: string;
  description: string;
  status: 'pending' | 'inProgress' | 'completed';
  assignedTo: string;
  dueDate: string;
  completedDate?: string;
  notes?: string;
}

export interface ControlMapping {
  frameworkId: string;
  controlId: string;
  mappingType: 'direct' | 'partial' | 'derived';
  confidence: number; // 0-100
}

// Audit Report Generation Types
export interface AuditReportArgs {
  framework: 'hitrust' | 'iso27001' | 'soc2' | 'cis';
  reportType: 'full' | 'summary' | 'gaps' | 'evidence' | 'executive' | 'control_matrix' | 'risk_assessment';
  dateRange: { 
    startDate: string; 
    endDate: string; 
  };
  format: 'csv' | 'html' | 'pdf' | 'xlsx';
  includeEvidence: boolean;
  outputPath?: string;
  customTemplate?: string;
  filters?: {
    controlIds?: string[];
    riskLevels?: ('low' | 'medium' | 'high' | 'critical')[];
    implementationStatus?: ('implemented' | 'partiallyImplemented' | 'notImplemented' | 'notApplicable')[];
    testingStatus?: ('passed' | 'failed' | 'notTested' | 'inProgress')[];
    owners?: string[];
  };
}

export interface AuditReport {
  id: string;
  framework: string;
  reportType: string;
  generatedDate: string;
  period: {
    startDate: string;
    endDate: string;
  };
  summary: ComplianceSummary;
  controls: ControlAssessment[];
  gaps: ComplianceGap[];
  recommendations: Recommendation[];
  evidence: ComplianceEvidence[];
  metadata: ReportMetadata;
}

export interface ComplianceSummary {
  totalControls: number;
  implementedControls: number;
  partiallyImplementedControls: number;
  notImplementedControls: number;
  notApplicableControls: number;
  compliancePercentage: number;
  riskScore: number;
  lastAssessmentDate: string;
  trendsData?: ComplianceTrend[];
}

export interface ComplianceTrend {
  date: string;
  compliancePercentage: number;
  implementedControls: number;
  riskScore: number;
}

export interface ControlAssessment {
  controlId: string;
  controlName: string;
  category: string;
  implementationStatus: string;
  testingStatus: string;
  riskLevel: string;
  lastTested: string;
  nextAssessment: string;
  owner: string;
  evidenceCount: number;
  findings?: Finding[];
  score?: number;
}

export interface Finding {
  id: string;
  type: 'deficiency' | 'observation' | 'recommendation';
  severity: 'low' | 'medium' | 'high' | 'critical';
  description: string;
  impact: string;
  recommendation: string;
  status: 'open' | 'inProgress' | 'resolved' | 'deferred';
  identifiedDate: string;
  targetResolutionDate?: string;
  actualResolutionDate?: string;
  assignedTo?: string;
}

export interface ComplianceGap {
  controlId: string;
  controlName: string;
  category: string;
  currentStatus: string;
  requiredStatus: string;
  riskLevel: string;
  impact: string;
  recommendedActions: string[];
  estimatedEffort: string;
  priority: number;
}

export interface Recommendation {
  id: string;
  type: 'immediate' | 'shortTerm' | 'longTerm';
  priority: 'low' | 'medium' | 'high' | 'critical';
  title: string;
  description: string;
  impact: string;
  effort: string;
  resources: string[];
  timeline: string;
  relatedControls: string[];
}

export interface ReportMetadata {
  generatedBy: string;
  generationTime: number; // milliseconds
  dataSource: string;
  version: string;
  totalPages?: number;
  fileSize?: number;
  checksum?: string;
}

// Compliance Assessment Types
export interface ComplianceAssessmentArgs {
  action: 'create' | 'update' | 'execute' | 'schedule' | 'cancel' | 'get_results';
  assessmentId?: string;
  framework: 'hitrust' | 'iso27001' | 'soc2';
  scope: Record<string, unknown>;
  settings?: Record<string, unknown>;
}

export interface ComplianceAssessment {
  id: string;
  name: string;
  framework: string;
  status: 'scheduled' | 'running' | 'completed' | 'failed' | 'cancelled';
  createdDate: string;
  scheduledDate?: string;
  startedDate?: string;
  completedDate?: string;
  progress: number; // 0-100
  scope: AssessmentScope;
  results?: AssessmentResults;
  errors?: AssessmentError[];
}

export interface AssessmentScope {
  totalControls: number;
  controlIds: string[];
  categories: string[];
  automated: boolean;
  manualSteps: number;
}

export interface AssessmentResults {
  overallScore: number;
  controlResults: ControlResult[];
  summary: ResultSummary;
  executionTime: number;
  evidenceCollected: number;
}

export interface ControlResult {
  controlId: string;
  status: 'pass' | 'fail' | 'partial' | 'skip' | 'error';
  score: number;
  automated: boolean;
  testDate: string;
  evidence: string[];
  notes?: string;
  errors?: string[];
}

export interface ResultSummary {
  passed: number;
  failed: number;
  partial: number;
  skipped: number;
  errors: number;
}

export interface AssessmentError {
  controlId: string;
  errorType: 'connection' | 'permission' | 'configuration' | 'data' | 'system';
  message: string;
  details?: string;
  timestamp: string;
}

// Compliance Monitoring Types
export interface ComplianceMonitoringArgs {
  action: 'get_status' | 'get_alerts' | 'get_trends' | 'configure_monitoring';
  framework?: 'hitrust' | 'iso27001' | 'soc2';
  filters?: Record<string, unknown>;
  monitoringSettings?: Record<string, unknown>;
}

export interface AlertThreshold {
  metric: 'complianceScore' | 'riskScore' | 'failedControls' | 'overdueAssessments';
  operator: 'greaterThan' | 'lessThan' | 'equals' | 'notEquals';
  value: number;
  severity: 'info' | 'warning' | 'error' | 'critical';
}

export interface NotificationSettings {
  email: boolean;
  teams: boolean;
  webhook?: string;
  recipients: string[];
  escalation?: EscalationSettings;
}

export interface EscalationSettings {
  enabled: boolean;
  delayMinutes: number;
  escalationRecipients: string[];
  maxEscalations: number;
}

export interface ComplianceStatus {
  framework: string;
  overallScore: number;
  riskLevel: 'low' | 'medium' | 'high' | 'critical';
  lastAssessmentDate: string;
  nextAssessmentDate: string;
  activeAlerts: ComplianceAlert[];
  trends: ComplianceTrend[];
  controlSummary: {
    total: number;
    compliant: number;
    nonCompliant: number;
    partiallyCompliant: number;
    notAssessed: number;
  };
}

export interface ComplianceAlert {
  id: string;
  type: 'complianceThreshold' | 'controlFailure' | 'assessmentOverdue' | 'riskIncrease';
  severity: 'info' | 'warning' | 'error' | 'critical';
  title: string;
  description: string;
  affectedControls: string[];
  createdDate: string;
  status: 'active' | 'acknowledged' | 'resolved';
  assignedTo?: string;
  resolvedDate?: string;
  metadata?: Record<string, any>;
}

// Evidence Collection Types
export interface EvidenceCollectionArgs {
  action: 'collect' | 'schedule' | 'get_status' | 'download';
  collectionId?: string;
  framework?: 'hitrust' | 'iso27001' | 'soc2';
  controlIds?: string[];
  evidenceTypes?: ('configuration' | 'logs' | 'policies' | 'screenshots' | 'documents')[];
  settings?: {
    automated: boolean;
    scheduledTime?: string;
    retention: number; // days
    encryption: boolean;
    compression: boolean;
  };
}

export interface EvidenceCollection {
  id: string;
  name: string;
  status: 'scheduled' | 'running' | 'completed' | 'failed';
  framework: string;
  startedDate: string;
  completedDate?: string;
  progress: number;
  totalItems: number;
  collectedItems: number;
  failedItems: number;
  evidence: CollectedEvidence[];
  errors?: string[];
}

export interface CollectedEvidence {
  id: string;
  controlId: string;
  type: string;
  name: string;
  source: string;
  collectedDate: string;
  size: number; // bytes
  format: string;
  encrypted: boolean;
  filePath: string;
  checksum: string;
  metadata: Record<string, any>;
}

// Framework-Specific Types

// HITRUST CSF Types
export interface HITRUSTControl {
  controlId: string; // Format: XX.XX.XX
  controlPoint: string;
  implementationLevel: 1 | 2 | 3;
  maturityLevel: 'Ad hoc' | 'Developing' | 'Defined' | 'Managed' | 'Optimized';
  requirementStatement: string;
  implementationGuidance: string;
  assessmentProcedure: string;
  references: string[];
  controlFamily: string; // e.g., "Information Security Management Program"
}

// ISO 27001 Types
export interface ISO27001Control {
  controlId: string; // Format: A.X.X.X
  controlTitle: string;
  controlObjective: string;
  implementationGuidance: string;
  otherInformation: string;
  domain: string; // e.g., "A.5 Information Security Policies"
  subcategory: string;
  controlType: 'preventive' | 'detective' | 'corrective';
  controlNature: 'management' | 'operational' | 'technical';
}

// SOC 2 Types
export interface SOC2Control {
  controlId: string;
  trustServicesCriteria: 'Security' | 'Availability' | 'ProcessingIntegrity' | 'Confidentiality' | 'PrivacyProtection';
  controlActivity: string;
  controlObjective: string;
  riskRating: 'low' | 'medium' | 'high';
  testingProcedure: string;
  frequency: 'daily' | 'weekly' | 'monthly' | 'quarterly' | 'annually';
  pointInTimeTest: boolean;
  periodOfTimeTest: boolean;
}

// CIS Controls Types
export interface CISControl {
  controlId: string; // Format: CIS-X.X
  title: string;
  description: string;
  assetType: 'Devices' | 'Software' | 'Network' | 'Users' | 'Data';
  securityFunction: 'Identify' | 'Protect' | 'Detect' | 'Respond' | 'Recover';
  implementationGroup: 1 | 2 | 3; // IG1, IG2, IG3
  subControls: CISSubControl[];
  automatable: boolean;
  category: 'Basic Hygiene' | 'Foundational' | 'Organizational';
  nistCsfMapping?: string[];
  references: string[];
}

export interface CISSubControl {
  subControlId: string; // Format: CIS-X.X.X
  title: string;
  description: string;
  implementationGroup: 1 | 2 | 3;
  automatable: boolean;
  difficulty: 'Low' | 'Medium' | 'High';
  testingProcedure: string;
  remediation: string;
  relatedSafeguards?: string[];
}

// CIS Compliance Assessment Types
export interface CISComplianceArgs {
  action: 'assess' | 'get_benchmark' | 'generate_report' | 'configure_monitoring' | 'remediate';
  benchmark?: 'windows-10' | 'windows-11' | 'windows-server-2019' | 'windows-server-2022' | 'office365' | 'azure' | 'intune';
  implementationGroup?: '1' | '2' | '3';
  controlIds?: string[];
  scope?: {
    devices?: string[];
    users?: string[];
    policies?: string[];
  };
  settings?: {
    automated?: boolean;
    generateRemediation?: boolean;
    includeEvidence?: boolean;
    riskPrioritization?: boolean;
  };
}

export interface CISBenchmark {
  id: string;
  name: string;
  version: string;
  platform: string;
  description: string;
  controls: CISControl[];
  profiles: CISProfile[];
  lastUpdated: string;
  applicability: string[];
}

export interface CISProfile {
  name: string;
  description: string;
  level: 1 | 2;
  controls: string[];
  applicableEnvironments: string[];
}

export interface CISAssessmentResult {
  assessmentId: string;
  benchmark: string;
  implementationGroup: number;
  executedDate: string;
  overallScore: number;
  compliancePercentage: number;
  totalControls: number;
  passedControls: number;
  failedControls: number;
  notApplicableControls: number;
  controlResults: CISControlResult[];
  riskSummary: CISRiskSummary;
  remediationPlan: CISRemediationPlan;
}

export interface CISControlResult {
  controlId: string;
  subControlId?: string;
  title: string;
  status: 'Pass' | 'Fail' | 'Not Applicable' | 'Manual Review Required';
  score: number;
  implementationGroup: number;
  automatable: boolean;
  testResult: string;
  evidence: CISEvidence[];
  remediation?: CISRemediation;
  riskLevel: 'Low' | 'Medium' | 'High' | 'Critical';
}

export interface CISEvidence {
  type: 'Registry' | 'Policy' | 'Configuration' | 'Log' | 'Script';
  source: string;
  value: string;
  expected: string;
  collected: string;
  automated: boolean;
}

export interface CISRemediation {
  title: string;
  description: string;
  steps: string[];
  automatable: boolean;
  scriptPath?: string;
  estimatedTime: string;
  difficulty: 'Low' | 'Medium' | 'High';
  impact: 'Low' | 'Medium' | 'High';
  prerequisites: string[];
}

export interface CISRiskSummary {
  criticalFindings: number;
  highRiskFindings: number;
  mediumRiskFindings: number;
  lowRiskFindings: number;
  topRisks: CISRiskFinding[];
  riskTrend: CISRiskTrend[];
}

export interface CISRiskFinding {
  controlId: string;
  title: string;
  riskLevel: 'Low' | 'Medium' | 'High' | 'Critical';
  impact: string;
  likelihood: string;
  affectedAssets: number;
  businessImpact: string;
  technicalImpact: string;
}

export interface CISRiskTrend {
  date: string;
  overallScore: number;
  criticalCount: number;
  highCount: number;
  mediumCount: number;
  lowCount: number;
}

export interface CISRemediationPlan {
  planId: string;
  generatedDate: string;
  totalTasks: number;
  estimatedTime: string;
  phases: CISRemediationPhase[];
  dependencies: CISDependency[];
}

export interface CISRemediationPhase {
  phase: number;
  name: string;
  description: string;
  tasks: CISRemediationTask[];
  startDate: string;
  endDate: string;
  priority: 'Critical' | 'High' | 'Medium' | 'Low';
}

export interface CISRemediationTask {
  taskId: string;
  controlId: string;
  title: string;
  description: string;
  automatable: boolean;
  estimatedTime: string;
  difficulty: 'Low' | 'Medium' | 'High';
  impact: 'Low' | 'Medium' | 'High';
  status: 'Pending' | 'In Progress' | 'Completed' | 'Skipped';
  assignedTo?: string;
  dueDate: string;
  prerequisites: string[];
  validationSteps: string[];
}

export interface CISDependency {
  taskId: string;
  dependsOn: string[];
  type: 'blocking' | 'soft';
  description: string;
}

// Gap Analysis Types
export interface GapAnalysisArgs {
  action: 'generate' | 'get_results' | 'export';
  analysisId?: string;
  framework: 'hitrust' | 'iso27001' | 'soc2';
  targetFramework?: 'hitrust' | 'iso27001' | 'soc2'; // For cross-framework mapping
  scope?: {
    controlIds?: string[];
    categories?: string[];
  };
  settings?: {
    includeRecommendations: boolean;
    prioritizeByRisk: boolean;
    includeTimeline: boolean;
    includeCostEstimate: boolean;
  };
}

export interface GapAnalysisResult {
  id: string;
  framework: string;
  generatedDate: string;
  summary: GapSummary;
  gaps: ControlGap[];
  recommendations: GapRecommendation[];
  timeline: ImplementationTimeline[];
  costEstimate?: CostEstimate;
}

export interface GapSummary {
  totalControls: number;
  compliantControls: number;
  gapControls: number;
  partialControls: number;
  priorityGaps: {
    critical: number;
    high: number;
    medium: number;
    low: number;
  };
  estimatedImplementationTime: string;
}

export interface ControlGap {
  controlId: string;
  controlName: string;
  currentState: string;
  requiredState: string;
  gapType: 'missing' | 'partial' | 'inadequate';
  riskLevel: 'low' | 'medium' | 'high' | 'critical';
  businessImpact: string;
  technicalImpact: string;
  effort: 'low' | 'medium' | 'high';
  dependencies: string[];
  recommendations: string[];
}

export interface GapRecommendation {
  id: string;
  title: string;
  description: string;
  affectedControls: string[];
  implementationType: 'technical' | 'procedural' | 'administrative';
  priority: 'critical' | 'high' | 'medium' | 'low';
  effort: string;
  timeline: string;
  resources: string[];
  dependencies: string[];
  riskReduction: number; // percentage
}

export interface ImplementationTimeline {
  phase: string;
  startDate: string;
  endDate: string;
  controlIds: string[];
  dependencies: string[];
  milestones: Milestone[];
  resources: Resource[];
}

export interface Milestone {
  name: string;
  date: string;
  description: string;
  deliverables: string[];
}

export interface Resource {
  type: 'internal' | 'external' | 'technology' | 'training';
  name: string;
  allocation: string;
  cost?: number;
}

export interface CostEstimate {
  totalCost: number;
  breakdown: CostBreakdown[];
  assumptions: string[];
  contingency: number; // percentage
}

export interface CostBreakdown {
  category: 'personnel' | 'technology' | 'consulting' | 'training' | 'other';
  description: string;
  cost: number;
  recurring: boolean;
  frequency?: 'monthly' | 'quarterly' | 'annually';
}
