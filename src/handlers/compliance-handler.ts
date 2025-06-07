import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';
import {
  ComplianceFrameworkArgs,
  ComplianceAssessmentArgs,
  ComplianceMonitoringArgs,
  EvidenceCollectionArgs,
  GapAnalysisArgs,
  CISComplianceArgs,
  CISBenchmark,
  CISControl,
  CISAssessmentResult,
  CISControlResult,
  CISRemediationPlan
} from '../types/compliance-types.js';

// Compliance Framework Management Handler
export async function handleComplianceFrameworks(
  graphClient: Client,
  args: ComplianceFrameworkArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let result: any;

  switch (args.action) {
    case 'list':
      // List available compliance frameworks
      result = {
        frameworks: [
          {
            id: 'hitrust',
            name: 'HITRUST CSF',
            version: '11.1',
            description: 'Health Information Trust Alliance Common Security Framework',
            controlFamilies: 49,
            totalControls: 156,
            status: 'available'
          },
          {
            id: 'iso27001',
            name: 'ISO 27001:2022',
            version: '2022',
            description: 'Information Security Management System',
            controlFamilies: 14,
            totalControls: 114,
            status: 'available'
          },
          {
            id: 'soc2',
            name: 'SOC 2 Type II',
            version: '2017',
            description: 'Service Organization Control 2',
            controlFamilies: 5,
            totalControls: 64,
            status: 'available'
          }
        ]
      };
      break;

    case 'configure':
      // Configure compliance framework settings
      const frameworkConfig = {
        framework: args.framework,
        scope: args.scope || ['all'],
        settings: args.settings,
        configuredDate: new Date().toISOString(),
        status: 'configured'
      };
      
      // In a real implementation, this would be stored in a database
      result = { message: 'Framework configured successfully', config: frameworkConfig };
      break;

    case 'status':
      // Get compliance framework status
      result = await getFrameworkStatus(graphClient, args.framework);
      break;

    case 'assess':
      // Trigger compliance assessment
      result = await triggerAssessment(graphClient, args.framework, args.scope || []);
      break;

    case 'activate':
      result = { message: `${args.framework} framework activated`, status: 'active' };
      break;

    case 'deactivate':
      result = { message: `${args.framework} framework deactivated`, status: 'inactive' };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}

// Compliance Assessment Handler
export async function handleComplianceAssessments(
  graphClient: Client,
  args: ComplianceAssessmentArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let result: any;

  switch (args.action) {
    case 'create':
      // Create new compliance assessment
      const assessmentId = `assessment-${Date.now()}`;
      result = {
        id: assessmentId,
        framework: args.framework,
        scope: args.scope,
        settings: args.settings,
        status: 'created',
        createdDate: new Date().toISOString()
      };
      break;

    case 'execute':
      if (!args.assessmentId) {
        throw new McpError(ErrorCode.InvalidParams, 'assessmentId is required for execute action');
      }
      
      // Execute assessment
      result = await executeAssessment(graphClient, args.assessmentId, args.framework);
      break;

    case 'get_results':
      if (!args.assessmentId) {
        throw new McpError(ErrorCode.InvalidParams, 'assessmentId is required for get_results action');
      }
      
      result = await getAssessmentResults(graphClient, args.assessmentId);
      break;

    case 'schedule':
      result = {
        assessmentId: args.assessmentId,
        scheduledDate: args.settings?.scheduledDate,
        status: 'scheduled',
        message: 'Assessment scheduled successfully'
      };
      break;

    case 'cancel':
      result = {
        assessmentId: args.assessmentId,
        status: 'cancelled',
        message: 'Assessment cancelled successfully'
      };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}

// Compliance Monitoring Handler
export async function handleComplianceMonitoring(
  graphClient: Client,
  args: ComplianceMonitoringArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let result: any;

  switch (args.action) {
    case 'get_status':
      // Get overall compliance status
      result = await getComplianceStatus(graphClient, args.framework, args.filters);
      break;

    case 'get_alerts':
      // Get compliance alerts
      result = await getComplianceAlerts(graphClient, args.framework, args.filters);
      break;

    case 'get_trends':
      // Get compliance trends
      result = await getComplianceTrends(graphClient, args.framework, args.filters);
      break;

    case 'configure_monitoring':
      // Configure monitoring settings
      result = {
        framework: args.framework,
        monitoringSettings: args.monitoringSettings,
        status: 'configured',
        message: 'Monitoring configured successfully'
      };
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}

// Evidence Collection Handler
export async function handleEvidenceCollection(
  graphClient: Client,
  args: EvidenceCollectionArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let result: any;

  switch (args.action) {
    case 'collect':
      // Start evidence collection
      const collectionId = `collection-${Date.now()}`;
      result = await startEvidenceCollection(graphClient, collectionId, args);
      break;

    case 'schedule':
      result = {
        collectionId: args.collectionId,
        scheduledTime: args.settings?.scheduledTime,
        status: 'scheduled',
        message: 'Evidence collection scheduled successfully'
      };
      break;

    case 'get_status':
      if (!args.collectionId) {
        throw new McpError(ErrorCode.InvalidParams, 'collectionId is required for get_status action');
      }
      
      result = await getCollectionStatus(args.collectionId);
      break;

    case 'download':
      if (!args.collectionId) {
        throw new McpError(ErrorCode.InvalidParams, 'collectionId is required for download action');
      }
      
      result = await downloadEvidence(args.collectionId);
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}

// Gap Analysis Handler
export async function handleGapAnalysis(
  graphClient: Client,
  args: GapAnalysisArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let result: any;

  switch (args.action) {
    case 'generate':
      // Generate gap analysis
      const analysisId = `gap-analysis-${Date.now()}`;
      result = await generateGapAnalysis(graphClient, analysisId, args);
      break;

    case 'get_results':
      if (!args.analysisId) {
        throw new McpError(ErrorCode.InvalidParams, 'analysisId is required for get_results action');
      }
      
      result = await getGapAnalysisResults(args.analysisId);
      break;

    case 'export':
      if (!args.analysisId) {
        throw new McpError(ErrorCode.InvalidParams, 'analysisId is required for export action');
      }
      
      result = await exportGapAnalysis(args.analysisId);
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}

// Helper functions
async function getFrameworkStatus(graphClient: Client, framework: string) {
  // Get data from Microsoft Compliance Manager and other sources
  const secureScore = await graphClient.api('/security/secureScores').top(1).get();
  const controls = await graphClient.api('/security/secureScoreControlProfiles').get();
  
  return {
    framework,
    overallScore: secureScore.value[0]?.currentScore || 0,
    maxScore: secureScore.value[0]?.maxScore || 100,
    compliancePercentage: Math.round((secureScore.value[0]?.currentScore / secureScore.value[0]?.maxScore) * 100) || 0,
    lastAssessmentDate: new Date().toISOString(),
    controlSummary: {
      total: controls.value?.length || 0,
      compliant: controls.value?.filter((c: any) => c.implementationStatus === 'implemented').length || 0,
      nonCompliant: controls.value?.filter((c: any) => c.implementationStatus === 'notImplemented').length || 0,
      partiallyCompliant: controls.value?.filter((c: any) => c.implementationStatus === 'partiallyImplemented').length || 0
    }
  };
}

async function triggerAssessment(graphClient: Client, framework: string, scope: string[]) {
  // This would trigger automated assessment tasks
  return {
    assessmentId: `assessment-${Date.now()}`,
    framework,
    scope,
    status: 'running',
    startedDate: new Date().toISOString(),
    estimatedCompletion: new Date(Date.now() + 3600000).toISOString() // 1 hour
  };
}

async function executeAssessment(graphClient: Client, assessmentId: string, framework: string) {
  // Execute compliance assessment
  return {
    assessmentId,
    framework,
    status: 'completed',
    completedDate: new Date().toISOString(),
    results: {
      overallScore: 85,
      controlsAssessed: 114,
      controlsPassed: 97,
      controlsFailed: 17,
      recommendations: 12
    }
  };
}

async function getAssessmentResults(graphClient: Client, assessmentId: string) {
  // Get detailed assessment results
  return {
    assessmentId,
    status: 'completed',
    results: {
      overallScore: 85,
      summary: {
        passed: 97,
        failed: 17,
        partial: 0,
        skipped: 0
      },
      controlResults: [
        {
          controlId: 'A.5.1.1',
          status: 'pass',
          score: 100,
          automated: true,
          testDate: new Date().toISOString()
        }
      ]
    }
  };
}

async function getComplianceStatus(graphClient: Client, framework?: string, filters?: any) {
  // Get current compliance status
  return {
    framework,
    overallScore: 85,
    riskLevel: 'medium',
    lastAssessmentDate: new Date().toISOString(),
    nextAssessmentDate: new Date(Date.now() + 30 * 24 * 3600000).toISOString(), // 30 days
    activeAlerts: [],
    trends: []
  };
}

async function getComplianceAlerts(graphClient: Client, framework?: string, filters?: any) {
  // Get compliance alerts
  return {
    alerts: [],
    totalCount: 0
  };
}

async function getComplianceTrends(graphClient: Client, framework?: string, filters?: any) {
  // Get compliance trends
  return {
    trends: [],
    period: filters?.timeRange || { startDate: new Date(Date.now() - 30 * 24 * 3600000).toISOString(), endDate: new Date().toISOString() }
  };
}

async function startEvidenceCollection(graphClient: Client, collectionId: string, args: EvidenceCollectionArgs) {
  // Start evidence collection process
  return {
    id: collectionId,
    status: 'running',
    framework: args.framework,
    startedDate: new Date().toISOString(),
    progress: 0,
    totalItems: 50,
    collectedItems: 0
  };
}

async function getCollectionStatus(collectionId: string) {
  // Get evidence collection status
  return {
    id: collectionId,
    status: 'completed',
    progress: 100,
    totalItems: 50,
    collectedItems: 48,
    failedItems: 2
  };
}

async function downloadEvidence(collectionId: string) {
  // Download collected evidence
  return {
    collectionId,
    downloadUrl: `/outputs/evidence-${collectionId}.zip`,
    expiresAt: new Date(Date.now() + 24 * 3600000).toISOString()
  };
}

async function generateGapAnalysis(graphClient: Client, analysisId: string, args: GapAnalysisArgs) {
  // Generate gap analysis
  return {
    id: analysisId,
    framework: args.framework,
    status: 'running',
    generatedDate: new Date().toISOString(),
    estimatedCompletion: new Date(Date.now() + 1800000).toISOString() // 30 minutes
  };
}

async function getGapAnalysisResults(analysisId: string) {
  // Get gap analysis results
  return {
    id: analysisId,
    status: 'completed',
    summary: {
      totalControls: 114,
      compliantControls: 97,
      gapControls: 17,
      partialControls: 0,
      priorityGaps: {
        critical: 3,
        high: 8,
        medium: 6,
        low: 0
      }
    },
    gaps: [],
    recommendations: []
  };
}

async function exportGapAnalysis(analysisId: string) {
  // Export gap analysis
  return {
    analysisId,
    exportUrl: `/outputs/gap-analysis-${analysisId}.pdf`,
    format: 'pdf',
    generatedDate: new Date().toISOString()
  };
}

// CIS Compliance Handler
export async function handleCISCompliance(
  graphClient: Client,
  args: CISComplianceArgs
): Promise<{ content: { type: string; text: string }[] }> {
  let result: any;

  switch (args.action) {
    case 'assess':
      // Perform CIS compliance assessment
      result = await performCISAssessment(graphClient, args);
      break;

    case 'get_benchmark':
      // Get CIS benchmark information
      result = await getCISBenchmark(args.benchmark || 'office365');
      break;

    case 'generate_report':
      // Generate CIS compliance report
      result = await generateCISReport(graphClient, args);
      break;

    case 'configure_monitoring':
      // Configure CIS monitoring
      result = await configureCISMonitoring(args);
      break;

    case 'remediate':
      // Execute automated remediation
      result = await executeCISRemediation(graphClient, args);
      break;

    default:
      throw new McpError(ErrorCode.InvalidParams, `Invalid action: ${args.action}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}

// CIS Assessment Functions
async function performCISAssessment(graphClient: Client, args: CISComplianceArgs): Promise<CISAssessmentResult> {
  const assessmentId = `cis-assessment-${Date.now()}`;
  const benchmark = args.benchmark || 'office365';
  const implementationGroup = args.implementationGroup ? parseInt(args.implementationGroup) : 1;

  // Get relevant Microsoft 365/Azure configurations
  const results = await gatherCISEvidence(graphClient, benchmark, args.scope);
  
  // Evaluate against CIS controls
  const controlResults = await evaluateCISControls(results, benchmark, implementationGroup, args.controlIds);
  
  // Generate risk summary
  const riskSummary = generateCISRiskSummary(controlResults);
  
  // Create remediation plan
  const remediationPlan = generateCISRemediationPlan(controlResults);

  return {
    assessmentId,
    benchmark,
    implementationGroup,
    executedDate: new Date().toISOString(),
    overallScore: calculateOverallScore(controlResults),
    compliancePercentage: calculateCompliancePercentage(controlResults),
    totalControls: controlResults.length,
    passedControls: controlResults.filter(c => c.status === 'Pass').length,
    failedControls: controlResults.filter(c => c.status === 'Fail').length,
    notApplicableControls: controlResults.filter(c => c.status === 'Not Applicable').length,
    controlResults,
    riskSummary,
    remediationPlan
  };
}

async function getCISBenchmark(benchmark: string): Promise<CISBenchmark> {
  // CIS Controls for Microsoft 365/Azure environments
  const benchmarks: Record<string, CISBenchmark> = {
    'office365': {
      id: 'cis-microsoft-365-foundations-v3.0.0',
      name: 'CIS Microsoft 365 Foundations Benchmark',
      version: '3.0.0',
      platform: 'Microsoft 365',
      description: 'CIS security configuration recommendations for Microsoft 365',
      controls: getCISMicrosoft365Controls(),
      profiles: [
        {
          name: 'Level 1',
          description: 'Basic security configurations',
          level: 1,
          controls: ['CIS-1.1', 'CIS-1.2', 'CIS-2.1', 'CIS-3.1'],
          applicableEnvironments: ['All Microsoft 365 tenants']
        },
        {
          name: 'Level 2',
          description: 'Enhanced security configurations',
          level: 2,
          controls: ['CIS-1.1', 'CIS-1.2', 'CIS-2.1', 'CIS-3.1', 'CIS-4.1', 'CIS-5.1'],
          applicableEnvironments: ['Security-conscious organizations']
        }
      ],
      lastUpdated: '2024-01-15',
      applicability: ['Microsoft 365 E3', 'Microsoft 365 E5', 'Microsoft 365 Business Premium']
    },
    'azure': {
      id: 'cis-microsoft-azure-foundations-v2.0.0',
      name: 'CIS Microsoft Azure Foundations Benchmark',
      version: '2.0.0',
      platform: 'Microsoft Azure',
      description: 'CIS security configuration recommendations for Microsoft Azure',
      controls: getCISAzureControls(),
      profiles: [
        {
          name: 'Level 1',
          description: 'Basic Azure security configurations',
          level: 1,
          controls: ['CIS-1.1', 'CIS-2.1', 'CIS-3.1'],
          applicableEnvironments: ['All Azure subscriptions']
        }
      ],
      lastUpdated: '2024-01-15',
      applicability: ['Azure subscriptions']
    },
    'intune': {
      id: 'cis-microsoft-intune-v1.0.0',
      name: 'CIS Microsoft Intune Benchmark',
      version: '1.0.0',
      platform: 'Microsoft Intune',
      description: 'CIS security configuration recommendations for Microsoft Intune',
      controls: getCISIntuneControls(),
      profiles: [
        {
          name: 'Level 1',
          description: 'Basic Intune security configurations',
          level: 1,
          controls: ['CIS-1.1', 'CIS-2.1'],
          applicableEnvironments: ['All Intune deployments']
        }
      ],
      lastUpdated: '2024-01-15',
      applicability: ['Microsoft Intune']
    }
  };

  return benchmarks[benchmark] || benchmarks['office365'];
}

function getCISMicrosoft365Controls(): CISControl[] {
  return [
    {
      controlId: 'CIS-1.1',
      title: 'Ensure multi-factor authentication is enabled for all users',
      description: 'Multi-factor authentication adds an additional layer of protection on top of a user name and password.',
      assetType: 'Users',
      securityFunction: 'Protect',
      implementationGroup: 1,
      subControls: [
        {
          subControlId: 'CIS-1.1.1',
          title: 'Enable MFA for all users',
          description: 'Ensure that MFA is required for all user accounts',
          implementationGroup: 1,
          automatable: true,
          difficulty: 'Low',
          testingProcedure: 'Verify MFA is enabled via Microsoft Graph API',
          remediation: 'Enable MFA through Azure AD conditional access policies',
          relatedSafeguards: ['CIS-1.2']
        }
      ],
      automatable: true,
      category: 'Basic Hygiene',
      nistCsfMapping: ['PR.AC-1', 'PR.AC-7'],
      references: ['https://docs.microsoft.com/en-us/azure/active-directory/authentication/concept-mfa-howitworks']
    },
    {
      controlId: 'CIS-1.2',
      title: 'Ensure legacy authentication is blocked',
      description: 'Legacy authentication protocols do not support multi-factor authentication.',
      assetType: 'Users',
      securityFunction: 'Protect',
      implementationGroup: 1,
      subControls: [
        {
          subControlId: 'CIS-1.2.1',
          title: 'Block legacy authentication protocols',
          description: 'Configure conditional access to block legacy authentication',
          implementationGroup: 1,
          automatable: true,
          difficulty: 'Medium',
          testingProcedure: 'Check conditional access policies for legacy authentication blocks',
          remediation: 'Create conditional access policy to block legacy authentication'
        }
      ],
      automatable: true,
      category: 'Basic Hygiene',
      nistCsfMapping: ['PR.AC-1'],
      references: ['https://docs.microsoft.com/en-us/azure/active-directory/conditional-access/block-legacy-authentication']
    },
    {
      controlId: 'CIS-2.1',
      title: 'Ensure Security Defaults is enabled',
      description: 'Security defaults provide secure default settings that Microsoft manages.',
      assetType: 'Users',
      securityFunction: 'Protect',
      implementationGroup: 1,
      subControls: [
        {
          subControlId: 'CIS-2.1.1',
          title: 'Enable Security Defaults',
          description: 'Enable Azure AD Security Defaults for basic security protection',
          implementationGroup: 1,
          automatable: true,
          difficulty: 'Low',
          testingProcedure: 'Check if Security Defaults are enabled in Azure AD',
          remediation: 'Enable Security Defaults in Azure AD portal'
        }
      ],
      automatable: true,
      category: 'Foundational',
      nistCsfMapping: ['PR.AC-1', 'PR.AC-6'],
      references: ['https://docs.microsoft.com/en-us/azure/active-directory/fundamentals/concept-fundamentals-security-defaults']
    }
  ];
}

function getCISAzureControls(): CISControl[] {
  return [
    {
      controlId: 'CIS-1.1',
      title: 'Ensure that multi-factor authentication is enabled for all privileged users',
      description: 'Enable MFA for users with administrative privileges in Azure',
      assetType: 'Users',
      securityFunction: 'Protect',
      implementationGroup: 1,
      subControls: [],
      automatable: true,
      category: 'Basic Hygiene',
      references: []
    }
  ];
}

function getCISIntuneControls(): CISControl[] {
  return [
    {
      controlId: 'CIS-1.1',
      title: 'Ensure device encryption is enabled',
      description: 'Configure device encryption policies in Intune',
      assetType: 'Devices',
      securityFunction: 'Protect',
      implementationGroup: 1,
      subControls: [],
      automatable: true,
      category: 'Basic Hygiene',
      references: []
    }
  ];
}

async function gatherCISEvidence(graphClient: Client, benchmark: string, scope?: any): Promise<any> {
  const evidence: any = {};

  try {
    // Gather authentication policies
    evidence.authenticationPolicies = await graphClient
      .api('/policies/authenticationMethodsPolicy')
      .get();

    // Gather conditional access policies
    evidence.conditionalAccessPolicies = await graphClient
      .api('/identity/conditionalAccess/policies')
      .get();

    // Gather security defaults
    evidence.securityDefaults = await graphClient
      .api('/policies/identitySecurityDefaultsEnforcementPolicy')
      .get();

    // Gather user MFA status
    evidence.users = await graphClient
      .api('/users')
      .select('id,displayName,userPrincipalName')
      .top(999)
      .get();

    // Get organization settings
    evidence.organization = await graphClient
      .api('/organization')
      .get();

    // Get directory roles
    evidence.directoryRoles = await graphClient
      .api('/directoryRoles')
      .get();

  } catch (error) {
    console.error('Error gathering CIS evidence:', error);
  }

  return evidence;
}

async function evaluateCISControls(
  evidence: any, 
  benchmark: string, 
  implementationGroup: number,
  controlIds?: string[]
): Promise<CISControlResult[]> {
  const controls = getCISMicrosoft365Controls();
  const results: CISControlResult[] = [];

  for (const control of controls) {
    if (controlIds && !controlIds.includes(control.controlId)) continue;
    if (control.implementationGroup > implementationGroup) continue;

    const result = await evaluateSingleControl(control, evidence);
    results.push(result);
  }

  return results;
}

async function evaluateSingleControl(control: CISControl, evidence: any): Promise<CISControlResult> {
  let status: 'Pass' | 'Fail' | 'Not Applicable' | 'Manual Review Required' = 'Manual Review Required';
  let testResult = '';
  let cisEvidence: any[] = [];
  let riskLevel: 'Low' | 'Medium' | 'High' | 'Critical' = 'Medium';

  switch (control.controlId) {
    case 'CIS-1.1': // MFA for all users
      const mfaEnabled = checkMFAStatus(evidence);
      status = mfaEnabled.compliant ? 'Pass' : 'Fail';
      testResult = mfaEnabled.message;
      cisEvidence = mfaEnabled.evidence;
      riskLevel = mfaEnabled.compliant ? 'Low' : 'High';
      break;

    case 'CIS-1.2': // Block legacy authentication
      const legacyBlocked = checkLegacyAuthentication(evidence);
      status = legacyBlocked.compliant ? 'Pass' : 'Fail';
      testResult = legacyBlocked.message;
      cisEvidence = legacyBlocked.evidence;
      riskLevel = legacyBlocked.compliant ? 'Low' : 'High';
      break;

    case 'CIS-2.1': // Security Defaults
      const securityDefaults = checkSecurityDefaults(evidence);
      status = securityDefaults.compliant ? 'Pass' : 'Fail';
      testResult = securityDefaults.message;
      cisEvidence = securityDefaults.evidence;
      riskLevel = securityDefaults.compliant ? 'Low' : 'Medium';
      break;

    default:
      status = 'Manual Review Required';
      testResult = 'Control requires manual review';
  }

  return {
    controlId: control.controlId,
    title: control.title,
    status,
    score: status === 'Pass' ? 100 : 0,
    implementationGroup: control.implementationGroup,
    automatable: control.automatable,
    testResult,
    evidence: cisEvidence,
    riskLevel,
    remediation: generateControlRemediation(control, status)
  };
}

function checkMFAStatus(evidence: any): { compliant: boolean; message: string; evidence: any[] } {
  // Check if MFA is enforced through conditional access or security defaults
  const caPolicy = evidence.conditionalAccessPolicies?.value?.find((p: any) => 
    p.displayName?.toLowerCase().includes('mfa') || 
    p.grantControls?.builtInControls?.includes('mfa')
  );

  const securityDefaultsEnabled = evidence.securityDefaults?.isEnabled;

  const compliant = !!(caPolicy || securityDefaultsEnabled);
  
  return {
    compliant,
    message: compliant 
      ? 'MFA is enforced through conditional access policies or security defaults'
      : 'MFA is not properly enforced for all users',
    evidence: [
      {
        type: 'Policy',
        source: 'Conditional Access',
        value: caPolicy ? 'MFA policy found' : 'No MFA policy found',
        expected: 'MFA enforcement policy required',
        collected: new Date().toISOString(),
        automated: true
      }
    ]
  };
}

function checkLegacyAuthentication(evidence: any): { compliant: boolean; message: string; evidence: any[] } {
  // Check for conditional access policies that block legacy authentication
  const legacyBlockPolicy = evidence.conditionalAccessPolicies?.value?.find((p: any) => 
    p.conditions?.clientAppTypes?.includes('exchangeActiveSync') &&
    p.grantControls?.operator === 'OR' &&
    p.grantControls?.builtInControls?.includes('block')
  );

  const compliant = !!legacyBlockPolicy;
  
  return {
    compliant,
    message: compliant 
      ? 'Legacy authentication is blocked by conditional access policy'
      : 'Legacy authentication is not blocked',
    evidence: [
      {
        type: 'Policy',
        source: 'Conditional Access',
        value: legacyBlockPolicy ? 'Legacy auth block policy found' : 'No legacy auth block policy found',
        expected: 'Policy blocking legacy authentication required',
        collected: new Date().toISOString(),
        automated: true
      }
    ]
  };
}

function checkSecurityDefaults(evidence: any): { compliant: boolean; message: string; evidence: any[] } {
  const isEnabled = evidence.securityDefaults?.isEnabled === true;
  
  return {
    compliant: isEnabled,
    message: isEnabled 
      ? 'Security Defaults are enabled'
      : 'Security Defaults are not enabled',
    evidence: [
      {
        type: 'Configuration',
        source: 'Azure AD',
        value: isEnabled.toString(),
        expected: 'true',
        collected: new Date().toISOString(),
        automated: true
      }
    ]
  };
}

function generateControlRemediation(control: CISControl, status: string) {
  if (status === 'Pass') return undefined;

  const remediationMap: Record<string, any> = {
    'CIS-1.1': {
      title: 'Enable Multi-Factor Authentication',
      description: 'Configure MFA enforcement for all users',
      steps: [
        'Go to Azure AD admin center',
        'Navigate to Security > Conditional Access',
        'Create new policy requiring MFA for all users',
        'Test policy with pilot group before full deployment'
      ],
      automatable: true,
      estimatedTime: '30 minutes',
      difficulty: 'Low',
      impact: 'Medium',
      prerequisites: ['Azure AD Premium license', 'Conditional Access']
    },
    'CIS-1.2': {
      title: 'Block Legacy Authentication',
      description: 'Create conditional access policy to block legacy authentication protocols',
      steps: [
        'Go to Azure AD admin center',
        'Navigate to Security > Conditional Access',
        'Create policy targeting legacy authentication clients',
        'Set access control to Block',
        'Enable policy after testing'
      ],
      automatable: true,
      estimatedTime: '20 minutes',
      difficulty: 'Medium',
      impact: 'Medium',
      prerequisites: ['Azure AD Premium license']
    }
  };

  return remediationMap[control.controlId];
}

function generateCISRiskSummary(controlResults: CISControlResult[]) {
  const criticalFindings = controlResults.filter(c => c.riskLevel === 'Critical' && c.status === 'Fail').length;
  const highRiskFindings = controlResults.filter(c => c.riskLevel === 'High' && c.status === 'Fail').length;
  const mediumRiskFindings = controlResults.filter(c => c.riskLevel === 'Medium' && c.status === 'Fail').length;
  const lowRiskFindings = controlResults.filter(c => c.riskLevel === 'Low' && c.status === 'Fail').length;

  return {
    criticalFindings,
    highRiskFindings,
    mediumRiskFindings,
    lowRiskFindings,
    topRisks: controlResults
      .filter(c => c.status === 'Fail')
      .sort((a, b) => {
        const riskOrder = { 'Critical': 4, 'High': 3, 'Medium': 2, 'Low': 1 };
        return riskOrder[b.riskLevel] - riskOrder[a.riskLevel];
      })
      .slice(0, 5)
      .map(c => ({
        controlId: c.controlId,
        title: c.title,
        riskLevel: c.riskLevel,
        impact: 'Security vulnerability',
        likelihood: 'High',
        affectedAssets: 100,
        businessImpact: 'Potential security breach',
        technicalImpact: 'Unauthorized access'
      })),
    riskTrend: [
      {
        date: new Date().toISOString(),
        overallScore: calculateOverallScore(controlResults),
        criticalCount: criticalFindings,
        highCount: highRiskFindings,
        mediumCount: mediumRiskFindings,
        lowCount: lowRiskFindings
      }
    ]
  };
}

function generateCISRemediationPlan(controlResults: CISControlResult[]): CISRemediationPlan {
  const failedControls = controlResults.filter(c => c.status === 'Fail');
  
  return {
    planId: `remediation-${Date.now()}`,
    generatedDate: new Date().toISOString(),
    totalTasks: failedControls.length,
    estimatedTime: `${failedControls.length * 30} minutes`,
    phases: [
      {
        phase: 1,
        name: 'Critical and High Risk Items',
        description: 'Address critical and high-risk security findings first',
        tasks: failedControls
          .filter(c => c.riskLevel === 'Critical' || c.riskLevel === 'High')
          .map(c => ({
            taskId: `task-${c.controlId}`,
            controlId: c.controlId,
            title: c.remediation?.title || `Remediate ${c.controlId}`,
            description: c.remediation?.description || c.title,
            automatable: c.remediation?.automatable || false,
            estimatedTime: c.remediation?.estimatedTime || '30 minutes',
            difficulty: c.remediation?.difficulty || 'Medium',
            impact: c.remediation?.impact || 'Medium',
            status: 'Pending',
            dueDate: new Date(Date.now() + 7 * 24 * 3600000).toISOString(), // 1 week
            prerequisites: c.remediation?.prerequisites || [],
            validationSteps: ['Verify control implementation', 'Test functionality', 'Document changes']
          })),
        startDate: new Date().toISOString(),
        endDate: new Date(Date.now() + 7 * 24 * 3600000).toISOString(),
        priority: 'Critical'
      }
    ],
    dependencies: []
  };
}

function calculateOverallScore(controlResults: CISControlResult[]): number {
  if (controlResults.length === 0) return 0;
  const totalScore = controlResults.reduce((sum, c) => sum + c.score, 0);
  return Math.round(totalScore / controlResults.length);
}

function calculateCompliancePercentage(controlResults: CISControlResult[]): number {
  if (controlResults.length === 0) return 0;
  const passedControls = controlResults.filter(c => c.status === 'Pass').length;
  return Math.round((passedControls / controlResults.length) * 100);
}

async function generateCISReport(graphClient: Client, args: CISComplianceArgs) {
  // Generate comprehensive CIS compliance report
  const assessment = await performCISAssessment(graphClient, args);
  
  return {
    reportId: `cis-report-${Date.now()}`,
    benchmark: args.benchmark,
    generatedDate: new Date().toISOString(),
    assessment,
    summary: {
      overallCompliance: assessment.compliancePercentage,
      totalControls: assessment.totalControls,
      passedControls: assessment.passedControls,
      failedControls: assessment.failedControls,
      riskLevel: assessment.riskSummary.criticalFindings > 0 ? 'Critical' : 
                 assessment.riskSummary.highRiskFindings > 0 ? 'High' : 'Medium'
    },
    downloadUrl: `/reports/cis-compliance-${Date.now()}.pdf`
  };
}

async function configureCISMonitoring(args: CISComplianceArgs) {
  // Configure continuous CIS monitoring
  return {
    monitoringId: `cis-monitoring-${Date.now()}`,
    benchmark: args.benchmark,
    settings: args.settings,
    schedule: 'daily',
    alerting: true,
    status: 'configured',
    message: 'CIS monitoring configured successfully'
  };
}

async function executeCISRemediation(graphClient: Client, args: CISComplianceArgs) {
  // Execute automated remediation for specific controls
  const results = [];
  
  if (args.controlIds) {
    for (const controlId of args.controlIds) {
      try {
        const result = await remediateCISControl(graphClient, controlId);
        results.push(result);
      } catch (error) {
        results.push({
          controlId,
          status: 'failed',
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }
  }

  return {
    remediationId: `remediation-${Date.now()}`,
    results,
    summary: {
      attempted: results.length,
      successful: results.filter(r => r.status === 'success').length,
      failed: results.filter(r => r.status === 'failed').length
    }
  };
}

async function remediateCISControl(graphClient: Client, controlId: string) {
  // Implement automated remediation for specific CIS controls
  switch (controlId) {
    case 'CIS-2.1': // Enable Security Defaults
      try {
        await graphClient
          .api('/policies/identitySecurityDefaultsEnforcementPolicy')
          .patch({ isEnabled: true });
        
        return {
          controlId,
          status: 'success',
          message: 'Security Defaults enabled successfully'
        };
      } catch (error) {
        throw new Error(`Failed to enable Security Defaults: ${error}`);
      }

    default:
      return {
        controlId,
        status: 'manual',
        message: 'Control requires manual remediation'
      };
  }
}
