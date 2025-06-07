import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { AuditReportArgs } from '../types/compliance-types.js';
import * as fs from 'fs';
import * as path from 'path';
import { createObjectCsvWriter } from 'csv-writer';
import * as XLSX from 'xlsx';
import Handlebars from 'handlebars';

// Audit Report Generation Handler
export async function handleAuditReports(
  graphClient: Client,
  args: AuditReportArgs
): Promise<{ content: { type: string; text: string }[] }> {
  
  // Generate the report data
  const reportData = await generateReportData(graphClient, args);
  
  // Generate the report in the requested format
  let result: any;
  switch (args.format) {
    case 'csv':
      result = await generateCSVReport(reportData, args);
      break;
    case 'html':
      result = await generateHTMLReport(reportData, args);
      break;
    case 'pdf':
      result = await generatePDFReport(reportData, args);
      break;
    case 'xlsx':
      result = await generateExcelReport(reportData, args);
      break;
    default:
      throw new McpError(ErrorCode.InvalidParams, `Unsupported format: ${args.format}`);
  }

  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}

// Generate report data based on framework and type
async function generateReportData(graphClient: Client, args: AuditReportArgs) {
  const reportId = `${args.framework}-${args.reportType}-${Date.now()}`;
  
  // Get compliance data from Microsoft Graph
  const secureScore = await getSecureScoreData(graphClient);
  const controls = await getControlsData(graphClient, args.framework);
  const complianceData = await getComplianceData(graphClient, args);
  
  // Generate report based on type
  switch (args.reportType) {
    case 'full':
      return generateFullReport(reportId, args, secureScore, controls, complianceData);
    case 'summary':
      return generateSummaryReport(reportId, args, secureScore, controls);
    case 'gaps':
      return generateGapsReport(reportId, args, controls);
    case 'evidence':
      return generateEvidenceReport(reportId, args, complianceData);
    case 'executive':
      return generateExecutiveReport(reportId, args, secureScore, controls);
    case 'control_matrix':
      return generateControlMatrixReport(reportId, args, controls);
    case 'risk_assessment':
      return generateRiskAssessmentReport(reportId, args, controls, complianceData);
    default:
      throw new McpError(ErrorCode.InvalidParams, `Unsupported report type: ${args.reportType}`);
  }
}

// Data collection functions
async function getSecureScoreData(graphClient: Client) {
  try {
    const secureScore = await graphClient.api('/security/secureScores').top(1).get();
    return secureScore.value[0] || {};
  } catch (error) {
    console.warn('Could not fetch secure score data:', error);
    return {};
  }
}

async function getControlsData(graphClient: Client, framework: string) {
  try {
    const controls = await graphClient.api('/security/secureScoreControlProfiles').get();
    return controls.value || [];
  } catch (error) {
    console.warn('Could not fetch controls data:', error);
    return [];
  }
}

async function getComplianceData(graphClient: Client, args: AuditReportArgs) {
  try {
    const devices = await graphClient.api('/deviceManagement/managedDevices').get();
    const policies = await graphClient.api('/deviceManagement/deviceCompliancePolicies').get();
    return { devices: devices.value || [], policies: policies.value || [] };
  } catch (error) {
    console.warn('Could not fetch compliance data:', error);
    return { devices: [], policies: [] };
  }
}

// Report generation functions
function generateFullReport(reportId: string, args: AuditReportArgs, secureScore: any, controls: any[], complianceData: any) {
  return {
    id: reportId,
    framework: args.framework,
    reportType: args.reportType,
    generatedDate: new Date().toISOString(),
    period: args.dateRange,
    summary: {
      totalControls: controls.length,
      implementedControls: controls.filter(c => c.implementationStatus === 'implemented').length,
      partiallyImplementedControls: controls.filter(c => c.implementationStatus === 'partiallyImplemented').length,
      notImplementedControls: controls.filter(c => c.implementationStatus === 'notImplemented').length,
      notApplicableControls: controls.filter(c => c.implementationStatus === 'notApplicable').length,
      compliancePercentage: Math.round((secureScore.currentScore / secureScore.maxScore) * 100) || 0,
      riskScore: secureScore.averageComparativeScores?.[0]?.averageScore || 0,
      lastAssessmentDate: new Date().toISOString()
    },
    controls: controls.map(control => ({
      controlId: control.id,
      controlName: control.title,
      category: control.category,
      implementationStatus: control.implementationStatus || 'notAssessed',
      testingStatus: control.userImpact || 'notTested',
      riskLevel: control.maxScore > 7 ? 'high' : control.maxScore > 4 ? 'medium' : 'low',
      lastTested: control.lastModifiedDateTime || new Date().toISOString(),
      nextAssessment: new Date(Date.now() + 90 * 24 * 3600000).toISOString(), // 90 days
      owner: 'System Administrator',
      evidenceCount: 0,
      score: control.currentScore || 0
    })),
    gaps: controls
      .filter(c => c.implementationStatus !== 'implemented')
      .map(control => ({
        controlId: control.id,
        controlName: control.title,
        category: control.category,
        currentStatus: control.implementationStatus || 'notImplemented',
        requiredStatus: 'implemented',
        riskLevel: control.maxScore > 7 ? 'high' : control.maxScore > 4 ? 'medium' : 'low',
        impact: control.description || 'Security impact not assessed',
        recommendedActions: [control.remediationImpact || 'Implement control as specified'],
        estimatedEffort: control.implementationCost || 'Medium',
        priority: control.maxScore || 5
      })),
    recommendations: [
      {
        id: 'rec-001',
        type: 'immediate',
        priority: 'high',
        title: 'Address High-Risk Control Gaps',
        description: 'Focus on implementing high-risk controls first',
        impact: 'Significant risk reduction',
        effort: 'Medium',
        resources: ['Security Team', 'IT Team'],
        timeline: '30 days',
        relatedControls: controls.filter(c => c.maxScore > 7).map(c => c.id)
      }
    ],
    evidence: [],
    metadata: {
      generatedBy: 'M365 Core MCP Server',
      generationTime: Date.now(),
      dataSource: 'Microsoft Graph API',
      version: '1.0'
    }
  };
}

function generateSummaryReport(reportId: string, args: AuditReportArgs, secureScore: any, controls: any[]) {
  return {
    id: reportId,
    framework: args.framework,
    reportType: args.reportType,
    generatedDate: new Date().toISOString(),
    period: args.dateRange,
    summary: {
      totalControls: controls.length,
      implementedControls: controls.filter(c => c.implementationStatus === 'implemented').length,
      compliancePercentage: Math.round((secureScore.currentScore / secureScore.maxScore) * 100) || 0,
      riskScore: secureScore.averageComparativeScores?.[0]?.averageScore || 0,
      lastAssessmentDate: new Date().toISOString()
    }
  };
}

function generateGapsReport(reportId: string, args: AuditReportArgs, controls: any[]) {
  return {
    id: reportId,
    framework: args.framework,
    reportType: args.reportType,
    gaps: controls
      .filter(c => c.implementationStatus !== 'implemented')
      .map(control => ({
        controlId: control.id,
        controlName: control.title,
        category: control.category,
        currentStatus: control.implementationStatus || 'notImplemented',
        requiredStatus: 'implemented',
        riskLevel: control.maxScore > 7 ? 'high' : control.maxScore > 4 ? 'medium' : 'low',
        priority: control.maxScore || 5
      }))
  };
}

function generateEvidenceReport(reportId: string, args: AuditReportArgs, complianceData: any) {
  return {
    id: reportId,
    framework: args.framework,
    reportType: args.reportType,
    evidence: [] // Evidence would be collected from various sources
  };
}

function generateExecutiveReport(reportId: string, args: AuditReportArgs, secureScore: any, controls: any[]) {
  return {
    id: reportId,
    framework: args.framework,
    reportType: args.reportType,
    executiveSummary: {
      overallComplianceScore: Math.round((secureScore.currentScore / secureScore.maxScore) * 100) || 0,
      keyFindings: [
        'Organization maintains good security posture',
        'Some controls require immediate attention',
        'Regular assessment schedule is recommended'
      ],
      recommendations: [
        'Implement missing high-priority controls',
        'Establish regular compliance monitoring',
        'Enhance security awareness training'
      ],
      riskLevel: 'Medium'
    }
  };
}

function generateControlMatrixReport(reportId: string, args: AuditReportArgs, controls: any[]) {
  return {
    id: reportId,
    framework: args.framework,
    reportType: args.reportType,
    controlMatrix: controls.map(control => ({
      controlId: control.id,
      controlName: control.title,
      category: control.category,
      implementationStatus: control.implementationStatus || 'notAssessed',
      testingStatus: control.userImpact || 'notTested',
      owner: 'System Administrator',
      lastTested: control.lastModifiedDateTime || new Date().toISOString()
    }))
  };
}

function generateRiskAssessmentReport(reportId: string, args: AuditReportArgs, controls: any[], complianceData: any) {
  return {
    id: reportId,
    framework: args.framework,
    reportType: args.reportType,
    riskAssessment: {
      overallRiskLevel: 'Medium',
      criticalRisks: controls.filter(c => c.maxScore > 8).length,
      highRisks: controls.filter(c => c.maxScore > 6 && c.maxScore <= 8).length,
      mediumRisks: controls.filter(c => c.maxScore > 3 && c.maxScore <= 6).length,
      lowRisks: controls.filter(c => c.maxScore <= 3).length,
      riskTrends: [] // Would include historical risk data
    }
  };
}

// Format-specific generation functions
async function generateCSVReport(reportData: any, args: AuditReportArgs): Promise<any> {
  const outputDir = './outputs';
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const fileName = `${reportData.id}.csv`;
  const filePath = path.join(outputDir, fileName);

  // Convert report data to CSV format
  const csvData = reportData.controls || reportData.gaps || [reportData.summary];
  
  const csvWriter = createObjectCsvWriter({
    path: filePath,
    header: Object.keys(csvData[0] || {}).map(key => ({ id: key, title: key }))
  });

  await csvWriter.writeRecords(csvData);

  return {
    reportId: reportData.id,
    format: 'csv',
    filePath: filePath,
    fileName: fileName,
    generatedDate: new Date().toISOString(),
    recordCount: csvData.length
  };
}

async function generateHTMLReport(reportData: any, args: AuditReportArgs): Promise<any> {
  const outputDir = './outputs';
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const fileName = `${reportData.id}.html`;
  const filePath = path.join(outputDir, fileName);

  // Create HTML template
  const template = Handlebars.compile(`
    <!DOCTYPE html>
    <html>
    <head>
        <title>{{framework}} {{reportType}} Report</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 20px; }
            .header { background-color: #0078d4; color: white; padding: 20px; }
            .summary { background-color: #f5f5f5; padding: 15px; margin: 20px 0; }
            table { width: 100%; border-collapse: collapse; margin: 20px 0; }
            th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            .risk-high { color: #d13438; }
            .risk-medium { color: #ff8c00; }
            .risk-low { color: #107c10; }
        </style>
    </head>
    <body>
        <div class="header">
            <h1>{{framework}} {{reportType}} Report</h1>
            <p>Generated on: {{generatedDate}}</p>
        </div>
        
        {{#if summary}}
        <div class="summary">
            <h2>Summary</h2>
            <p><strong>Total Controls:</strong> {{summary.totalControls}}</p>
            <p><strong>Implemented Controls:</strong> {{summary.implementedControls}}</p>
            <p><strong>Compliance Percentage:</strong> {{summary.compliancePercentage}}%</p>
        </div>
        {{/if}}
        
        {{#if controls}}
        <h2>Controls</h2>
        <table>
            <tr>
                <th>Control ID</th>
                <th>Control Name</th>
                <th>Category</th>
                <th>Implementation Status</th>
                <th>Risk Level</th>
            </tr>
            {{#each controls}}
            <tr>
                <td>{{controlId}}</td>
                <td>{{controlName}}</td>
                <td>{{category}}</td>
                <td>{{implementationStatus}}</td>
                <td class="risk-{{riskLevel}}">{{riskLevel}}</td>
            </tr>
            {{/each}}
        </table>
        {{/if}}
    </body>
    </html>
  `);

  const html = template(reportData);
  fs.writeFileSync(filePath, html);

  return {
    reportId: reportData.id,
    format: 'html',
    filePath: filePath,
    fileName: fileName,
    generatedDate: new Date().toISOString()
  };
}

async function generatePDFReport(reportData: any, args: AuditReportArgs): Promise<any> {
  // First generate HTML
  const htmlReport = await generateHTMLReport(reportData, args);
  
  const outputDir = './outputs';
  const fileName = `${reportData.id}.pdf`;
  const filePath = path.join(outputDir, fileName);

  // Convert HTML to PDF
  const htmlContent = fs.readFileSync(htmlReport.filePath, 'utf8');
  const options = { format: 'A4', printBackground: true };
  const file = { content: htmlContent };
  
  try {
    // Dynamically import html-pdf-node
    const htmlPdf: any = await import('html-pdf-node');
    const pdfBuffer = await htmlPdf.generatePdf(file, options);
    fs.writeFileSync(filePath, pdfBuffer);
    
    // Clean up temporary HTML file
    fs.unlinkSync(htmlReport.filePath);
    
    return {
      reportId: reportData.id,
      format: 'pdf',
      filePath: filePath,
      fileName: fileName,
      generatedDate: new Date().toISOString()
    };
  } catch (error) {
    console.error('PDF generation failed:', error);
    // Return HTML report as fallback
    return htmlReport;
  }
}

async function generateExcelReport(reportData: any, args: AuditReportArgs): Promise<any> {
  const outputDir = './outputs';
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const fileName = `${reportData.id}.xlsx`;
  const filePath = path.join(outputDir, fileName);

  // Create workbook
  const workbook = XLSX.utils.book_new();

  // Add summary sheet
  if (reportData.summary) {
    const summaryData = [
      ['Framework', reportData.framework],
      ['Report Type', reportData.reportType],
      ['Generated Date', reportData.generatedDate],
      ['Total Controls', reportData.summary.totalControls],
      ['Implemented Controls', reportData.summary.implementedControls],
      ['Compliance Percentage', `${reportData.summary.compliancePercentage}%`]
    ];
    const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary');
  }

  // Add controls sheet
  if (reportData.controls) {
    const controlsSheet = XLSX.utils.json_to_sheet(reportData.controls);
    XLSX.utils.book_append_sheet(workbook, controlsSheet, 'Controls');
  }

  // Add gaps sheet
  if (reportData.gaps) {
    const gapsSheet = XLSX.utils.json_to_sheet(reportData.gaps);
    XLSX.utils.book_append_sheet(workbook, gapsSheet, 'Gaps');
  }

  // Write file
  XLSX.writeFile(workbook, filePath);

  return {
    reportId: reportData.id,
    format: 'xlsx',
    filePath: filePath,
    fileName: fileName,
    generatedDate: new Date().toISOString(),
    sheets: workbook.SheetNames
  };
}
