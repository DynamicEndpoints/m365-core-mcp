/**
 * HTML Report Handler
 * Creates professional HTML reports with charts and styling
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import {
  HTMLReportArgs,
  HTMLSection,
  HTMLTemplate,
  ChartData,
  TableData
} from '../types/document-generation-types.js';

/**
 * Handle HTML report operations
 */
export async function handleHTMLReports(
  args: HTMLReportArgs,
  graphClient: Client
): Promise<string> {
  try {
    switch (args.action) {
      case 'create':
        return await createHTMLReport(args, graphClient);
      case 'get':
        return await getHTMLReport(args, graphClient);
      case 'list':
        return await listHTMLReports(args, graphClient);
      default:
        throw new McpError(
          ErrorCode.InvalidRequest,
          `Unknown action: ${args.action}`
        );
    }
  } catch (error) {
    if (error instanceof McpError) throw error;
    throw new McpError(
      ErrorCode.InternalError,
      `HTML report operation failed: ${error instanceof Error ? error.message : 'Unknown error'}`
    );
  }
}

/**
 * Create a new HTML report
 */
async function createHTMLReport(
  args: HTMLReportArgs,
  graphClient: Client
): Promise<string> {
  if (!args.fileName) {
    throw new McpError(ErrorCode.InvalidRequest, 'fileName is required for create action');
  }

  if (!args.template) {
    throw new McpError(ErrorCode.InvalidRequest, 'template is required for create action');
  }

  if (!args.sections || args.sections.length === 0) {
    throw new McpError(ErrorCode.InvalidRequest, 'At least one section is required');
  }

  // Determine drive location (default to user's OneDrive)
  const driveId = args.driveId || 'me';
  const folderPath = args.folderId ? `/items/${args.folderId}` : '/root';

  // Create HTML file
  const fileName = args.fileName.endsWith('.html') ? args.fileName : `${args.fileName}.html`;
  
  // Generate HTML content
  const htmlContent = generateHTMLContent(args.sections, args.template, args.includeCharts);
  
  const uploadedFile = await graphClient
    .api(`/drives/${driveId}${folderPath}:/${fileName}:/content`)
    .header('Content-Type', 'text/html')
    .put(Buffer.from(htmlContent, 'utf-8'));

  return JSON.stringify({
    success: true,
    fileId: uploadedFile.id,
    fileName: uploadedFile.name,
    webUrl: uploadedFile.webUrl,
    driveId: uploadedFile.parentReference?.driveId,
    message: `HTML report "${fileName}" created successfully with ${args.sections.length} sections`
  }, null, 2);
}

/**
 * Get an existing HTML report
 */
async function getHTMLReport(
  args: HTMLReportArgs,
  graphClient: Client
): Promise<string> {
  if (!args.fileId) {
    throw new McpError(ErrorCode.InvalidRequest, 'fileId is required for get action');
  }

  const driveId = args.driveId || 'me';
  
  const file = await graphClient
    .api(`/drives/${driveId}/items/${args.fileId}`)
    .get();

  // Optionally get content
  const content = await graphClient
    .api(`/drives/${driveId}/items/${args.fileId}/content`)
    .get();

  return JSON.stringify({
    success: true,
    file: {
      id: file.id,
      name: file.name,
      size: file.size,
      webUrl: file.webUrl,
      createdDateTime: file.createdDateTime,
      lastModifiedDateTime: file.lastModifiedDateTime,
      createdBy: file.createdBy?.user?.displayName,
      lastModifiedBy: file.lastModifiedBy?.user?.displayName
    },
    contentPreview: content ? content.substring(0, 500) : null
  }, null, 2);
}

/**
 * List HTML reports
 */
async function listHTMLReports(
  args: HTMLReportArgs,
  graphClient: Client
): Promise<string> {
  const driveId = args.driveId || 'me';
  const folderPath = args.folderId ? `/items/${args.folderId}` : '/root';
  
  let query = graphClient
    .api(`/drives/${driveId}${folderPath}/children`)
    .filter("endsWith(name,'.html')");

  if (args.filter) {
    query = query.filter(args.filter);
  }

  if (args.top) {
    query = query.top(args.top);
  }

  const result = await query.get();

  return JSON.stringify({
    success: true,
    count: result.value?.length || 0,
    reports: result.value?.map((file: any) => ({
      id: file.id,
      name: file.name,
      size: file.size,
      webUrl: file.webUrl,
      createdDateTime: file.createdDateTime,
      lastModifiedDateTime: file.lastModifiedDateTime
    })) || []
  }, null, 2);
}

/**
 * Generate complete HTML content
 */
function generateHTMLContent(
  sections: HTMLSection[],
  template: HTMLTemplate,
  includeCharts?: boolean
): string {
  const theme = getThemeStyles(template.theme || 'modern');
  
  let html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHtml(template.title)}</title>
  ${template.author ? `<meta name="author" content="${escapeHtml(template.author)}">` : ''}
  ${template.description ? `<meta name="description" content="${escapeHtml(template.description)}">` : ''}
  
  ${template.includeBootstrap ? '<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">' : ''}
  ${includeCharts || template.includeChartJS ? '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.js"></script>' : ''}
  
  <style>
    ${theme}
    ${template.customCSS || ''}
  </style>
</head>
<body>
  <div class="container">
    <header class="report-header">
      ${template.companyLogo ? `<img src="${escapeHtml(template.companyLogo)}" alt="Company Logo" class="company-logo">` : ''}
      <h1>${escapeHtml(template.title)}</h1>
      ${template.description ? `<p class="report-description">${escapeHtml(template.description)}</p>` : ''}
      ${template.companyName ? `<p class="company-name">${escapeHtml(template.companyName)}</p>` : ''}
      <p class="report-date">Generated: ${new Date().toLocaleString()}</p>
    </header>
    
    <main class="report-content">`;

  // Generate sections
  sections.forEach((section, index) => {
    html += generateHTMLSection(section, index);
  });

  html += `
    </main>
    
    <footer class="report-footer">
      ${template.companyName ? `<p>&copy; ${new Date().getFullYear()} ${escapeHtml(template.companyName)}</p>` : ''}
      <p>Report generated by M365 Core MCP Server</p>
    </footer>
  </div>
  
  ${includeCharts ? getChartScripts(sections) : ''}
</body>
</html>`;

  return html;
}

/**
 * Generate HTML for a section
 */
function generateHTMLSection(section: HTMLSection, index: number): string {
  let html = '<section class="report-section">';

  if (section.heading) {
    const level = section.level || 2;
    html += `<h${level}>${escapeHtml(section.heading)}</h${level}>`;
  }

  switch (section.type) {
    case 'paragraph':
      html += `<p>${section.content || ''}</p>`;
      break;

    case 'table':
      if (section.tableData) {
        html += generateHTMLTable(section.tableData);
      }
      break;

    case 'chart':
      if (section.chartData) {
        html += generateChartHTML(section.chartData, `chart-${index}`);
      }
      break;

    case 'list':
      if (section.items) {
        html += '<ul>';
        section.items.forEach(item => {
          html += `<li>${escapeHtml(item)}</li>`;
        });
        html += '</ul>';
      }
      break;

    case 'grid':
      html += '<div class="grid-container" style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px;">';
      if (section.content) {
        html += section.content;
      }
      html += '</div>';
      break;

    case 'card':
      if (section.cardData) {
        html += '<div class="card-container">';
        section.cardData.forEach(card => {
          html += `
            <div class="card" style="background: ${card.color || '#f8f9fa'};">
              ${card.icon ? `<div class="card-icon">${card.icon}</div>` : ''}
              <h3 class="card-title">${escapeHtml(card.title)}</h3>
              <p class="card-value">${card.value}</p>
            </div>`;
        });
        html += '</div>';
      }
      break;

    case 'alert':
      const alertClass = section.alertType || 'info';
      html += `<div class="alert alert-${alertClass}">${section.content || ''}</div>`;
      break;

    default:
      if (section.content) {
        html += section.content;
      }
  }

  html += '</section>';
  return html;
}

/**
 * Generate HTML table
 */
function generateHTMLTable(tableData: TableData): string {
  const tableClass = tableData.style === 'striped' ? 'table-striped' : tableData.style === 'bordered' ? 'table-bordered' : '';
  
  let html = `<table class="report-table ${tableClass}">
    <thead>
      <tr>`;
  
  tableData.headers.forEach(header => {
    html += `<th>${escapeHtml(header)}</th>`;
  });
  
  html += `</tr>
    </thead>
    <tbody>`;
  
  tableData.rows.forEach(row => {
    html += '<tr>';
    row.forEach(cell => {
      html += `<td>${escapeHtml(cell)}</td>`;
    });
    html += '</tr>';
  });
  
  html += `</tbody>
  </table>`;
  
  return html;
}

/**
 * Generate chart HTML canvas
 */
function generateChartHTML(chartData: ChartData, chartId: string): string {
  return `<div class="chart-container">
    <canvas id="${chartId}" width="400" height="200"></canvas>
  </div>`;
}

/**
 * Generate Chart.js scripts for all charts
 */
function getChartScripts(sections: HTMLSection[]): string {
  let scripts = '<script>';
  
  sections.forEach((section, index) => {
    if (section.type === 'chart' && section.chartData) {
      const chartId = `chart-${index}`;
      const chartData = section.chartData;
      
      scripts += `
        const ctx${index} = document.getElementById('${chartId}').getContext('2d');
        new Chart(ctx${index}, {
          type: '${chartData.type}',
          data: {
            labels: ${JSON.stringify(chartData.categories)},
            datasets: ${JSON.stringify(chartData.series.map(s => ({
              label: s.name,
              data: s.values,
              backgroundColor: 'rgba(54, 162, 235, 0.2)',
              borderColor: 'rgba(54, 162, 235, 1)',
              borderWidth: 1
            })))}
          },
          options: {
            responsive: true,
            plugins: {
              title: {
                display: true,
                text: '${chartData.title}'
              }
            }
          }
        });`;
    }
  });
  
  scripts += '</script>';
  return scripts;
}

/**
 * Get theme CSS styles
 */
function getThemeStyles(theme: string): string {
  const themes: Record<string, string> = {
    modern: `
      body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif; background: #f5f5f5; margin: 0; padding: 20px; }
      .container { max-width: 1200px; margin: 0 auto; background: white; padding: 40px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); border-radius: 8px; }
      .report-header { border-bottom: 3px solid #0078d4; padding-bottom: 20px; margin-bottom: 30px; }
      .report-header h1 { color: #0078d4; margin: 0; font-size: 2.5em; }
      .report-description { color: #666; font-size: 1.1em; margin-top: 10px; }
      .report-date { color: #999; font-size: 0.9em; }
      .company-logo { max-height: 60px; margin-bottom: 20px; }
      .report-section { margin-bottom: 40px; }
      .report-section h2 { color: #333; border-left: 4px solid #0078d4; padding-left: 15px; }
      .report-table { width: 100%; border-collapse: collapse; margin: 20px 0; }
      .report-table th { background: #0078d4; color: white; padding: 12px; text-align: left; }
      .report-table td { padding: 10px; border-bottom: 1px solid #ddd; }
      .report-table tr:hover { background: #f5f5f5; }
      .chart-container { margin: 20px 0; max-width: 800px; }
      .card-container { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin: 20px 0; }
      .card { padding: 20px; border-radius: 8px; text-align: center; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
      .card-title { font-size: 1em; color: #666; margin: 10px 0; }
      .card-value { font-size: 2em; font-weight: bold; color: #0078d4; margin: 0; }
      .alert { padding: 15px; border-radius: 5px; margin: 20px 0; }
      .alert-info { background: #d1ecf1; border-left: 4px solid #0c5460; color: #0c5460; }
      .alert-success { background: #d4edda; border-left: 4px solid #155724; color: #155724; }
      .alert-warning { background: #fff3cd; border-left: 4px solid #856404; color: #856404; }
      .alert-danger { background: #f8d7da; border-left: 4px solid #721c24; color: #721c24; }
      .report-footer { margin-top: 60px; padding-top: 20px; border-top: 1px solid #ddd; text-align: center; color: #999; font-size: 0.9em; }
    `,
    classic: `
      body { font-family: Georgia, 'Times New Roman', serif; background: #fff; margin: 0; padding: 20px; line-height: 1.6; }
      .container { max-width: 900px; margin: 0 auto; }
      .report-header { text-align: center; border-bottom: 2px solid #333; padding-bottom: 20px; margin-bottom: 40px; }
      .report-header h1 { margin: 0; font-size: 2.5em; }
      .report-table { width: 100%; border: 1px solid #333; margin: 20px 0; }
      .report-table th, .report-table td { border: 1px solid #333; padding: 8px; }
    `,
    minimal: `
      body { font-family: 'Helvetica Neue', Arial, sans-serif; background: #fff; margin: 0; padding: 40px; color: #333; }
      .container { max-width: 800px; margin: 0 auto; }
      .report-header h1 { font-weight: 300; font-size: 3em; margin: 0 0 10px 0; }
      .report-section { margin: 60px 0; }
      .report-table { width: 100%; margin: 30px 0; }
      .report-table th { font-weight: 500; padding: 15px 0; border-bottom: 2px solid #333; text-align: left; }
      .report-table td { padding: 15px 0; border-bottom: 1px solid #ddd; }
    `,
    dashboard: `
      body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #1e1e1e; margin: 0; padding: 20px; color: #fff; }
      .container { max-width: 1400px; margin: 0 auto; }
      .report-header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 30px; border-radius: 10px; margin-bottom: 30px; }
      .report-section { background: #2d2d2d; padding: 25px; border-radius: 8px; margin-bottom: 20px; }
      .card-container { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; }
      .card { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 25px; border-radius: 10px; text-align: center; }
    `
  };

  return themes[theme] || themes.modern;
}

/**
 * Escape HTML special characters
 */
function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
