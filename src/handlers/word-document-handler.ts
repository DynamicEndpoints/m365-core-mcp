/**
 * Word Document Handler
 * Creates professional Word documents from data using Microsoft Graph API
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import {
  WordDocumentArgs,
  WordSection,
  WordTemplate,
  TableData
} from '../types/document-generation-types.js';

/**
 * Handle Word document operations
 */
export async function handleWordDocuments(
  args: WordDocumentArgs,
  graphClient: Client
): Promise<string> {
  try {
    switch (args.action) {
      case 'create':
        return await createDocument(args, graphClient);
      case 'get':
        return await getDocument(args, graphClient);
      case 'list':
        return await listDocuments(args, graphClient);
      case 'export':
        return await exportDocument(args, graphClient);
      case 'append':
        return await appendToDocument(args, graphClient);
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
      `Word document operation failed: ${error instanceof Error ? error.message : 'Unknown error'}`
    );
  }
}

/**
 * Create a new Word document
 */
async function createDocument(
  args: WordDocumentArgs,
  graphClient: Client
): Promise<string> {
  if (!args.fileName) {
    throw new McpError(ErrorCode.InvalidRequest, 'fileName is required for create action');
  }

  if (!args.sections || args.sections.length === 0) {
    throw new McpError(ErrorCode.InvalidRequest, 'At least one section is required');
  }

  // Determine drive location (default to user's OneDrive)
  const driveId = args.driveId || 'me';
  const folderPath = args.folderId ? `/items/${args.folderId}` : '/root';

  // Create Word document file
  const fileName = args.fileName.endsWith('.docx') ? args.fileName : `${args.fileName}.docx`;
  
  // Generate document content
  const documentContent = generateDocumentXML(args.sections, args.template);
  
  const uploadedFile = await graphClient
    .api(`/drives/${driveId}${folderPath}:/${fileName}:/content`)
    .header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    .put(documentContent);

  return JSON.stringify({
    success: true,
    fileId: uploadedFile.id,
    fileName: uploadedFile.name,
    webUrl: uploadedFile.webUrl,
    driveId: uploadedFile.parentReference?.driveId,
    message: `Word document "${fileName}" created successfully with ${args.sections.length} sections`
  }, null, 2);
}

/**
 * Get an existing Word document
 */
async function getDocument(
  args: WordDocumentArgs,
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
 * List Word documents
 */
async function listDocuments(
  args: WordDocumentArgs,
  graphClient: Client
): Promise<string> {
  const driveId = args.driveId || 'me';
  const folderPath = args.folderId ? `/items/${args.folderId}` : '/root';
  
  let query = graphClient
    .api(`/drives/${driveId}${folderPath}/children`)
    .filter("endsWith(name,'.docx')");

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
    documents: result.value?.map((file: any) => ({
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
 * Export document to different format
 */
async function exportDocument(
  args: WordDocumentArgs,
  graphClient: Client
): Promise<string> {
  if (!args.fileId) {
    throw new McpError(ErrorCode.InvalidRequest, 'fileId is required for export action');
  }

  if (!args.format) {
    throw new McpError(ErrorCode.InvalidRequest, 'format is required for export action');
  }

  const driveId = args.driveId || 'me';
  
  // Get file info
  const file = await graphClient
    .api(`/drives/${driveId}/items/${args.fileId}`)
    .get();

  // Convert format to MIME type
  const formatMap: Record<string, string> = {
    'pdf': 'application/pdf',
    'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'html': 'text/html',
    'txt': 'text/plain'
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
    message: `Document exported to ${args.format}`,
    originalFile: file.name,
    format: args.format,
    downloadUrl: file.webUrl
  }, null, 2);
}

/**
 * Append content to existing document
 */
async function appendToDocument(
  args: WordDocumentArgs,
  graphClient: Client
): Promise<string> {
  if (!args.fileId) {
    throw new McpError(ErrorCode.InvalidRequest, 'fileId is required for append action');
  }

  if (!args.content) {
    throw new McpError(ErrorCode.InvalidRequest, 'content is required for append action');
  }

  const driveId = args.driveId || 'me';

  // Note: This is a simplified implementation
  // For real append functionality, you would need to:
  // 1. Download the existing document
  // 2. Parse it
  // 3. Add new content
  // 4. Upload the modified version
  
  // For now, we'll use a simple approach with the Word API
  throw new McpError(
    ErrorCode.InvalidRequest,
    'Append functionality requires Office 365 Word API integration. ' +
    'Please use create action with complete content or manually edit the document.'
  );
}

/**
 * Generate Word document XML content (simplified version)
 * In production, use a library like docx or officegen
 */
function generateDocumentXML(sections: WordSection[], template?: WordTemplate): Buffer {
  // This is a simplified implementation
  // For production, integrate with docx library: https://docx.js.org/
  
  let content = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>`;

  // Add header if specified
  if (template?.header) {
    content += `
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Header"/>
      </w:pPr>
      <w:r>
        <w:t>${escapeXml(template.header)}</w:t>
      </w:r>
    </w:p>`;
  }

  // Add sections
  sections.forEach(section => {
    content += generateSectionXML(section);
  });

  // Add footer if specified
  if (template?.footer) {
    content += `
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Footer"/>
      </w:pPr>
      <w:r>
        <w:t>${escapeXml(template.footer)}</w:t>
      </w:r>
    </w:p>`;
  }

  content += `
  </w:body>
</w:document>`;

  return Buffer.from(content, 'utf-8');
}

/**
 * Generate XML for a document section
 */
function generateSectionXML(section: WordSection): string {
  let xml = '';

  switch (section.type) {
    case 'heading1':
    case 'heading2':
    case 'heading3':
      const headingLevel = section.type.replace('heading', '');
      xml = `
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading${headingLevel}"/>
      </w:pPr>
      <w:r>
        <w:t>${escapeXml(section.content || '')}</w:t>
      </w:r>
    </w:p>`;
      break;

    case 'paragraph':
      xml = `
    <w:p>
      <w:r>
        ${section.style?.bold ? '<w:rPr><w:b/></w:rPr>' : ''}
        <w:t>${escapeXml(section.content || '')}</w:t>
      </w:r>
    </w:p>`;
      break;

    case 'list':
      if (section.items) {
        xml = section.items.map(item => `
    <w:p>
      <w:pPr>
        <w:numPr>
          <w:ilvl w:val="0"/>
          <w:numId w:val="1"/>
        </w:numPr>
      </w:pPr>
      <w:r>
        <w:t>${escapeXml(item)}</w:t>
      </w:r>
    </w:p>`).join('');
      }
      break;

    case 'table':
      if (section.tableData) {
        xml = generateTableXML(section.tableData);
      }
      break;

    case 'pageBreak':
      xml = `
    <w:p>
      <w:r>
        <w:br w:type="page"/>
      </w:r>
    </w:p>`;
      break;
  }

  return xml;
}

/**
 * Generate table XML
 */
function generateTableXML(tableData: TableData): string {
  let xml = '<w:tbl>';
  
  // Add headers
  xml += '<w:tr>';
  tableData.headers.forEach(header => {
    xml += `
      <w:tc>
        <w:tcPr>
          <w:shd w:fill="D9D9D9"/>
        </w:tcPr>
        <w:p>
          <w:r>
            <w:rPr><w:b/></w:rPr>
            <w:t>${escapeXml(header)}</w:t>
          </w:r>
        </w:p>
      </w:tc>`;
  });
  xml += '</w:tr>';

  // Add rows
  tableData.rows.forEach(row => {
    xml += '<w:tr>';
    row.forEach(cell => {
      xml += `
      <w:tc>
        <w:p>
          <w:r>
            <w:t>${escapeXml(cell)}</w:t>
          </w:r>
        </w:p>
      </w:tc>`;
    });
    xml += '</w:tr>';
  });

  xml += '</w:tbl>';
  return xml;
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
