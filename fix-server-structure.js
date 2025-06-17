import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const serverPath = path.join(__dirname, 'src', 'server.ts');
let content = fs.readFileSync(serverPath, 'utf8');

console.log('Fixing major structural issues in server.ts...');

// Remove duplicate closing brackets and malformed structure
content = content.replace(/\}\s*\)\s*;\s*\}\s*;\s*\}\s*catch \(error\) \{\s*throw new McpError\(/g, '});');

// Fix the duplicate error handling sections
content = content.replace(/\}\s*\)\s*;\s*\}\s*;\s*\}\s*catch/g, '}\n    );\n  }\n\n  // Method continues...\n  } catch');

// Fix the malformed resource declarations
content = content.replace(/\s*\}\s*\}\s*\)\s*;\s*this\.server\.resource\(/g, '\n      }\n    );\n\n    this.server.resource(');

// Fix method declarations that are outside the class
const methodDeclarationPattern = /^(\s*)private async handle(\w+)\(/gm;
content = content.replace(methodDeclarationPattern, '  private async handle$2(');

const publicMethodPattern = /^(\s*)public (addSSEClient|removeSSEClient|broadcastUpdate|reportProgress|completeOperation|notifyResourceChange)\(/gm;
content = content.replace(publicMethodPattern, '  public $2(');

// Fix try-catch blocks that are malformed
content = content.replace(/\btry \{(\s*)this\.reportProgress/g, 'try {\n      this.reportProgress');

// Fix const declarations within methods
content = content.replace(/(\s*)const operationId = randomUUID\(\);(\s*)try \{/g, '$1const operationId = randomUUID();\n$1\n$1try {');

// Fix for loops
content = content.replace(/for \(let i = 0; i < total; i\+\+\) \{(\s*)\/\/ Process each item/g, 'for (let i = 0; i < total; i++) {\n        // Process each item');

// Fix object literals
content = content.replace(/const (progressUpdate|completion|notification) = \{(\s*)type:/g, 'const $1 = {\n      type:');

// Ensure proper class structure around the problematic methods
const classMethodsSection = `
  // --- Tool Handlers ---

  private async handleDistributionList(args: DistributionListArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    // Handler logic for managing distribution lists
    // This is a placeholder implementation - replace with actual logic
    return {
      content: [
        {
          type: "text",
          text: \`Handled Distribution List: \${JSON.stringify(args, null, 2)}\`,
        },
      ],
    };
  }

  private async handleSecurityGroup(args: SecurityGroupArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    // Handler logic for managing security groups
    // This is a placeholder implementation - replace with actual logic
    return {
      content: [
        {
          type: "text",
          text: \`Handled Security Group: \${JSON.stringify(args, null, 2)}\`,
        },
      ],
    };
  }

  private async handleM365Group(args: M365GroupArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    // Handler logic for managing M365 groups
    // This is a placeholder implementation - replace with actual logic
    return {
      content: [
        {
          type: "text",
          text: \`Handled M365 Group: \${JSON.stringify(args, null, 2)}\`,
        },
      ],
    };
  }

  private async handleAzureAdRoles(args: AzureAdRoleArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    // Handler logic for managing Azure AD roles
    // This is a placeholder implementation - replace with actual logic
    return {
      content: [
        {
          type: "text",
          text: \`Handled Azure AD Role: \${JSON.stringify(args, null, 2)}\`,
        },
      ],
    };
  }

  private async handleAzureAdApps(args: AzureAdAppArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    // Handler logic for managing Azure AD apps
    // This is a placeholder implementation - replace with actual logic
    return {
      content: [
        {
          type: "text",
          text: \`Handled Azure AD App: \${JSON.stringify(args, null, 2)}\`,
        },
      ],
    };
  }

  private async handleAzureAdDevices(args: AzureAdDeviceArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    // Handler logic for managing Azure AD devices
    // This is a placeholder implementation - replace with actual logic
    return {
      content: [
        {
          type: "text",
          text: \`Handled Azure AD Device: \${JSON.stringify(args, null, 2)}\`,
        },
      ],
    };
  }

  private async handleServicePrincipals(args: AzureAdSpArgs): Promise<{ content: { type: string; text: string; }[]; }> {
    // Handler logic for managing service principals
    // This is a placeholder implementation - replace with actual logic
    return {
      content: [
        {
          type: "text",
          text: \`Handled Service Principal: \${JSON.stringify(args, null, 2)}\`,
        },
      ],
    };
  }

  // Enhanced tool handler with progress reporting for bulk operations
  private async handleBulkOperation(args: any, operationType: string): Promise<string> {
    const operationId = randomUUID();
    
    try {
      this.reportProgress(operationId, 0, \`Starting \${operationType} operation\`);
      
      // Simulate bulk operation with progress updates
      const items = args.items || [];
      const total = items.length;
      
      for (let i = 0; i < total; i++) {
        // Process each item
        const progress = Math.round(((i + 1) / total) * 100);
        this.reportProgress(operationId, progress, \`Processing item \${i + 1} of \${total}\`);
        
        // Simulate processing time
        await new Promise(resolve => setTimeout(resolve, 100));
      }
      
      const result = \`Completed \${operationType} operation for \${total} items\`;
      this.completeOperation(operationId, result);
      
      return result;
    } catch (error) {
      this.completeOperation(operationId, { error: error instanceof Error ? error.message : 'Unknown error' });
      throw error;
    }
  }

  // Enhanced distribution list handler with progress reporting
  private async handleDistributionListWithProgress(args: DistributionListArgs): Promise<any> {
    const operationId = randomUUID();
    
    try {
      this.reportProgress(operationId, 0, \`Starting distribution list operation: \${args.action}\`);
      
      // Execute the actual operation
      this.reportProgress(operationId, 50, 'Executing distribution list operation...');
      const result = await this.handleDistributionList(args);
      
      this.reportProgress(operationId, 100, 'Distribution list operation completed');
      this.completeOperation(operationId, result);
      
      // Notify about resource changes for create, update, delete operations
      if (args.action === 'create' || args.action === 'update' || args.action === 'delete') {
        this.notifyResourceChange(\`m365://distribution-lists/\${args.listId || 'new'}\`,
          args.action === 'create' ? 'created' : args.action === 'update' ? 'updated' : 'deleted');
      }
      
      return result;
    } catch (error) {
      this.completeOperation(operationId, { error: error instanceof Error ? error.message : 'Unknown error' });
      throw error;
    }
  }

  // --- Real-time and Progress Reporting Methods ---

  public addSSEClient(client: any): void {
    this.sseClients.add(client);
    console.log(\`SSE client connected. Total clients: \${this.sseClients.size}\`);
  }

  public removeSSEClient(client: any): void {
    this.sseClients.delete(client);
    console.log(\`SSE client disconnected. Total clients: \${this.sseClients.size}\`);
  }

  public broadcastUpdate(update: any): void {
    this.sseClients.forEach(client => {
      try {
        client.write(\`data: \${JSON.stringify(update)}\\n\\n\`);
      } catch (error) {
        // Remove disconnected clients
        this.sseClients.delete(client);
      }
    });
  }

  public reportProgress(operationId: string, progress: number, message?: string): void {
    const progressUpdate = {
      type: 'progress',
      operationId,
      progress,
      message,
      timestamp: new Date().toISOString()
    };
    
    this.progressTrackers.set(operationId, progressUpdate);
    this.broadcastUpdate(progressUpdate);
  }

  public completeOperation(operationId: string, result: any): void {
    const completion = {
      type: 'completion',
      operationId,
      result,
      timestamp: new Date().toISOString()
    };
    
    this.progressTrackers.delete(operationId);
    this.broadcastUpdate(completion);
  }

  public notifyResourceChange(resourceUri: string, changeType: 'created' | 'updated' | 'deleted'): void {
    const notification = {
      type: 'resourceChange',
      resourceUri,
      changeType,
      timestamp: new Date().toISOString()
    };
    
    this.broadcastUpdate(notification);
  }
}`;

// Find where to place the class methods section and replace it
const classEndPattern = /\/\/ --- Tool Handlers ---[\s\S]*?}\s*$/;
if (classEndPattern.test(content)) {
  content = content.replace(classEndPattern, classMethodsSection);
} else {
  // If pattern not found, append at the end of the class
  content = content.replace(/}\s*$/, classMethodsSection);
}

// Write the fixed file
fs.writeFileSync(serverPath, content);
console.log('Fixed major structural issues in server.ts');
