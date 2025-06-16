# Modern MCP Tool Handling Implementation

Based on analysis of official MCP servers (filesystem, git, time, memory, etc.), here's how to modernize your M365 Core MCP server tool handling:

## Current vs. Modern Pattern

### **Current Implementation (Older Pattern):**
```typescript
private setupTools(): void {
  this.server.tool(
    "manage_distribution_lists",
    distributionListSchema,
    wrapToolHandler(async (args: DistributionListArgs) => {
      return await this.handleDistributionList(args);
    })
  );
}
```

### **Modern MCP Implementation (Recommended):**
```typescript
import {
  ListToolsRequestSchema,
  CallToolRequestSchema,
  ToolSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { zodToJsonSchema } from "zod-to-json-schema";

const ToolInputSchema = ToolSchema.shape.inputSchema;
type ToolInput = z.infer<typeof ToolInputSchema>;

private setupTools(): void {
  // List all available tools
  this.server.setRequestHandler(ListToolsRequestSchema, async () => {
    return {
      tools: [
        {
          name: "manage_distribution_lists",
          description: "Manage distribution lists in Microsoft 365",
          inputSchema: zodToJsonSchema(distributionListSchema) as ToolInput,
        },
        {
          name: "manage_security_groups", 
          description: "Manage security groups in Microsoft 365",
          inputSchema: zodToJsonSchema(securityGroupSchema) as ToolInput,
        },
        {
          name: "manage_m365_groups",
          description: "Manage Microsoft 365 groups",
          inputSchema: zodToJsonSchema(m365GroupSchema) as ToolInput,
        },
        // ... all your other tools
      ],
    };
  });

  // Handle tool calls
  this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;
    
    try {
      switch (name) {
        case "manage_distribution_lists":
          const result = await this.handleDistributionList(args as DistributionListArgs);
          return { 
            content: [{ 
              type: "text", 
              text: typeof result === 'string' ? result : JSON.stringify(result, null, 2)
            }] 
          };
        
        case "manage_security_groups":
          const secResult = await this.handleSecurityGroup(args as SecurityGroupArgs);
          return { 
            content: [{ 
              type: "text", 
              text: typeof secResult === 'string' ? secResult : JSON.stringify(secResult, null, 2)
            }] 
          };
          
        // ... all your other tools
        
        default:
          throw new McpError(
            ErrorCode.InvalidRequest,
            `Unknown tool: ${name}`
          );
      }
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      return {
        content: [{ 
          type: "text", 
          text: `Error: ${error instanceof Error ? error.message : 'Unknown error'}` 
        }],
        isError: true,
      };
    }
  });
}
```

## Key Benefits of Modern Approach

### **1. Better Request/Response Handling**
- Explicit separation of tool listing and execution
- Standardized response format with `content` array
- Better error handling with `isError` flag

### **2. Improved Type Safety**
- Proper use of `zodToJsonSchema()` for input schemas
- Better TypeScript integration with request handlers
- Cleaner parameter handling

### **3. Enhanced Debugging**
- Clear request/response flow
- Better error messages and stack traces
- Easier to unit test individual handlers

### **4. Future Compatibility**
- Aligns with latest MCP specifications
- Ready for new MCP client features
- Better performance with modern clients

## Implementation Steps

### **Step 1: Update Imports**
```typescript
import {
  ListToolsRequestSchema,
  CallToolRequestSchema,
  ToolSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { zodToJsonSchema } from "zod-to-json-schema";
```

### **Step 2: Add Type Definitions**
```typescript
const ToolInputSchema = ToolSchema.shape.inputSchema;
type ToolInput = z.infer<typeof ToolInputSchema>;
```

### **Step 3: Update setupTools Method**
Replace the entire `setupTools()` method with the modern pattern shown above.

### **Step 4: Update Response Format**
Ensure all tool handlers return the proper format:
```typescript
return { 
  content: [{ 
    type: "text", 
    text: JSON.stringify(result, null, 2)
  }] 
};
```

### **Step 5: Test with MCP Inspector**
```bash
npx @modelcontextprotocol/inspector npm start
```

## Validation Checklist

- [ ] Replace `this.server.tool()` with `setRequestHandler()`
- [ ] Add `ListToolsRequestSchema` handler
- [ ] Add `CallToolRequestSchema` handler
- [ ] Update tool schemas with `zodToJsonSchema()`
- [ ] Modify response format to include `content` array
- [ ] Add proper error handling with `isError: true`
- [ ] Test all tools with MCP inspector
- [ ] Update documentation and examples

## Expected Improvements

✅ **Better Error Handling** - Standardized error responses  
✅ **Improved Type Safety** - Better TypeScript integration  
✅ **Future Compatibility** - Aligns with latest MCP specifications  
✅ **Easier Debugging** - Clearer request/response flow  
✅ **Better Testing** - Request handlers are easier to unit test  

This migration will make your M365 Core MCP server more maintainable, debuggable, and aligned with modern MCP best practices!
