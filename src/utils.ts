import { McpError } from '@modelcontextprotocol/sdk/types.js';

/**
 * Formats a text response for MCP tools according to the SDK 1.12.0 requirements
 * @param text The text content to format
 * @param isError Whether this is an error response
 * @returns A properly formatted MCP tool response
 */
export function formatTextResponse(text: string, isError: boolean = false): {
  content: { type: "text"; text: string; }[];
  isError?: boolean;
} {
  return {
    content: [{ 
      type: "text" as const, 
      text 
    }],
    ...(isError ? { isError: true } : {})
  };
}

/**
 * Formats a JSON response for MCP tools with proper structure
 * @param data The data object to format
 * @param message Optional success message
 * @returns A properly formatted MCP tool response
 */
export function formatJsonResponse(data: any, message?: string): {
  content: { type: "text"; text: string; }[];
} {
  const responseText = message 
    ? `${message}\n\n${JSON.stringify(data, null, 2)}`
    : JSON.stringify(data, null, 2);
    
  return {
    content: [{ 
      type: "text" as const, 
      text: responseText
    }]
  };
}

/**
 * Validates that a response follows proper MCP format
 * @param response The response to validate
 * @returns True if valid, throws error if invalid
 */
export function validateMcpResponse(response: any): boolean {
  if (!response || typeof response !== 'object') {
    throw new Error('Response must be an object');
  }
  
  if (!response.content || !Array.isArray(response.content)) {
    throw new Error('Response must have a content array');
  }
  
  for (const item of response.content) {
    if (!item.type || !item.text) {
      throw new Error('Each content item must have type and text properties');
    }
  }
  
  return true;
}

/**
 * Creates a standardized error response for MCP tools
 * @param error The error to format
 * @param toolName Optional tool name for context
 * @returns A properly formatted MCP error response
 */
export function formatErrorResponse(error: any, toolName?: string): {
  content: { type: "text"; text: string; }[];
  isError: boolean;
} {
  const errorMessage = error instanceof Error ? error.message : String(error);
  const contextMessage = toolName ? `Error in ${toolName}: ${errorMessage}` : errorMessage;
  
  return {
    content: [{ 
      type: "text" as const, 
      text: contextMessage
    }],
    isError: true
  };
}

/**
 * Wraps a handler function to ensure its response is properly formatted
 * 
 * @param handler The original handler function
 * @returns A wrapped handler that ensures proper response formatting
 */
export function wrapToolHandler<T>(
  handler: (args: T) => Promise<{ content: { type: string; text: string; }[]; isError?: boolean }>
): (args: T) => Promise<{ content: { type: "text"; text: string; }[]; isError?: boolean }> {
  return async (args: T) => {
    try {
      const result = await handler(args);
      
      // Validate the response format
      validateMcpResponse(result);
      
      return {
        content: result.content.map(item => ({
          type: "text" as const,
          text: item.text
        })),
        ...(result.isError ? { isError: true } : {})
      };
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new Error(`Error executing tool: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  };
}
