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
 * Wraps a handler function to ensure its response is properly formatted
 * @param handler The original handler function
 * @returns A wrapped handler that ensures proper response formatting
 */
export function wrapToolHandler<T, R>(
  handler: (args: T) => Promise<{ content: { type: string; text: string; }[]; isError?: boolean }>
): (args: T) => Promise<R> {
  return async (args: T): Promise<R> => {
    try {
      const result = await handler(args);
      return {
        content: result.content.map(item => ({
          type: "text" as const,
          text: item.text
        })),
        ...(result.isError ? { isError: true } : {})
      } as unknown as R;
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new Error(`Error executing tool: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  };
}
