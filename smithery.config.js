/**
 * Smithery Build Configuration
 * https://smithery.ai/docs/build/deployments
 * 
 * This file customizes the esbuild options for Smithery builds.
 */

export default {
  esbuild: {
    // Mark packages with native bindings as external to prevent bundling issues
    external: [
      "puppeteer",
      "puppeteer-core", 
      "html-pdf-node",
      "@azure/identity",
      "@azure/msal-node"
    ],
    
    // Enable minification for production builds
    minify: true,
    
    // Target Node.js 18+ for compatibility
    target: "node18",
    
    // Keep names for better error stack traces
    keepNames: true
  }
};
