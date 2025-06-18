#!/usr/bin/env node

/**
 * Script to convert all validateCredentials() calls to lazy authentication
 */

import fs from 'fs';
import path from 'path';

const filePath = path.join(process.cwd(), 'src', 'server.ts');

console.log('🔄 Converting validateCredentials() to lazy authentication...\n');

try {
  // Read the file
  let content = fs.readFileSync(filePath, 'utf8');
  
  // Count occurrences before replacement
  const beforeCount = (content.match(/this\.validateCredentials\(\);/g) || []).length;
  
  console.log(`Found ${beforeCount} instances of this.validateCredentials()`);
  
  // Replace all occurrences in tool handlers (but not in getGraphClient)
  content = content.replace(
    /(\s+)this\.validateCredentials\(\);(\s+try \{)/g,
    '$1await this.ensureAuthenticated();$2'
  );
  
  // Special case for the getGraphClient method - keep the synchronous validation there
  content = content.replace(
    /(authProvider: async \(callback[^}]+)await this\.ensureAuthenticated\(\);/,
    '$1this.validateCredentials();'
  );
  
  // Count occurrences after replacement
  const afterCount = (content.match(/await this\.ensureAuthenticated\(\);/g) || []).length;
  const remainingValidateCount = (content.match(/this\.validateCredentials\(\);/g) || []).length;
  
  // Write the file back
  fs.writeFileSync(filePath, content, 'utf8');
  
  console.log(`✅ Conversion complete!`);
  console.log(`   - Converted ${beforeCount - remainingValidateCount} calls to lazy authentication`);
  console.log(`   - ${afterCount} tools now use await this.ensureAuthenticated()`);
  console.log(`   - ${remainingValidateCount} validateCredentials() calls remain (in auth provider)`);
  
} catch (error) {
  console.error('❌ Error during conversion:', error);
  process.exit(1);
}

console.log('\n🎉 Lazy loading conversion complete!');
console.log('All tools now authenticate on demand rather than at startup.');
