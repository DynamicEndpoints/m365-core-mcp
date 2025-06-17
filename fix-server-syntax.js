import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const serverPath = path.join(__dirname, 'src', 'server.ts');
let content = fs.readFileSync(serverPath, 'utf8');

console.log('Fixing syntax errors in server.ts...');

// Fix missing variable names in const declarations
const fixes = [
  { pattern: /const  = await this\.getGraphClient\(\)\s+\.api\(`\/sites\/\$\{variables\.siteId\}\/lists`\)/, replacement: 'const lists = await this.getGraphClient().api(`/sites/${variables.siteId}/lists`)' },
  { pattern: /const  = await this\.getGraphClient\(\)\s+\.api\(`\/sites\/\$\{variables\.siteId\}\/lists\/\$\{variables\.listId\}`\)/, replacement: 'const list = await this.getGraphClient().api(`/sites/${variables.siteId}/lists/${variables.listId}`)' },
  { pattern: /const  = await this\.getGraphClient\(\)\s+\.api\(`\/sites\/\$\{variables\.siteId\}\/lists\/\$\{variables\.listId\}\/items\?expand=fields`\)/, replacement: 'const items = await this.getGraphClient().api(`/sites/${variables.siteId}/lists/${variables.listId}/items?expand=fields`)' },
  { pattern: /const  = await this\.getGraphClient\(\)\s+\.api\(`\/users\/\$\{variables\.userId\}`\)/, replacement: 'const user = await this.getGraphClient().api(`/users/${variables.userId}`)' },
  { pattern: /const  = await this\.getGraphClient\(\)\s+\.api\(`\/groups\/\$\{variables\.groupId\}`\)/, replacement: 'const group = await this.getGraphClient().api(`/groups/${variables.groupId}`)' },
  { pattern: /const  = await this\.getGraphClient\(\)\s+\.api\(`\/applications\/\$\{variables\.appId\}`\)/, replacement: 'const app = await this.getGraphClient().api(`/applications/${variables.appId}`)' },
  { pattern: /const  = await this\.getGraphClient\(\)\s+\.api\(`\/devices\/\$\{variables\.deviceId\}`\)/, replacement: 'const device = await this.getGraphClient().api(`/devices/${variables.deviceId}`)' },
  { pattern: /const  = await this\.getGraphClient\(\)\s+\.api\(`\/servicePrincipals\/\$\{variables\.spId\}`\)/, replacement: 'const sp = await this.getGraphClient().api(`/servicePrincipals/${variables.spId}`)' },
  { pattern: /const  = await this\.getGraphClient\(\)\s+\.api\(`\/directoryRoles\/\$\{variables\.roleId\}`\)/, replacement: 'const role = await this.getGraphClient().api(`/directoryRoles/${variables.roleId}`)' },
  { pattern: /const  = await this\.getGraphClient\(\)\s+\.api\(`\/security\/alerts_v2\/\$\{variables\.alertId\}`\)/, replacement: 'const alert = await this.getGraphClient().api(`/security/alerts_v2/${variables.alertId}`)' },
  { pattern: /const  = await this\.getGraphClient\(\)\s+\.api\(`\/security\/incidents\/\$\{variables\.incidentId\}`\)/, replacement: 'const incident = await this.getGraphClient().api(`/security/incidents/${variables.incidentId}`)' },
  { pattern: /const  = await this\.getGraphClient\(\)\s+\.api\(`\/deviceManagement\/managedDevices\/\$\{variables\.deviceId\}`\)/, replacement: 'const device = await this.getGraphClient().api(`/deviceManagement/managedDevices/${variables.deviceId}`)' },
  { pattern: /const  = await this\.getGraphClient\(\)\s+\.api\(`\/deviceManagement\/deviceConfigurations\/\$\{variables\.configId\}`\)/, replacement: 'const config = await this.getGraphClient().api(`/deviceManagement/deviceConfigurations/${variables.configId}`)' },
  { pattern: /const  = await this\.getGraphClient\(\)\s+\.api\(`\/deviceManagement\/deviceCompliancePolicies\/\$\{variables\.policyId\}`\)/, replacement: 'const policy = await this.getGraphClient().api(`/deviceManagement/deviceCompliancePolicies/${variables.policyId}`)' },
  { pattern: /const  = await this\.getGraphClient\(\)\s+\.api\(`\/informationProtection\/policy\/labels\/\$\{variables\.labelId\}`\)/, replacement: 'const label = await this.getGraphClient().api(`/informationProtection/policy/labels/${variables.labelId}`)' },
  { pattern: /const  = await this\.getGraphClient\(\)\s+\.api\(`\/deviceAppManagement\/mobileApps\/\$\{variables\.appId\}`\)/, replacement: 'const app = await this.getGraphClient().api(`/deviceAppManagement/mobileApps/${variables.appId}`)' }
];

// Apply all fixes
fixes.forEach(fix => {
  const matches = content.match(fix.pattern);
  if (matches) {
    console.log(`Fixing: ${matches[0].substring(0, 50)}...`);
    content = content.replace(fix.pattern, fix.replacement);
  }
});

// Fix any remaining const  = patterns with generic variable names
let counter = 1;
content = content.replace(/const  = await/g, () => `const data${counter++} = await`);

// Fix malformed method declarations
content = content.replace(/async \(uri: URL, variables\) => \{        try \{/g, 'async (uri: URL, variables) => {\n        try {');

// Fix other structural issues
content = content.replace(/\}\s*\)\s*;\s*\}\s*\)\s*;/g, '}\n      )\n    );');

// Write the fixed file
fs.writeFileSync(serverPath, content);
console.log('Fixed syntax errors in server.ts');

// Run a basic validation
try {
  console.log('\nSkipping TypeScript validation (would require TypeScript module)');
} catch (error) {
  console.log('\nCould not validate TypeScript syntax:', error.message);
}
