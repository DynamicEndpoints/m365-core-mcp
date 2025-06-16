import { handleDLPPolicies, handleDLPIncidents, handleDLPSensitivityLabels } from './handlers/dlp-handler.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { DLPPolicyArgs, DLPIncidentArgs, DLPSensitivityLabelArgs } from './types/dlp-types.js';

// Mock Graph client
const mockGraphClient = {
  api: (path: string) => ({
    get: async () => {
      console.log(`GET request to: ${path}`);
      if (path === '/beta/security/dataLossPreventionPolicies') {
        return { value: [] };
      } else if (path === '/security/alerts_v2') {
        return { value: [] };
      } else if (path === '/informationProtection/policy/labels') {
        return { value: [] };
      }
      return {};
    },
    post: async (data: any) => {
      console.log(`POST request to: ${path} with data:`, data);
      return { id: 'new-id', ...data };
    },
    patch: async (data: any) => {
      console.log(`PATCH request to: ${path} with data:`, data);
      return { id: 'updated-id', ...data };
    },
    delete: async () => {
      console.log(`DELETE request to: ${path}`);
      return {};
    }
  })
} as unknown as Client;

// Test policy handler
async function testDLPPolicies() {
  console.log('\n--- Testing DLP Policies ---');
  
  // Test listing policies
  const listArgs: DLPPolicyArgs = { action: 'list' };
  console.log('List policies:');
  const listResult = await handleDLPPolicies(mockGraphClient, listArgs);
  console.log('Result:', JSON.stringify(listResult, null, 2));
  
  // Test getting a policy
  const getArgs: DLPPolicyArgs = { action: 'get', policyId: 'test-policy-id' };
  console.log('\nGet policy:');
  const getResult = await handleDLPPolicies(mockGraphClient, getArgs);
  console.log('Result:', JSON.stringify(getResult, null, 2));
  
  // Test creating a policy
  const createArgs: DLPPolicyArgs = { 
    action: 'create', 
    name: 'Test Policy',
    description: 'This is a test policy',
    settings: { enabled: true }
  };
  console.log('\nCreate policy:');
  const createResult = await handleDLPPolicies(mockGraphClient, createArgs);
  console.log('Result:', JSON.stringify(createResult, null, 2));
}

// Test incident handler
async function testDLPIncidents() {
  console.log('\n--- Testing DLP Incidents ---');
  
  // Test listing incidents
  const listArgs: DLPIncidentArgs = { action: 'list' };
  console.log('List incidents:');
  const listResult = await handleDLPIncidents(mockGraphClient, listArgs);
  console.log('Result:', JSON.stringify(listResult, null, 2));
  
  // Test getting an incident
  const getArgs: DLPIncidentArgs = { action: 'get', incidentId: 'test-incident-id' };
  console.log('\nGet incident:');
  const getResult = await handleDLPIncidents(mockGraphClient, getArgs);
  console.log('Result:', JSON.stringify(getResult, null, 2));
  
  // Test resolving an incident
  const resolveArgs: DLPIncidentArgs = { action: 'resolve', incidentId: 'test-incident-id' };
  console.log('\nResolve incident:');
  const resolveResult = await handleDLPIncidents(mockGraphClient, resolveArgs);
  console.log('Result:', JSON.stringify(resolveResult, null, 2));
}

// Test sensitivity label handler
async function testDLPSensitivityLabels() {
  console.log('\n--- Testing DLP Sensitivity Labels ---');
  
  // Test listing labels
  const listArgs: DLPSensitivityLabelArgs = { action: 'list' };
  console.log('List labels:');
  const listResult = await handleDLPSensitivityLabels(mockGraphClient, listArgs);
  console.log('Result:', JSON.stringify(listResult, null, 2));
  
  // Test getting a label
  const getArgs: DLPSensitivityLabelArgs = { action: 'get', labelId: 'test-label-id' };
  console.log('\nGet label:');
  const getResult = await handleDLPSensitivityLabels(mockGraphClient, getArgs);
  console.log('Result:', JSON.stringify(getResult, null, 2));
  
  // Test creating a label
  const createArgs: DLPSensitivityLabelArgs = { 
    action: 'create', 
    name: 'Test Label',
    description: 'This is a test label',
    settings: { color: 'red', sensitivity: 3 }
  };
  console.log('\nCreate label:');
  const createResult = await handleDLPSensitivityLabels(mockGraphClient, createArgs);
  console.log('Result:', JSON.stringify(createResult, null, 2));
}

// Run tests
async function runTests() {
  try {
    await testDLPPolicies();
    await testDLPIncidents();
    await testDLPSensitivityLabels();
    console.log('\nAll tests completed successfully!');
  } catch (error) {
    console.error('Test failed:', error);
  }
}

runTests();
