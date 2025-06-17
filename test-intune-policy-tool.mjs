import { M365CoreServer } from './build/server.js';
import assert from 'assert';

async function runIntunePolicyTests() {
  console.log('Starting Intune policy tool tests...');

  const server = new M365CoreServer();

  // Test case 1: Create a simple macOS Compliance Policy
  const macosPolicyArgs = {
    platform: 'macos',
    policyType: 'Compliance',
    displayName: 'Test macOS Compliance Policy',
    description: 'A test policy for macOS compliance.',
    settings: {
      compliance: {
        passwordRequired: true,
        passwordMinimumLength: 8,
        storageRequireEncryption: true
      }
    }
  };

  try {
    console.log('Test 1: Creating macOS Compliance Policy...');
    const macosResult = await server.server.tools['createIntunePolicy'].handler(macosPolicyArgs);
    assert(macosResult, 'macOS policy creation should return a result.');
    assert(macosResult.content[0].text.includes('Test macOS Compliance Policy'), 'macOS policy should have the correct display name.');
    console.log('Test 1 PASSED.');
  } catch (error) {
    console.error('Test 1 FAILED:', error);
  }

  // Test case 2: Create a simple Windows Configuration Policy
  const windowsPolicyArgs = {
    platform: 'windows',
    policyType: 'Configuration',
    displayName: 'Test Windows Configuration Policy',
    description: 'A test policy for Windows configuration.',
    settings: {
      windows10GeneralConfiguration: {
        passwordBlockSimple: true,
        passwordMinimumLength: 6
      }
    }
  };

  try {
    console.log('Test 2: Creating Windows Configuration Policy...');
    const windowsResult = await server.server.tools['createIntunePolicy'].handler(windowsPolicyArgs);
    assert(windowsResult, 'Windows policy creation should return a result.');
    assert(windowsResult.content[0].text.includes('Test Windows Configuration Policy'), 'Windows policy should have the correct display name.');
    console.log('Test 2 PASSED.');
  } catch (error) {
    console.error('Test 2 FAILED:', error);
  }

  console.log('Intune policy tool tests completed.');
  // In a real scenario, you might want to clean up the created policies.
}

runIntunePolicyTests();
