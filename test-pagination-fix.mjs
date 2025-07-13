import { createAuthenticatedClient } from './src/utils/modern-graph-client.js';
import { handleCallMicrosoftApi } from './src/handlers.js';
import assert from 'assert';

async function runTest() {
  console.log('Starting pagination fix validation test...');

  try {
    // Obtain a Graph client
    const graphClient = await createAuthenticatedClient();

    // Define arguments to fetch all users, simulating the original scenario
    const args = {
      apiType: 'graph',
      fetchAll: true,
      method: 'get',
      path: '/users',
      responseFormat: 'json',
      selectFields: ["id", "displayName", "userPrincipalName"]
    };

    // Capture console output to check for errors
    const originalConsoleError = console.error;
    let consoleErrorOutput = '';
    console.error = (message) => {
      consoleErrorOutput += message;
    };

    // Execute the API call
    const result = await handleCallMicrosoftApi(graphClient, args, () => Promise.resolve(''), {});

    // Restore console.error
    console.error = originalConsoleError;

    // 1. Check for JSON parsing errors in console output
    assert.strictEqual(consoleErrorOutput.includes('Unexpected token'), false, `Test failed: Found JSON parsing errors in console output: ${consoleErrorOutput}`);
    console.log('✔ Test Passed: No JSON parsing errors were found in the console output.');

    // 2. Check if the result content is valid
    assert.ok(result.content, 'Test failed: Result content is missing.');
    assert.strictEqual(Array.isArray(result.content), true, 'Test failed: Result content is not an array.');
    assert.strictEqual(result.content.length > 0, true, 'Test failed: Result content is empty.');
    console.log('✔ Test Passed: Result content is valid and not empty.');

    // 3. Check if the response was parsed correctly
    const responseText = result.content[0].text;
    let parsedResponse;
    try {
      parsedResponse = JSON.parse(responseText.substring(responseText.indexOf('{')));
    } catch (e) {
      assert.fail(`Test failed: Could not parse the final JSON response. Error: ${e.message}`);
    }
    
    assert.ok(parsedResponse.value, 'Test failed: Parsed response does not contain a "value" property.');
    assert.ok(Array.isArray(parsedResponse.value), 'Test failed: Parsed response "value" is not an array.');
    console.log(`✔ Test Passed: Successfully parsed the final JSON response and found ${parsedResponse.value.length} users.`);

    console.log('✔ All validation checks passed. The pagination fix is working correctly.');

  } catch (error) {
    console.error('Test failed with an unexpected error:', error);
  }
}

runTest();
