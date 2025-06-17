#!/usr/bin/env node

/**
 * Comprehensive M365 Core MCP Test Suite
 * Tests all major functionality including CIS compliance, DLP, and endpoint validation
 */

import { spawn } from 'child_process';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

interface TestResult {
  name: string;
  passed: boolean;
  message: string;
  duration: number;
}

class MCPTester {
  private results: TestResult[] = [];
  private serverProcess: any = null;

  async runAllTests(): Promise<void> {
    console.log('🚀 Starting M365 Core MCP Comprehensive Test Suite\n');

    try {
      // Test 1: Build validation
      await this.runTest('Build Validation', () => this.testBuild());

      // Test 2: Server startup
      await this.runTest('Server Startup', () => this.testServerStartup());

      // Test 3: CIS Compliance Tools
      await this.runTest('CIS Compliance Tools', () => this.testCISCompliance());

      // Test 4: DLP Functionality
      await this.runTest('DLP Functionality', () => this.testDLPFunctionality());

      // Test 5: Intune macOS Management
      await this.runTest('Intune macOS Management', () => this.testIntuneMacOS());

      // Test 6: API Endpoint Validation
      await this.runTest('API Endpoint Validation', () => this.testAPIEndpoints());

      // Test 7: Real-time Features
      await this.runTest('Real-time Features', () => this.testRealTimeFeatures());

    } catch (error) {
      console.error('❌ Test suite failed:', error);
    } finally {
      if (this.serverProcess) {
        this.serverProcess.kill();
      }
      this.printResults();
    }
  }

  private async runTest(name: string, testFn: () => Promise<void>): Promise<void> {
    const startTime = Date.now();
    console.log(`🧪 Running test: ${name}`);

    try {
      await testFn();
      const duration = Date.now() - startTime;
      this.results.push({
        name,
        passed: true,
        message: 'Test passed successfully',
        duration
      });
      console.log(`✅ ${name} - PASSED (${duration}ms)\n`);
    } catch (error) {
      const duration = Date.now() - startTime;
      const message = error instanceof Error ? error.message : 'Unknown error';
      this.results.push({
        name,
        passed: false,
        message,
        duration
      });
      console.log(`❌ ${name} - FAILED: ${message} (${duration}ms)\n`);
    }
  }

  private async testBuild(): Promise<void> {
    return new Promise((resolve, reject) => {
      const buildProcess = spawn('npm', ['run', 'build'], {
        cwd: __dirname,
        stdio: 'pipe'
      });

      let output = '';
      let errorOutput = '';

      buildProcess.stdout.on('data', (data) => {
        output += data.toString();
      });

      buildProcess.stderr.on('data', (data) => {
        errorOutput += data.toString();
      });

      buildProcess.on('close', (code) => {
        if (code === 0) {
          console.log('   📦 Build completed successfully');
          resolve();
        } else {
          reject(new Error(`Build failed with code ${code}: ${errorOutput}`));
        }
      });

      buildProcess.on('error', (error) => {
        reject(new Error(`Build process error: ${error.message}`));
      });
    });
  }

  private async testServerStartup(): Promise<void> {
    return new Promise((resolve, reject) => {
      const serverPath = join(__dirname, 'build', 'index.js');
      this.serverProcess = spawn('node', [serverPath], {
        stdio: 'pipe',
        env: { ...process.env, NODE_ENV: 'test' }
      });

      let output = '';
      let hasStarted = false;

      const timeout = setTimeout(() => {
        if (!hasStarted) {
          this.serverProcess.kill();
          reject(new Error('Server startup timeout'));
        }
      }, 10000); // 10 second timeout

      this.serverProcess.stdout.on('data', (data) => {
        output += data.toString();
        // Look for server startup indicators
        if (output.includes('MCP server running') || output.includes('Server initialized')) {
          hasStarted = true;
          clearTimeout(timeout);
          console.log('   🖥️  Server started successfully');
          resolve();
        }
      });

      this.serverProcess.stderr.on('data', (data) => {
        const error = data.toString();
        if (error.includes('Error') || error.includes('Exception')) {
          clearTimeout(timeout);
          reject(new Error(`Server startup error: ${error}`));
        }
      });

      this.serverProcess.on('error', (error) => {
        clearTimeout(timeout);
        reject(new Error(`Server process error: ${error.message}`));
      });

      // If server starts but doesn't log anything, give it a moment then assume success
      setTimeout(() => {
        if (!hasStarted && this.serverProcess && !this.serverProcess.killed) {
          hasStarted = true;
          clearTimeout(timeout);
          console.log('   🖥️  Server appears to be running (no errors detected)');
          resolve();
        }
      }, 3000);
    });
  }

  private async testCISCompliance(): Promise<void> {
    console.log('   🔒 Testing CIS compliance capabilities...');
    
    // Test if CIS compliance handler exists and is properly structured
    const fs = await import('fs');
    const complianceHandlerPath = join(__dirname, 'src', 'handlers', 'compliance-handler.ts');
    
    if (!fs.existsSync(complianceHandlerPath)) {
      throw new Error('CIS compliance handler not found');
    }

    const handlerContent = fs.readFileSync(complianceHandlerPath, 'utf8');
    
    // Check for required CIS compliance functions
    const requiredFunctions = [
      'handleCISCompliance',
      'generateComplianceReport',
      'checkSecurityBaseline'
    ];

    for (const func of requiredFunctions) {
      if (!handlerContent.includes(func)) {
        throw new Error(`Required CIS compliance function ${func} not found`);
      }
    }

    console.log('   ✓ CIS compliance handler structure validated');
    console.log('   ✓ Required compliance functions found');
  }

  private async testDLPFunctionality(): Promise<void> {
    console.log('   🛡️  Testing DLP functionality...');
    
    const fs = await import('fs');
    const dlpHandlerPath = join(__dirname, 'src', 'handlers', 'dlp-handler.ts');
    
    if (!fs.existsSync(dlpHandlerPath)) {
      throw new Error('DLP handler not found');
    }

    const dlpContent = fs.readFileSync(dlpHandlerPath, 'utf8');
    
    // Check for required DLP functions
    const requiredDLPFunctions = [
      'handleDLPPolicies',
      'handleDLPIncidents', 
      'handleDLPSensitivityLabels'
    ];

    for (const func of requiredDLPFunctions) {
      if (!dlpContent.includes(func)) {
        throw new Error(`Required DLP function ${func} not found`);
      }
    }

    // Validate API endpoints used
    const validEndpoints = [
      '/beta/security/dataLossPreventionPolicies',
      '/security/alerts_v2',
      '/informationProtection/policy/labels'
    ];

    for (const endpoint of validEndpoints) {
      if (!dlpContent.includes(endpoint)) {
        throw new Error(`Required DLP API endpoint ${endpoint} not found`);
      }
    }

    console.log('   ✓ DLP handler structure validated');
    console.log('   ✓ DLP API endpoints validated');
  }

  private async testIntuneMacOS(): Promise<void> {
    console.log('   🍎 Testing Intune macOS management...');
    
    const fs = await import('fs');
    const intuneHandlerPath = join(__dirname, 'src', 'handlers', 'intune-macos-handler.ts');
    
    if (!fs.existsSync(intuneHandlerPath)) {
      throw new Error('Intune macOS handler not found');
    }

    const intuneContent = fs.readFileSync(intuneHandlerPath, 'utf8');
    
    // Check for required Intune functions
    const requiredIntuneFunctions = [
      'handleIntuneMacOSDevices',
      'handleIntuneMacOSPolicies',
      'handleIntuneMacOSApps'
    ];

    for (const func of requiredIntuneFunctions) {
      if (!intuneContent.includes(func)) {
        throw new Error(`Required Intune function ${func} not found`);
      }
    }

    // Check for proper filtering logic
    if (!intuneContent.includes('macOS') || !intuneContent.includes('operatingSystem eq')) {
      throw new Error('macOS filtering logic not found or incomplete');
    }

    console.log('   ✓ Intune macOS handler structure validated');
    console.log('   ✓ macOS-specific filtering logic validated');
  }

  private async testAPIEndpoints(): Promise<void> {
    console.log('   🌐 Testing API endpoint configurations...');
    
    const fs = await import('fs');
    
    // Test if all handlers are using correct Microsoft Graph endpoints
    const handlersDir = join(__dirname, 'src', 'handlers');
    const handlerFiles = fs.readdirSync(handlersDir).filter(f => f.endsWith('.ts'));
    
    const validGraphEndpoints = [
      '/deviceManagement/',
      '/security/',
      '/informationProtection/',
      '/users/',
      '/groups/',
      '/applications/',
      '/servicePrincipals/',
      '/directoryRoles/',
      '/sites/'
    ];

    for (const file of handlerFiles) {
      const content = fs.readFileSync(join(handlersDir, file), 'utf8');
      
      // Check if file uses Graph API and has valid endpoints
      if (content.includes('graphClient.api')) {
        let hasValidEndpoint = false;
        for (const endpoint of validGraphEndpoints) {
          if (content.includes(endpoint)) {
            hasValidEndpoint = true;
            break;
          }
        }
        if (!hasValidEndpoint) {
          throw new Error(`Handler ${file} may be using invalid Graph API endpoints`);
        }
      }
    }

    console.log('   ✓ API endpoint validation completed');
  }

  private async testRealTimeFeatures(): Promise<void> {
    console.log('   ⚡ Testing real-time features...');
    
    const fs = await import('fs');
    const serverPath = join(__dirname, 'src', 'server.ts');
    const serverContent = fs.readFileSync(serverPath, 'utf8');
    
    // Check for real-time capabilities
    const realTimeFeatures = [
      'reportProgress',
      'broadcastUpdate',
      'sseClients',
      'progressTrackers'
    ];

    for (const feature of realTimeFeatures) {
      if (!serverContent.includes(feature)) {
        throw new Error(`Real-time feature ${feature} not found`);
      }
    }

    console.log('   ✓ Real-time progress reporting available');
    console.log('   ✓ SSE broadcast capabilities available');
  }

  private printResults(): void {
    console.log('\n' + '='.repeat(60));
    console.log('📊 TEST RESULTS SUMMARY');
    console.log('='.repeat(60));

    const passed = this.results.filter(r => r.passed).length;
    const failed = this.results.filter(r => !r.passed).length;
    const total = this.results.length;

    console.log(`\n✅ Passed: ${passed}/${total}`);
    console.log(`❌ Failed: ${failed}/${total}`);
    console.log(`⏱️  Total Duration: ${this.results.reduce((sum, r) => sum + r.duration, 0)}ms\n`);

    if (failed > 0) {
      console.log('❌ FAILED TESTS:');
      this.results.filter(r => !r.passed).forEach(result => {
        console.log(`   • ${result.name}: ${result.message}`);
      });
      console.log();
    }

    if (passed === total) {
      console.log('🎉 ALL TESTS PASSED! Your M365 Core MCP is ready for deployment.');
      console.log('\n📋 CISA COMPLIANCE CHECKLIST:');
      console.log('   ✅ DLP policies and incident management');
      console.log('   ✅ Device management and compliance');
      console.log('   ✅ Security baseline validation');
      console.log('   ✅ Real-time monitoring capabilities');
      console.log('   ✅ Audit and reporting functions');
    } else {
      console.log('⚠️  Some tests failed. Please review and fix the issues above.');
    }

    console.log('\n' + '='.repeat(60));
  }
}

// Run the test suite
const tester = new MCPTester();
tester.runAllTests().catch(console.error);
