/**
 * Test script for custom functions registration
 * Run this to verify all functions are properly configured
 */

import { AIANALYZE, AIGENERATE, AIEXPLAIN, AISUMMARIZE, AIUSAGE } from './functions';

// Mock Office context for testing
if (typeof Office === 'undefined') {
  (global as any).Office = {
    context: {
      document: {
        settings: {
          get: () => null,
          set: () => {},
          saveAsync: (callback: any) => callback({ status: 'succeeded' }),
          remove: () => {}
        }
      }
    },
    onReady: (callback: any) => callback(),
    AsyncResultStatus: {
      Succeeded: 'succeeded'
    }
  };
}

// Test data
const testData = [
  ['Product', 'Sales', 'Profit'],
  ['Widget A', 1000, 250],
  ['Widget B', 1500, 400],
  ['Widget C', 800, 150]
];

// Test functions
async function runTests() {
  console.log('Testing Excel Custom Functions...\n');

  // Test AIANALYZE
  console.log('1. Testing AIANALYZE...');
  try {
    const result = await AIANALYZE(testData, 'Analyze sales performance');
    console.log('✓ AIANALYZE:', result.substring(0, 100) + '...');
  } catch (error) {
    console.error('✗ AIANALYZE failed:', error);
  }

  // Test AIGENERATE
  console.log('\n2. Testing AIGENERATE...');
  try {
    const result = await AIGENERATE('Generate a product description for Widget A');
    console.log('✓ AIGENERATE:', result.substring(0, 100) + '...');
  } catch (error) {
    console.error('✗ AIGENERATE failed:', error);
  }

  // Test AIEXPLAIN
  console.log('\n3. Testing AIEXPLAIN...');
  try {
    const result = await AIEXPLAIN('=VLOOKUP(A2,B:C,2,FALSE)');
    console.log('✓ AIEXPLAIN:', result.substring(0, 100) + '...');
  } catch (error) {
    console.error('✗ AIEXPLAIN failed:', error);
  }

  // Test AISUMMARIZE
  console.log('\n4. Testing AISUMMARIZE...');
  try {
    const result = await AISUMMARIZE(testData, 50);
    console.log('✓ AISUMMARIZE:', result.substring(0, 100) + '...');
  } catch (error) {
    console.error('✗ AISUMMARIZE failed:', error);
  }

  // Test AIUSAGE
  console.log('\n5. Testing AIUSAGE...');
  try {
    const result = await AIUSAGE();
    console.log('✓ AIUSAGE:', result);
  } catch (error) {
    console.error('✗ AIUSAGE failed:', error);
  }

  console.log('\nAll tests completed!');
}

// Verify custom functions metadata
async function verifyMetadata() {
  console.log('\nVerifying Custom Functions Metadata...\n');
  
  // List of expected functions
  const expectedFunctions = [
    'AIANALYZE',
    'AIGENERATE', 
    'AIEXPLAIN',
    'AISUMMARIZE',
    'AIUSAGE'
  ];

  console.log('Expected functions:');
  expectedFunctions.forEach(fn => {
    console.log(`- ${fn}`);
  });

  console.log('\nTo verify in Excel:');
  console.log('1. Open Excel with the add-in loaded');
  console.log('2. In any cell, type "=AI" to see autocomplete suggestions');
  console.log('3. All functions should appear with descriptions');
}

// Run tests if executed directly
if (require.main === module) {
  console.log('ExcelAIRate Custom Functions Test Suite\n');
  console.log('Note: This test requires authentication.');
  console.log('Please ensure you are signed in via the task pane.\n');
  
  runTests().then(() => {
    verifyMetadata();
  }).catch(error => {
    console.error('Test suite failed:', error);
  });
}

export { runTests, verifyMetadata };