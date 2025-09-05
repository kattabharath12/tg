#!/usr/bin/env node

// Simple test for wages extraction logic
console.log('🧪 Testing W2 wages extraction logic...\n');

// Simulate the OCR extraction function from our TypeScript code
function extractWagesFromOCR(ocrText) {
  console.log('🔍 [OCR] Searching for wages in OCR text...');
  
  // Multiple regex patterns to match Box 1 wages
  const wagePatterns = [
    // Pattern: "1 Wages, tips, other compensation 161130.48"
    /\b1\s+Wages[,\s]*tips[,\s]*other\s+compensation\s+([\d,]+\.?\d*)/i,
    // Pattern: "1. Wages, tips, other compensation: $161,130.48"
    /\b1\.?\s*Wages[,\s]*tips[,\s]*other\s+compensation[:\s]+\$?([\d,]+\.?\d*)/i,
    // Pattern: "Box 1 161130.48" or "1 161130.48"
    /\b(?:Box\s*)?1\s+\$?([\d,]+\.?\d*)/i,
    // Pattern: "Wages and tips 161130.48"
    /Wages\s+and\s+tips\s+\$?([\d,]+\.?\d*)/i,
    // Pattern: "1 Wages, tips, other compensation" followed by amount on next line
    /\b1\s+Wages[,\s]*tips[,\s]*other\s+compensation[\s\n]+\$?([\d,]+\.?\d*)/i
  ];

  for (const pattern of wagePatterns) {
    const match = ocrText.match(pattern);
    if (match && match[1]) {
      const wageString = match[1];
      console.log('🔍 [OCR] Found wage match:', wageString, 'using pattern:', pattern.source);
      
      // Parse the amount
      const cleanedAmount = wageString.replace(/[,$\s]/g, '');
      const parsedAmount = parseFloat(cleanedAmount);
      
      if (!isNaN(parsedAmount) && parsedAmount > 0) {
        console.log('✅ [OCR] Successfully parsed wages:', parsedAmount);
        return parsedAmount;
      }
    }
  }

  console.log('⚠️ [OCR] No wages found in OCR text');
  return 0;
}

// Test cases
const testCases = [
  {
    name: 'Standard W2 format',
    ocrText: 'Form W-2 Wage and Tax Statement 2023\nEmployer: Test Company Inc.\nEmployee: John Doe\n1 Wages, tips, other compensation 161130.48\n2 Federal income tax withheld 25000.00\n3 Social security wages 161130.48',
    expected: 161130.48
  },
  {
    name: 'Formatted with commas and dollar signs',
    ocrText: 'W-2 Wage and Tax Statement\n1. Wages, tips, other compensation: $161,130.48\n2. Federal income tax withheld: $25,000.00\n3. Social security wages: $161,130.48',
    expected: 161130.48
  },
  {
    name: 'Simple Box 1 format',
    ocrText: 'Box 1 161130.48\nBox 2 25000.00\nBox 3 161130.48',
    expected: 161130.48
  },
  {
    name: 'Wages and tips format',
    ocrText: 'Employee Information\nWages and tips 161130.48\nFederal tax withheld 25000.00',
    expected: 161130.48
  },
  {
    name: 'No wages found',
    ocrText: 'Some random text without wage information\nBox 2 25000.00\nBox 3 161130.48',
    expected: 0
  }
];

console.log('Running test cases...\n');

let passedTests = 0;
let totalTests = testCases.length;

testCases.forEach((testCase, index) => {
  console.log(`📋 Test ${index + 1}: ${testCase.name}`);
  console.log('OCR Text:', JSON.stringify(testCase.ocrText));
  
  const result = extractWagesFromOCR(testCase.ocrText);
  
  if (result === testCase.expected) {
    console.log(`✅ PASSED: Expected ${testCase.expected}, got ${result}`);
    passedTests++;
  } else {
    console.log(`❌ FAILED: Expected ${testCase.expected}, got ${result}`);
  }
  
  console.log('─'.repeat(60));
});

console.log(`\n🏁 Test Results: ${passedTests}/${totalTests} tests passed`);

if (passedTests === totalTests) {
  console.log('🎉 All tests passed! The wages extraction logic is working correctly.');
} else {
  console.log('⚠️ Some tests failed. Please review the extraction logic.');
}
