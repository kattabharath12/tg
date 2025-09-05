#!/usr/bin/env node

// Test script to verify W2 wages extraction and mapping
const { W2ToForm1040Mapper } = require('./tax_filing_app/app/lib/w2-to-1040-mapping.ts');

console.log('🧪 Testing W2 wages extraction and mapping...\n');

// Test case 1: Simulated Azure DI response with missing wages field but OCR text
const testW2Data1 = {
  extractedData: {
    documentType: 'FORM_W2',
    employerName: 'Test Company Inc.',
    employerEIN: '12-3456789',
    employeeName: 'John Doe',
    employeeSSN: '123-45-6789',
    // wages field is missing (this is the bug)
    federalTaxWithheld: 25000.00,
    socialSecurityWages: 161130.48,
    medicareWages: 161130.48,
    socialSecurityTaxWithheld: 9990.09,
    medicareTaxWithheld: 2336.39,
    fullText: 'Form W-2 Wage and Tax Statement 2023\nEmployer: Test Company Inc.\nEmployee: John Doe\n1 Wages, tips, other compensation 161130.48\n2 Federal income tax withheld 25000.00\n3 Social security wages 161130.48\n4 Social security tax withheld 9990.09\n5 Medicare wages and tips 161130.48\n6 Medicare tax withheld 2336.39'
  }
};

console.log('📋 Test Case 1: Missing wages field with OCR fallback');
console.log('Input data structure:', JSON.stringify(testW2Data1, null, 2));

try {
  const form1040Data1 = W2ToForm1040Mapper.mapW2ToForm1040(testW2Data1);
  console.log('\n✅ Mapping Result:');
  console.log('Line 1 (Wages):', form1040Data1.line1);
  console.log('Line 25a (Federal Tax Withheld):', form1040Data1.line25a);
  
  if (form1040Data1.line1 === 161130.48) {
    console.log('🎉 SUCCESS: Wages correctly extracted and mapped to Line 1!');
  } else {
    console.log('❌ FAILED: Expected 161130.48, got', form1040Data1.line1);
  }
} catch (error) {
  console.error('❌ Test failed with error:', error.message);
}

console.log('\n' + '='.repeat(60) + '\n');

// Test case 2: Simulated Azure DI response with wages field present
const testW2Data2 = {
  extractedData: {
    documentType: 'FORM_W2',
    employerName: 'Test Company Inc.',
    employerEIN: '12-3456789',
    employeeName: 'John Doe',
    employeeSSN: '123-45-6789',
    wages: 161130.48, // wages field is present
    federalTaxWithheld: 25000.00,
    socialSecurityWages: 161130.48,
    medicareWages: 161130.48,
    socialSecurityTaxWithheld: 9990.09,
    medicareTaxWithheld: 2336.39,
    fullText: 'Form W-2 Wage and Tax Statement 2023\nEmployer: Test Company Inc.\nEmployee: John Doe\n1 Wages, tips, other compensation 161130.48\n2 Federal income tax withheld 25000.00\n3 Social security wages 161130.48\n4 Social security tax withheld 9990.09\n5 Medicare wages and tips 161130.48\n6 Medicare tax withheld 2336.39'
  }
};

console.log('📋 Test Case 2: Wages field present (normal case)');
console.log('Input data structure:', JSON.stringify(testW2Data2, null, 2));

try {
  const form1040Data2 = W2ToForm1040Mapper.mapW2ToForm1040(testW2Data2);
  console.log('\n✅ Mapping Result:');
  console.log('Line 1 (Wages):', form1040Data2.line1);
  console.log('Line 25a (Federal Tax Withheld):', form1040Data2.line25a);
  
  if (form1040Data2.line1 === 161130.48) {
    console.log('🎉 SUCCESS: Wages correctly mapped to Line 1!');
  } else {
    console.log('❌ FAILED: Expected 161130.48, got', form1040Data2.line1);
  }
} catch (error) {
  console.error('❌ Test failed with error:', error.message);
}

console.log('\n' + '='.repeat(60) + '\n');

// Test case 3: Different OCR text format
const testW2Data3 = {
  extractedData: {
    documentType: 'FORM_W2',
    employerName: 'Test Company Inc.',
    employerEIN: '12-3456789',
    employeeName: 'John Doe',
    employeeSSN: '123-45-6789',
    // wages field is missing
    federalTaxWithheld: 25000.00,
    socialSecurityWages: 161130.48,
    fullText: 'W-2 Wage and Tax Statement\n1. Wages, tips, other compensation: $161,130.48\n2. Federal income tax withheld: $25,000.00\n3. Social security wages: $161,130.48'
  }
};

console.log('📋 Test Case 3: Different OCR format with commas and dollar signs');
console.log('Input data structure:', JSON.stringify(testW2Data3, null, 2));

try {
  const form1040Data3 = W2ToForm1040Mapper.mapW2ToForm1040(testW2Data3);
  console.log('\n✅ Mapping Result:');
  console.log('Line 1 (Wages):', form1040Data3.line1);
  console.log('Line 25a (Federal Tax Withheld):', form1040Data3.line25a);
  
  if (form1040Data3.line1 === 161130.48) {
    console.log('🎉 SUCCESS: Wages correctly extracted from formatted OCR text!');
  } else {
    console.log('❌ FAILED: Expected 161130.48, got', form1040Data3.line1);
  }
} catch (error) {
  console.error('❌ Test failed with error:', error.message);
}

console.log('\n🏁 Test completed!');
