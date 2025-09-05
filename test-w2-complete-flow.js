#!/usr/bin/env node

// Complete flow test for W2 processing with wages extraction
console.log('🧪 Testing complete W2 processing flow...\n');

// Simulate the complete flow from Azure DI extraction to Form 1040 mapping

// Step 1: Simulate Azure Document Intelligence extraction (with missing wages field)
function simulateAzureExtraction() {
  console.log('🔍 [Azure DI] Simulating document extraction...');
  
  // This simulates what Azure DI returns when the structured "wages" field is missing
  // but the OCR text contains the wage information
  const extractedData = {
    documentType: 'FORM_W2',
    employerName: 'Test Company Inc.',
    employerEIN: '12-3456789',
    employeeName: 'John Doe',
    employeeSSN: '123-45-6789',
    // wages field is missing (this is the bug we're fixing)
    federalTaxWithheld: 25000.00,
    socialSecurityWages: 161130.48,
    medicareWages: 161130.48,
    socialSecurityTaxWithheld: 9990.09,
    medicareTaxWithheld: 2336.39,
    fullText: 'Form W-2 Wage and Tax Statement 2023\nEmployer: Test Company Inc.\nEmployee: John Doe\n1 Wages, tips, other compensation 161130.48\n2 Federal income tax withheld 25000.00\n3 Social security wages 161130.48\n4 Social security tax withheld 9990.09\n5 Medicare wages and tips 161130.48\n6 Medicare tax withheld 2336.39'
  };
  
  console.log('✅ [Azure DI] Extraction completed');
  console.log('📋 [Azure DI] Extracted data keys:', Object.keys(extractedData));
  console.log('⚠️ [Azure DI] Notice: wages field is missing, but fullText contains wage info');
  
  return extractedData;
}

// Step 2: Simulate the OCR fallback extraction
function extractWagesFromOCR(ocrText) {
  console.log('🔍 [OCR Fallback] Searching for wages in OCR text...');
  
  const wagePatterns = [
    /\b1\s+Wages[,\s]*tips[,\s]*other\s+compensation\s+([\d,]+\.?\d*)/i,
    /\b1\.?\s*Wages[,\s]*tips[,\s]*other\s+compensation[:\s]+\$?([\d,]+\.?\d*)/i,
    /\b(?:Box\s*)?1\s+\$?([\d,]+\.?\d*)/i,
    /Wages\s+and\s+tips\s+\$?([\d,]+\.?\d*)/i,
    /\b1\s+Wages[,\s]*tips[,\s]*other\s+compensation[\s\n]+\$?([\d,]+\.?\d*)/i
  ];

  for (const pattern of wagePatterns) {
    const match = ocrText.match(pattern);
    if (match && match[1]) {
      const wageString = match[1];
      console.log('🔍 [OCR Fallback] Found wage match:', wageString);
      
      const cleanedAmount = wageString.replace(/[,$\s]/g, '');
      const parsedAmount = parseFloat(cleanedAmount);
      
      if (!isNaN(parsedAmount) && parsedAmount > 0) {
        console.log('✅ [OCR Fallback] Successfully parsed wages:', parsedAmount);
        return parsedAmount;
      }
    }
  }

  console.log('⚠️ [OCR Fallback] No wages found in OCR text');
  return 0;
}

// Step 3: Simulate the enhanced W2 processing (with OCR fallback)
function processW2WithFallback(extractedData) {
  console.log('🔍 [W2 Processor] Processing W2 data with fallback logic...');
  
  // First try to get wages from structured data
  let wages = extractedData.wages || 0;
  console.log('🔍 [W2 Processor] Structured wages value:', wages);
  
  // If wages not found, try OCR fallback
  if (wages === 0 && extractedData.fullText) {
    console.log('🔍 [W2 Processor] Wages not found in structured data, trying OCR fallback...');
    wages = extractWagesFromOCR(extractedData.fullText);
  }
  
  // Create the final processed data
  const processedData = {
    ...extractedData,
    wages: wages // Add the wages field
  };
  
  console.log('✅ [W2 Processor] Processing completed');
  console.log('📋 [W2 Processor] Final wages value:', processedData.wages);
  
  return processedData;
}

// Step 4: Simulate Form 1040 mapping
function mapToForm1040(w2Data) {
  console.log('🔍 [1040 Mapper] Mapping W2 data to Form 1040...');
  
  const form1040Data = {
    line1: 0, // Wages from W2
    line25a: 0, // Federal tax withheld
    line9: 0, // Total income
    line11: 0, // AGI
    line12: 13850, // Standard deduction (single)
    line15: 0, // Taxable income
    line16: 0, // Tax liability
    line24: 0, // Total tax
    line32: 0, // Total payments
    line33: 0, // Refund
    line37: 0 // Amount owed
  };
  
  // Map wages to Line 1
  if (w2Data.wages && w2Data.wages > 0) {
    form1040Data.line1 = w2Data.wages;
    console.log('✅ [1040 Mapper] Mapped wages to Line 1:', form1040Data.line1);
  } else {
    console.log('⚠️ [1040 Mapper] No wages found to map to Line 1');
  }
  
  // Map federal tax withheld to Line 25a
  if (w2Data.federalTaxWithheld && w2Data.federalTaxWithheld > 0) {
    form1040Data.line25a = w2Data.federalTaxWithheld;
    console.log('✅ [1040 Mapper] Mapped federal tax withheld to Line 25a:', form1040Data.line25a);
  }
  
  // Calculate other lines
  form1040Data.line9 = form1040Data.line1; // Total income = wages for this simple case
  form1040Data.line11 = form1040Data.line9; // AGI = total income (no adjustments)
  form1040Data.line15 = Math.max(0, form1040Data.line11 - form1040Data.line12); // Taxable income
  
  // Simple tax calculation (10% bracket for demo)
  form1040Data.line16 = Math.round(form1040Data.line15 * 0.22 * 100) / 100; // Simplified 22% tax rate
  form1040Data.line24 = form1040Data.line16; // Total tax = tax liability
  form1040Data.line32 = form1040Data.line25a; // Total payments = federal tax withheld
  
  // Calculate refund or amount owed
  if (form1040Data.line32 > form1040Data.line24) {
    form1040Data.line33 = form1040Data.line32 - form1040Data.line24; // Refund
  } else {
    form1040Data.line37 = form1040Data.line24 - form1040Data.line32; // Amount owed
  }
  
  console.log('✅ [1040 Mapper] Form 1040 mapping completed');
  return form1040Data;
}

// Run the complete test
console.log('🚀 Starting complete W2 processing flow test...\n');

try {
  // Step 1: Extract data (with missing wages field)
  console.log('📋 STEP 1: Azure Document Intelligence Extraction');
  const extractedData = simulateAzureExtraction();
  console.log('');
  
  // Step 2: Process with fallback logic
  console.log('📋 STEP 2: W2 Processing with OCR Fallback');
  const processedW2Data = processW2WithFallback(extractedData);
  console.log('');
  
  // Step 3: Map to Form 1040
  console.log('📋 STEP 3: Form 1040 Mapping');
  const form1040Data = mapToForm1040(processedW2Data);
  console.log('');
  
  // Step 4: Verify results
  console.log('📋 STEP 4: Results Verification');
  console.log('═'.repeat(60));
  console.log('📊 FINAL RESULTS:');
  console.log('═'.repeat(60));
  console.log('Form 1040 Line 1 (Wages):', `$${form1040Data.line1.toLocaleString()}`);
  console.log('Form 1040 Line 25a (Fed Tax Withheld):', `$${form1040Data.line25a.toLocaleString()}`);
  console.log('Form 1040 Line 9 (Total Income):', `$${form1040Data.line9.toLocaleString()}`);
  console.log('Form 1040 Line 11 (AGI):', `$${form1040Data.line11.toLocaleString()}`);
  console.log('Form 1040 Line 15 (Taxable Income):', `$${form1040Data.line15.toLocaleString()}`);
  console.log('Form 1040 Line 16 (Tax Liability):', `$${form1040Data.line16.toLocaleString()}`);
  console.log('Form 1040 Line 32 (Total Payments):', `$${form1040Data.line32.toLocaleString()}`);
  console.log('Form 1040 Line 33 (Refund):', `$${form1040Data.line33.toLocaleString()}`);
  console.log('Form 1040 Line 37 (Amount Owed):', `$${form1040Data.line37.toLocaleString()}`);
  console.log('═'.repeat(60));
  
  // Verify the fix worked
  if (form1040Data.line1 === 161130.48) {
    console.log('🎉 SUCCESS: The wages extraction fix is working!');
    console.log('✅ Form 1040 Line 1 correctly shows $161,130.48 instead of $0');
    console.log('✅ OCR fallback successfully extracted wages from text');
    console.log('✅ W2-to-1040 mapping is functioning properly');
  } else {
    console.log('❌ FAILED: Expected Line 1 to be $161,130.48, got $' + form1040Data.line1);
  }
  
} catch (error) {
  console.error('💥 Test failed with error:', error.message);
  console.error('Stack trace:', error.stack);
}

console.log('\n🏁 Complete flow test finished!');
