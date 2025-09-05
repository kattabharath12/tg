
import { NextRequest, NextResponse } from 'next/server';
import { getServerSession } from 'next-auth/next';
import { authOptions } from '@/lib/auth';
import { prisma } from '@/lib/prisma';
import { AzureDocumentIntelligenceService, ExtractedFieldData } from '@/lib/azure-document-intelligence-service';
import { DocumentType } from '@prisma/client';

export async function POST(
  request: NextRequest,
  { params }: { params: { id: string } }
) {
  try {
    const session = await getServerSession(authOptions);
    if (!session?.user?.id) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    const documentId = params.id;

    // Get the document from the database
    const document = await prisma.document.findUnique({
      where: {
        id: documentId,
      },
    });

    if (!document) {
      return NextResponse.json({ error: 'Document not found' }, { status: 404 });
    }

    if (document.processingStatus === 'COMPLETED') {
      return NextResponse.json({ error: 'Document already processed' }, { status: 400 });
    }

    try {
      // Initialize Azure Document Intelligence service
      const azureService = new AzureDocumentIntelligenceService({
        endpoint: process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT!,
        apiKey: process.env.AZURE_DOCUMENT_INTELLIGENCE_API_KEY!,
      });

      console.log("🔍 [PROCESS] Starting document processing for:", document.fileName);
      console.log("🔍 [PROCESS] Document type:", document.documentType);

      // Extract data from the document
      const extractedTaxData = await azureService.extractDataFromDocument(
        document.filePath,
        document.documentType
      );

      console.log("✅ [PROCESS] Extraction completed. Extracted data:", JSON.stringify(extractedTaxData, null, 2));

      // Handle document type correction if needed
      let finalDocumentType = document.documentType;
      if (extractedTaxData.correctedDocumentType) {
        console.log(`🔄 [PROCESS] Document type corrected: ${document.documentType} → ${extractedTaxData.correctedDocumentType}`);
        finalDocumentType = extractedTaxData.correctedDocumentType;
        
        // Update document type in database
        await prisma.document.update({
          where: { id: documentId },
          data: { documentType: finalDocumentType },
        });
      }

      // Process the extracted data based on document type
      let processedEntries: any[] = [];
      
      switch (finalDocumentType) {
        case 'W2':
        case 'W2_CORRECTED':
          console.log("🔍 [PROCESS] Processing W2 document...")
          processedEntries = await processW2Document(extractedTaxData.extractedData || extractedTaxData);
          break;
        case 'FORM_1099_INT':
          console.log("🔍 [PROCESS] Processing 1099-INT document...")
          processedEntries = await process1099IntDocument(extractedTaxData.extractedData || extractedTaxData);
          break;
        case 'FORM_1099_DIV':
          console.log("🔍 [PROCESS] Processing 1099-DIV document...")
          processedEntries = await process1099DivDocument(extractedTaxData.extractedData || extractedTaxData);
          break;
        case 'FORM_1099_MISC':
          console.log("🔍 [PROCESS] Processing 1099-MISC document...")
          processedEntries = await process1099MiscDocument(extractedTaxData.extractedData || extractedTaxData);
          break;
        case 'FORM_1099_NEC':
          console.log("🔍 [PROCESS] Processing 1099-NEC document...")
          processedEntries = await process1099NecDocument(extractedTaxData.extractedData || extractedTaxData);
          break;
        default:
          console.log("⚠️ [PROCESS] Unknown document type, creating generic entries...")
          processedEntries = await processGenericDocument(extractedTaxData.extractedData || extractedTaxData);
          break;
      }

      console.log("✅ [PROCESS] Processing completed. Entries created:", processedEntries.length);

      // Update the document with extracted data and mark as completed
      const updatedDocument = await prisma.document.update({
        where: { id: documentId },
        data: {
          extractedData: extractedTaxData,
          processingStatus: 'COMPLETED',
          processedAt: new Date(),
        },
      });

      return NextResponse.json({
        success: true,
        document: updatedDocument,
        extractedData: extractedTaxData,
        processedEntries: processedEntries,
      });

    } catch (processingError: any) {
      console.error('❌ [PROCESS] Error processing document:', processingError);
      
      // Update processing status to FAILED
      await prisma.document.update({
        where: { id: documentId },
        data: { 
          processingStatus: 'FAILED',
          processingError: processingError.message 
        },
      });

      return NextResponse.json(
        { error: `Processing failed: ${processingError.message}` },
        { status: 500 }
      );
    }

  } catch (error: any) {
    console.error('❌ [PROCESS] Unexpected error:', error);
    return NextResponse.json(
      { error: 'Internal server error' },
      { status: 500 }
    );
  }
}

// Enhanced 1099-INT document processing function
async function process1099IntDocument(extractedData: ExtractedFieldData): Promise<any[]> {
  const entries = [];
  
  // Enhanced field mappings for all 1099-INT boxes and fields
  const fieldMappings = {
    // Payer and recipient information
    'payerName': 'Payer Name',
    'payerTIN': 'Payer TIN',
    'payerAddress': 'Payer Address',
    'recipientName': 'Recipient Name',
    'recipientTIN': 'Recipient TIN',
    'recipientAddress': 'Recipient Address',
    'accountNumber': 'Account Number',
    
    // Box 1-17 mappings
    'interestIncome': 'Box 1 - Interest Income',
    'earlyWithdrawalPenalty': 'Box 2 - Early Withdrawal Penalty',
    'interestOnUSavingsBonds': 'Box 3 - Interest on US Savings Bonds and Treasury Obligations',
    'federalTaxWithheld': 'Box 4 - Federal Income Tax Withheld',
    'investmentExpenses': 'Box 5 - Investment Expenses',
    'foreignTaxPaid': 'Box 6 - Foreign Tax Paid',
    'taxExemptInterest': 'Box 8 - Tax-Exempt Interest',
    'specifiedPrivateActivityBondInterest': 'Box 9 - Specified Private Activity Bond Interest',
    'marketDiscount': 'Box 10 - Market Discount',
    'bondPremium': 'Box 11 - Bond Premium',
    'bondPremiumOnTreasuryObligations': 'Box 12 - Bond Premium on Treasury Obligations',
    'bondPremiumOnTaxExemptBond': 'Box 13 - Bond Premium on Tax-Exempt Bond',
    'taxExemptAndTaxCreditBondCUSIPNo': 'Box 14 - Tax-Exempt and Tax Credit Bond CUSIP No.',
    'state': 'Box 15 - State',
    'stateIdentificationNo': 'Box 16 - State Identification No.',
    'stateTaxWithheld': 'Box 17 - State Tax Withheld'
  };
  
  for (const [fieldKey, displayName] of Object.entries(fieldMappings)) {
    if (extractedData[fieldKey] !== undefined && extractedData[fieldKey] !== null && extractedData[fieldKey] !== '') {
      entries.push({
        fieldName: displayName,
        fieldValue: String(extractedData[fieldKey]),
        confidence: 0.95
      });
    }
  }
  
  return entries;
}

async function process1099DivDocument(extractedData: ExtractedFieldData): Promise<any[]> {
  const entries = [];
  
  const fieldMappings = {
    'payerName': 'Payer Name',
    'payerTIN': 'Payer TIN',
    'payerAddress': 'Payer Address',
    'recipientName': 'Recipient Name',
    'recipientTIN': 'Recipient TIN',
    'recipientAddress': 'Recipient Address',
    'ordinaryDividends': 'Box 1a - Ordinary Dividends',
    'qualifiedDividends': 'Box 1b - Qualified Dividends',
    'totalCapitalGain': 'Box 2a - Total Capital Gain Distributions',
    'nondividendDistributions': 'Box 3 - Nondividend Distributions',
    'federalTaxWithheld': 'Box 4 - Federal Income Tax Withheld',
    'section199ADividends': 'Box 5 - Section 199A Dividends'
  };
  
  for (const [fieldKey, displayName] of Object.entries(fieldMappings)) {
    if (extractedData[fieldKey] !== undefined && extractedData[fieldKey] !== null && extractedData[fieldKey] !== '') {
      entries.push({
        fieldName: displayName,
        fieldValue: String(extractedData[fieldKey]),
        confidence: 0.95
      });
    }
  }
  
  return entries;
}

async function process1099MiscDocument(extractedData: ExtractedFieldData): Promise<any[]> {
  const entries = [];
  
  const fieldMappings = {
    'payerName': 'Payer Name',
    'payerTIN': 'Payer TIN',
    'payerAddress': 'Payer Address',
    'recipientName': 'Recipient Name',
    'recipientTIN': 'Recipient TIN',
    'recipientAddress': 'Recipient Address',
    'accountNumber': 'Account Number',
    'rents': 'Box 1 - Rents',
    'royalties': 'Box 2 - Royalties',
    'otherIncome': 'Box 3 - Other Income',
    'federalTaxWithheld': 'Box 4 - Federal Income Tax Withheld',
    'fishingBoatProceeds': 'Box 5 - Fishing Boat Proceeds',
    'medicalHealthPayments': 'Box 6 - Medical and Health Care Payments',
    'nonemployeeCompensation': 'Box 7 - Nonemployee Compensation',
    'substitutePayments': 'Box 8 - Substitute Payments',
    'cropInsuranceProceeds': 'Box 9 - Crop Insurance Proceeds',
    'grossProceedsAttorney': 'Box 10 - Gross Proceeds Paid to an Attorney',
    'fishPurchases': 'Box 11 - Fish Purchased for Resale',
    'section409ADeferrals': 'Box 12 - Section 409A Deferrals',
    'excessGoldenParachutePayments': 'Box 13 - Excess Golden Parachute Payments',
    'nonqualifiedDeferredCompensation': 'Box 14 - Nonqualified Deferred Compensation',
    'section409AIncome': 'Box 15a - Section 409A Income',
    'stateTaxWithheld': 'Box 16 - State Tax Withheld',
    'statePayerNumber': 'Box 17 - State/Payer\'s State No.',
    'stateIncome': 'Box 18 - State Income'
  };
  
  for (const [fieldKey, displayName] of Object.entries(fieldMappings)) {
    if (extractedData[fieldKey] !== undefined && extractedData[fieldKey] !== null && extractedData[fieldKey] !== '') {
      entries.push({
        fieldName: displayName,
        fieldValue: String(extractedData[fieldKey]),
        confidence: 0.95
      });
    }
  }
  
  return entries;
}

async function process1099NecDocument(extractedData: ExtractedFieldData): Promise<any[]> {
  const entries = [];
  
  const fieldMappings = {
    'payerName': 'Payer Name',
    'payerTIN': 'Payer TIN',
    'payerAddress': 'Payer Address',
    'recipientName': 'Recipient Name',
    'recipientTIN': 'Recipient TIN',
    'recipientAddress': 'Recipient Address',
    'nonemployeeCompensation': 'Box 1 - Nonemployee Compensation',
    'federalTaxWithheld': 'Box 4 - Federal Income Tax Withheld',
    'stateTaxWithheld': 'Box 5 - State Tax Withheld',
    'statePayerNumber': 'Box 6 - State/Payer\'s State No.',
    'stateIncome': 'Box 7 - State Income'
  };
  
  for (const [fieldKey, displayName] of Object.entries(fieldMappings)) {
    if (extractedData[fieldKey] !== undefined && extractedData[fieldKey] !== null && extractedData[fieldKey] !== '') {
      entries.push({
        fieldName: displayName,
        fieldValue: String(extractedData[fieldKey]),
        confidence: 0.95
      });
    }
  }
  
  return entries;
}

async function processW2Document(extractedData: ExtractedFieldData): Promise<any[]> {
  const entries = [];
  
  const fieldMappings = {
    'employerName': 'Employer Name',
    'employerTIN': 'Employer TIN',
    'employerAddress': 'Employer Address',
    'employeeName': 'Employee Name',
    'employeeSSN': 'Employee SSN',
    'employeeAddress': 'Employee Address',
    'wages': 'Box 1 - Wages, Tips, Other Compensation',
    'federalTaxWithheld': 'Box 2 - Federal Income Tax Withheld',
    'socialSecurityWages': 'Box 3 - Social Security Wages',
    'socialSecurityTaxWithheld': 'Box 4 - Social Security Tax Withheld',
    'medicareWages': 'Box 5 - Medicare Wages and Tips',
    'medicareTaxWithheld': 'Box 6 - Medicare Tax Withheld'
  };
  
  for (const [fieldKey, displayName] of Object.entries(fieldMappings)) {
    if (extractedData[fieldKey] !== undefined && extractedData[fieldKey] !== null && extractedData[fieldKey] !== '') {
      entries.push({
        fieldName: displayName,
        fieldValue: String(extractedData[fieldKey]),
        confidence: 0.95
      });
    }
  }
  
  return entries;
}

async function processGenericDocument(extractedData: ExtractedFieldData): Promise<any[]> {
  const entries = [];
  
  // Process all available fields generically
  for (const [fieldKey, fieldValue] of Object.entries(extractedData)) {
    if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '' && fieldKey !== 'fullText') {
      entries.push({
        fieldName: fieldKey,
        fieldValue: String(fieldValue),
        confidence: 0.8
      });
    }
  }
  
  return entries;
}
