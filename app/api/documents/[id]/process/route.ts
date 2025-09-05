


import { NextResponse } from "next/server"
import { getServerSession } from "next-auth"
import { authOptions } from "@/lib/auth"
import { prisma } from "@/lib/db"
import { DocumentType, ProcessingStatus, EntryType } from "@prisma/client"
import { getAzureDocumentIntelligenceService, type ExtractedFieldData } from "@/lib/azure-document-intelligence-service"
import { DuplicateDetectionService, type DuplicateDetectionResult } from "@/lib/duplicate-detection"

export const dynamic = "force-dynamic"

export async function POST(
  request: Request,
  { params }: { params: { id: string } }
) {
  console.log("🔍 [PROCESS] Starting document processing for ID:", params.id)
  
  try {
    console.log("🔍 [PROCESS] Step 1: Getting server session...")
    const session = await getServerSession(authOptions)
    
    if (!session?.user?.id) {
      console.log("❌ [PROCESS] No valid session found")
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 })
    }
    
    console.log("✅ [PROCESS] Session validated for user:", session.user.id)

    console.log("🔍 [PROCESS] Step 2: Fetching document from database...")
    const document = await prisma.document.findUnique({
      where: { id: params.id },
      include: {
        taxReturn: {
          select: {
            userId: true
          }
        }
      }
    })

    if (!document) {
      console.log("❌ [PROCESS] Document not found")
      return NextResponse.json({ error: "Document not found" }, { status: 404 })
    }

    if (document.taxReturn.userId !== session.user.id) {
      console.log("❌ [PROCESS] Document does not belong to user")
      return NextResponse.json({ error: "Unauthorized" }, { status: 403 })
    }

    console.log("✅ [PROCESS] Document found:", {
      id: document.id,
      fileName: document.fileName,
      documentType: document.documentType,
      processingStatus: document.processingStatus
    })

    console.log("🔍 [PROCESS] Step 3: Checking if document is already processed...")
    if (document.processingStatus === 'COMPLETED') {
      console.log("ℹ️ [PROCESS] Document already processed, returning existing data")
      
      const existingEntries = await prisma.documentExtractedEntry.findMany({
        where: { documentId: document.id }
      })
      
      return NextResponse.json({
        success: true,
        message: "Document already processed",
        data: {
          documentId: document.id,
          documentType: document.documentType,
          extractedData: document.extractedData, // Include raw extracted data for frontend processing
          extractedEntries: existingEntries,
          ocrText: document.ocrText,
          confidence: 0.95 // Default confidence for already processed documents
        }
      })
    }

    console.log("🔍 [PROCESS] Step 4: Updating processing status to PROCESSING...")
    await prisma.document.update({
      where: { id: params.id },
      data: { processingStatus: ProcessingStatus.PROCESSING }
    })
    console.log("✅ [PROCESS] Processing status updated")

    console.log("🔍 [PROCESS] Step 5: Checking Azure Document Intelligence configuration...")
    const azureEndpoint = process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT
    const azureApiKey = process.env.AZURE_DOCUMENT_INTELLIGENCE_API_KEY

    if (!azureEndpoint || !azureApiKey) {
      console.log("❌ [PROCESS] Azure Document Intelligence not configured")
      await prisma.document.update({
        where: { id: params.id },
        data: { processingStatus: ProcessingStatus.FAILED }
      })
      return NextResponse.json({ 
        error: "Azure Document Intelligence service not configured" 
      }, { status: 500 })
    }
    
    console.log("✅ [PROCESS] Azure Document Intelligence configured")

    console.log("🔍 [PROCESS] Step 6: Processing with Azure Document Intelligence...")
    let extractedTaxData: any;
    let finalDocumentType = document.documentType;
    
    try {
      const azureService = getAzureDocumentIntelligenceService();
      const extractedData = await azureService.extractDataFromDocument(document.filePath, document.documentType);
      
      // Check if document type was corrected based on OCR analysis
      if (extractedData.correctedDocumentType) {
        console.log(`🔄 [PROCESS] Document type corrected: ${document.documentType} → ${extractedData.correctedDocumentType}`);
        
        // Convert string to DocumentType enum with validation
        const correctedType = extractedData.correctedDocumentType as string;
        if (Object.values(DocumentType).includes(correctedType as DocumentType)) {
          finalDocumentType = correctedType as DocumentType;
          
          // Update the document type in the database
          await prisma.document.update({
            where: { id: params.id },
            data: { documentType: finalDocumentType }
          });
          console.log("✅ [PROCESS] Document type updated in database");
        } else {
          console.log(`⚠️ [PROCESS] Invalid document type returned: ${correctedType}, keeping original: ${document.documentType}`);
        }
      }
      
      extractedTaxData = {
        documentType: finalDocumentType,
        ocrText: extractedData.fullText || '',
        extractedData: extractedData,
        confidence: 0.95 // Azure typically has high confidence
      };
      
      console.log("✅ [PROCESS] Final document type:", finalDocumentType)
      console.log("✅ [PROCESS] Azure Document Intelligence processing completed")
    } catch (azureError: any) {
      console.error("❌ [PROCESS] Azure Document Intelligence processing failed:", azureError)
      await prisma.document.update({
        where: { id: params.id },
        data: { processingStatus: ProcessingStatus.FAILED }
      })
      return NextResponse.json({ 
        error: `Document processing failed: ${azureError?.message || 'Unknown error'}` 
      }, { status: 500 })
    }

    console.log("🔍 [PROCESS] Step 7: Checking for duplicate documents...")
    let duplicateInfo: DuplicateDetectionResult | null = null;
    
    try {
      duplicateInfo = await DuplicateDetectionService.checkForDuplicates({
        documentType: finalDocumentType,
        extractedData: extractedTaxData.extractedData,
        taxReturnId: document.taxReturnId
      });
      
      if (duplicateInfo.isDuplicate) {
        console.log("⚠️ [PROCESS] Duplicate document detected:", duplicateInfo);
      } else {
        console.log("✅ [PROCESS] No duplicates found");
      }
    } catch (duplicateError: any) {
      console.error("❌ [PROCESS] Duplicate detection failed:", duplicateError);
      // Continue processing even if duplicate detection fails
      duplicateInfo = { 
        isDuplicate: false, 
        confidence: 0, 
        matchingDocuments: [],
        matchCriteria: {
          documentType: false,
          employerInfo: false,
          recipientInfo: false,
          amountSimilarity: false,
          nameSimilarity: false
        }
      };
    }

    console.log("🔍 [PROCESS] Step 8: Extracting structured data...")
    const extractedEntries: any[] = [];
    
    try {
      // Process different document types
      switch (finalDocumentType) {
        case 'W2':
          console.log("🔍 [PROCESS] Processing W2 document...")
          const w2Entries = await processW2Document(extractedTaxData.extractedData);
          extractedEntries.push(...w2Entries);
          break;
          
        case 'FORM_1099_INT':
          console.log("🔍 [PROCESS] Processing 1099-INT document...")
          const intEntries = await process1099IntDocument(extractedTaxData.extractedData);
          extractedEntries.push(...intEntries);
          break;
          
        case 'FORM_1099_DIV':
          console.log("🔍 [PROCESS] Processing 1099-DIV document...")
          const divEntries = await process1099DivDocument(extractedTaxData.extractedData);
          extractedEntries.push(...divEntries);
          break;
          
        case 'FORM_1099_MISC':
          console.log("🔍 [PROCESS] Processing 1099-MISC document...")
          const miscEntries = await process1099MiscDocument(extractedTaxData.extractedData);
          extractedEntries.push(...miscEntries);
          break;
          
        case 'FORM_1099_NEC':
          console.log("🔍 [PROCESS] Processing 1099-NEC document...")
          const necEntries = await process1099NecDocument(extractedTaxData.extractedData);
          extractedEntries.push(...necEntries);
          break;
          
        default:
          console.log("🔍 [PROCESS] Processing generic document...")
          const genericEntries = await processGenericDocument(extractedTaxData.extractedData);
          extractedEntries.push(...genericEntries);
          break;
      }
      
      console.log(`✅ [PROCESS] Extracted ${extractedEntries.length} entries`)
    } catch (extractionError: any) {
      console.error("❌ [PROCESS] Data extraction failed:", extractionError);
      await prisma.document.update({
        where: { id: params.id },
        data: { processingStatus: ProcessingStatus.FAILED }
      })
      return NextResponse.json({ 
        error: `Data extraction failed: ${extractionError?.message || 'Unknown error'}` 
      }, { status: 500 })
    }

    console.log("🔍 [PROCESS] Step 9: Saving extracted entries to database...")
    try {
      // Save extracted entries to database
      const savedEntries = await Promise.all(
        extractedEntries.map(entry => 
          prisma.documentExtractedEntry.create({
            data: {
              documentId: document.id,
              entryType: finalDocumentType === 'W2' ? EntryType.INCOME : EntryType.INCOME, // Default to INCOME for now
              extractedData: {
                fieldName: entry.fieldName,
                fieldValue: entry.fieldValue,
                confidence: entry.confidence || 0.95,
                boundingBox: entry.boundingBox || null
              }
            }
          })
        )
      );
      
      console.log(`✅ [PROCESS] Saved ${savedEntries.length} entries to database`)
    } catch (saveError: any) {
      console.error("❌ [PROCESS] Failed to save entries:", saveError);
      await prisma.document.update({
        where: { id: params.id },
        data: { processingStatus: ProcessingStatus.FAILED }
      })
      return NextResponse.json({ 
        error: `Failed to save extracted data: ${saveError?.message || 'Unknown error'}` 
      }, { status: 500 })
    }

    console.log("🔍 [PROCESS] Step 10: Updating document status to COMPLETED and saving extractedData...")
    // CRITICAL FIX: Save the extractedData to the Document model for Form 1040 retrieval
    await prisma.document.update({
      where: { id: params.id },
      data: { 
        processingStatus: ProcessingStatus.COMPLETED,
        extractedData: extractedTaxData.extractedData, // This is the key fix!
        ocrText: extractedTaxData.ocrText
      }
    })

    console.log("✅ [PROCESS] Document processing completed successfully")
    
    return NextResponse.json({
      success: true,
      message: "Document processed successfully",
      data: {
        documentId: document.id,
        documentType: finalDocumentType,
        extractedData: extractedTaxData.extractedData, // Include raw extracted data for frontend processing
        extractedEntries: extractedEntries,
        duplicateInfo: duplicateInfo,
        ocrText: extractedTaxData.ocrText,
        confidence: extractedTaxData.confidence
      }
    })

  } catch (error: any) {
    console.error("❌ [PROCESS] Unexpected error:", error)
    
    // Update document status to failed if possible
    try {
      await prisma.document.update({
        where: { id: params.id },
        data: { processingStatus: ProcessingStatus.FAILED }
      })
    } catch (updateError) {
      console.error("❌ [PROCESS] Failed to update document status:", updateError)
    }
    
    return NextResponse.json({ 
      error: `Processing failed: ${error?.message || 'Unknown error'}` 
    }, { status: 500 })
  }
}

// Helper methods for processing different document types
async function processW2Document(extractedData: ExtractedFieldData): Promise<any[]> {
    const entries = [];
    
    // Map W2 fields to database entries
    const w2FieldMappings = {
      'employeeName': 'Employee Name',
      'employeeSSN': 'Employee SSN',
      'employeeAddress': 'Employee Address',
      'employerName': 'Employer Name',
      'employerEIN': 'Employer EIN',
      'employerAddress': 'Employer Address',
      'wages': 'Box 1 - Wages',
      'federalTaxWithheld': 'Box 2 - Federal Tax Withheld',
      'socialSecurityWages': 'Box 3 - Social Security Wages',
      'socialSecurityTaxWithheld': 'Box 4 - Social Security Tax Withheld',
      'medicareWages': 'Box 5 - Medicare Wages',
      'medicareTaxWithheld': 'Box 6 - Medicare Tax Withheld',
      'socialSecurityTips': 'Box 7 - Social Security Tips',
      'allocatedTips': 'Box 8 - Allocated Tips',
      'stateWages': 'Box 15 - State Wages',
      'stateTaxWithheld': 'Box 17 - State Tax Withheld',
      'localWages': 'Box 18 - Local Wages',
      'localTaxWithheld': 'Box 19 - Local Tax Withheld'
    };
    
    for (const [fieldKey, displayName] of Object.entries(w2FieldMappings)) {
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

async function process1099IntDocument(extractedData: ExtractedFieldData): Promise<any[]> {
    const entries = [];
    
    const fieldMappings = {
      'payerName': 'Payer Name',
      'payerTIN': 'Payer TIN',
      'payerAddress': 'Payer Address',
      'recipientName': 'Recipient Name',
      'recipientTIN': 'Recipient TIN',
      'recipientAddress': 'Recipient Address',
      'interestIncome': 'Box 1 - Interest Income',
      'earlyWithdrawalPenalty': 'Box 2 - Early Withdrawal Penalty',
      'interestOnUSavingsBonds': 'Box 3 - Interest on US Savings Bonds',
      'federalTaxWithheld': 'Box 4 - Federal Tax Withheld',
      'investmentExpenses': 'Box 5 - Investment Expenses',
      'foreignTaxPaid': 'Box 6 - Foreign Tax Paid',
      'taxExemptInterest': 'Box 8 - Tax-Exempt Interest'
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
      'federalTaxWithheld': 'Box 4 - Federal Tax Withheld',
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
    
    // Comprehensive field mappings for all 1099-MISC boxes and information
    const fieldMappings = {
      // Payer and recipient information
      'payerName': 'Payer Name',
      'payerTIN': 'Payer TIN',
      'payerAddress': 'Payer Address',
      'recipientName': 'Recipient Name',
      'recipientTIN': 'Recipient TIN',
      'recipientAddress': 'Recipient Address',
      'accountNumber': 'Account Number',
      
      // Box 1-18 mappings
      'rents': 'Box 1 - Rents',
      'royalties': 'Box 2 - Royalties',
      'otherIncome': 'Box 3 - Other Income',
      'federalTaxWithheld': 'Box 4 - Federal Income Tax Withheld',
      'fishingBoatProceeds': 'Box 5 - Fishing Boat Proceeds',
      'medicalHealthPayments': 'Box 6 - Medical and Health Care Payments',
      'nonemployeeCompensation': 'Box 7 - Nonemployee Compensation',
      'substitutePayments': 'Box 8 - Substitute Payments in Lieu of Dividends or Interest',
      'cropInsuranceProceeds': 'Box 9 - Crop Insurance Proceeds',
      'attorneyProceeds': 'Box 10 - Gross Proceeds Paid to an Attorney',
      'fishPurchases': 'Box 11 - Fish Purchased for Resale',
      'section409ADeferrals': 'Box 12 - Section 409A Deferrals',
      'excessGoldenParachutePayments': 'Box 13 - Excess Golden Parachute Payments',
      'nonqualifiedDeferredCompensation': 'Box 14 - Nonqualified Deferred Compensation',
      'section409AIncome': 'Box 15a - Section 409A Income',
      'stateTaxWithheld': 'Box 16 - State Tax Withheld',
      'statePayerNumber': 'Box 17 - State/Payer\'s State No.',
      'stateIncome': 'Box 18 - State Income',
      
      // Additional fields that might be extracted
      'medicalPaymentsMultiple': 'Multiple Medical Payments Found'
    };
    
    for (const [fieldKey, displayName] of Object.entries(fieldMappings)) {
      if (extractedData[fieldKey] !== undefined && extractedData[fieldKey] !== null && extractedData[fieldKey] !== '') {
        let fieldValue = extractedData[fieldKey];
        
        // Handle special cases
        if (fieldKey === 'medicalPaymentsMultiple' && Array.isArray(fieldValue)) {
          // Convert array of medical payments to string
          fieldValue = fieldValue.map(amount => `$${amount}`).join(', ');
        }
        
        entries.push({
          fieldName: displayName,
          fieldValue: String(fieldValue),
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
      'federalTaxWithheld': 'Box 4 - Federal Tax Withheld'
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
    
    // Process all available fields for generic documents
    for (const [fieldKey, fieldValue] of Object.entries(extractedData)) {
      if (fieldKey !== 'fullText' && fieldKey !== 'correctedDocumentType' && 
          fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
        entries.push({
          fieldName: fieldKey,
          fieldValue: String(fieldValue),
          confidence: 0.85
        });
      }
    }
    
    return entries;
  }

