import { NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { W2ToForm1040Mapper } from "@/lib/w2-to-1040-mapping";
import { Form1099ToForm1040Mapper } from "@/lib/1099-to-1040-mapping";
import { Form1040Data } from "@/lib/form-1040-types";

export const dynamic = "force-dynamic";

// GET: Retrieve 1040 form data with W2 mappings
export async function GET(
  request: Request,
  { params }: { params: { id: string } }
) {
  console.log("🔍 [1040 GET] Starting form 1040 data retrieval for tax return:", params.id);
  
  try {
    const session = await getServerSession(authOptions);
    
    if (!session?.user?.email) {
      console.log("❌ [1040 GET] No session found");
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
    }

    const user = await prisma.user.findUnique({
      where: { email: session.user.email }
    });

    if (!user) {
      console.log("❌ [1040 GET] User not found");
      return NextResponse.json({ error: "User not found" }, { status: 404 });
    }

    // Get tax return with all related data
    const taxReturn = await prisma.taxReturn.findFirst({
      where: { 
        id: params.id,
        userId: user.id 
      },
      include: {
        incomeEntries: true,
        deductionEntries: true,
        dependents: true,
        documents: {
          where: { 
            OR: [
              { documentType: 'W2', processingStatus: 'COMPLETED' },
              { documentType: 'FORM_1099_INT', processingStatus: 'COMPLETED' },
              { documentType: 'FORM_1099_DIV', processingStatus: 'COMPLETED' },
              { documentType: 'FORM_1099_MISC', processingStatus: 'COMPLETED' },
              { documentType: 'FORM_1099_NEC', processingStatus: 'COMPLETED' }
            ]
          },
          include: {
            extractedEntries: true
          }
        }
      }
    });

    if (!taxReturn) {
      console.log("❌ [1040 GET] Tax return not found");
      return NextResponse.json({ error: "Tax return not found" }, { status: 404 });
    }

    // Check if there's existing 1040 data
    let form1040Data: Partial<Form1040Data> = {};
    
    // If there's saved 1040 data in a separate table or JSON field, load it
    // For now, we'll construct it from the tax return data
    
    // Get W2 and 1099 documents and their extracted data
    const w2Documents = taxReturn.documents.filter((doc: any) => doc.documentType === 'W2');
    const form1099Documents = taxReturn.documents.filter((doc: any) => 
      ['FORM_1099_INT', 'FORM_1099_DIV', 'FORM_1099_MISC', 'FORM_1099_NEC'].includes(doc.documentType)
    );
    const w2MappingData = [];
    const form1099MappingData = [];

    // Process each W2 document and map to 1040 form
    console.log(`🔍 [1040 GET] Processing ${w2Documents.length} W2 documents`);
    
    for (const w2Doc of w2Documents) {
      console.log(`🔍 [1040 GET] Processing W2 document: ${w2Doc.fileName} (ID: ${w2Doc.id})`);
      console.log(`🔍 [1040 GET] W2 document extractedData:`, JSON.stringify(w2Doc.extractedData, null, 2));
      
      if (w2Doc.extractedData && typeof w2Doc.extractedData === 'object') {
        const extractedData = w2Doc.extractedData as any;
        
        // Try different data structure paths
        let w2DataToMap = extractedData.extractedData || extractedData;
        
        if (process.env.NODE_ENV === 'development') {
          console.log(`🔍 [1040 GET] Data to map to 1040:`, JSON.stringify(w2DataToMap, null, 2));
          console.log(`🔍 [1040 GET] Current form1040Data before mapping:`, JSON.stringify(form1040Data, null, 2));
        }
        
        // DEBUG: Check for personal info fields in the W2 data
        console.log(`🔍 [1040 GET DEBUG] Checking W2 personal info fields:`);
        console.log(`  - employeeName: ${w2DataToMap.employeeName}`);
        console.log(`  - Employee?.Name: ${w2DataToMap.Employee?.Name}`);
        console.log(`  - Employee.Name: ${w2DataToMap['Employee.Name']}`);
        console.log(`  - employeeSSN: ${w2DataToMap.employeeSSN}`);
        console.log(`  - Employee?.SSN: ${w2DataToMap.Employee?.SSN}`);
        console.log(`  - Employee.SSN: ${w2DataToMap['Employee.SSN']}`);
        console.log(`  - employeeAddress: ${w2DataToMap.employeeAddress}`);
        console.log(`  - Employee?.Address: ${w2DataToMap.Employee?.Address}`);
        console.log(`  - Employee.Address: ${w2DataToMap['Employee.Address']}`);
        
        // Map W2 data to 1040 form fields
        const mappedData = W2ToForm1040Mapper.mapW2ToForm1040(
          w2DataToMap, 
          form1040Data
        );
        
        if (process.env.NODE_ENV === 'development') {
          console.log(`🔍 [1040 GET] Mapped data from W2:`, JSON.stringify(mappedData, null, 2));
        }
        
        // DEBUG: Check if personalInfo was created in mappedData
        if (mappedData.personalInfo) {
          console.log(`✅ [1040 GET DEBUG] personalInfo was created:`, JSON.stringify(mappedData.personalInfo, null, 2));
        } else {
          console.log(`❌ [1040 GET DEBUG] personalInfo was NOT created in mappedData`);
        }
        
        // Merge the mapped data
        form1040Data = { ...form1040Data, ...mappedData };
        
        if (process.env.NODE_ENV === 'development') {
          console.log(`🔍 [1040 GET] Form1040Data after merging:`, JSON.stringify(form1040Data, null, 2));
        }
        
        // DEBUG: Check if personalInfo exists in final form1040Data
        if (form1040Data.personalInfo) {
          console.log(`✅ [1040 GET DEBUG] personalInfo exists in final form1040Data:`, JSON.stringify(form1040Data.personalInfo, null, 2));
        } else {
          console.log(`❌ [1040 GET DEBUG] personalInfo is MISSING from final form1040Data`);
        }
        
        // Create mapping summary
        const mappingSummary = W2ToForm1040Mapper.createMappingSummary(w2DataToMap);
        
        w2MappingData.push({
          documentId: w2Doc.id,
          fileName: w2Doc.fileName,
          mappings: mappingSummary
        });
        
        console.log(`✅ [1040 GET] Successfully processed W2 document: ${w2Doc.fileName}`);
      } else {
        console.log(`⚠️ [1040 GET] W2 document ${w2Doc.fileName} has no extractedData or invalid format`);
      }
    }

    // Process each 1099 document and map to 1040 form
    console.log(`🔍 [1040 GET] Processing ${form1099Documents.length} 1099 documents`);
    
    for (const form1099Doc of form1099Documents) {
      console.log(`🔍 [1040 GET] Processing 1099 document: ${form1099Doc.fileName} (ID: ${form1099Doc.id})`);
      console.log(`🔍 [1040 GET] 1099 document extractedData:`, JSON.stringify(form1099Doc.extractedData, null, 2));
      
      if (form1099Doc.extractedData && typeof form1099Doc.extractedData === 'object') {
        const extractedData = form1099Doc.extractedData as any;
        
        // Try different data structure paths
        let form1099DataToMap = extractedData.extractedData || extractedData;
        
        if (process.env.NODE_ENV === 'development') {
          console.log(`🔍 [1040 GET] 1099 data to map to 1040:`, JSON.stringify(form1099DataToMap, null, 2));
          console.log(`🔍 [1040 GET] Current form1040Data before 1099 mapping:`, JSON.stringify(form1040Data, null, 2));
        }
        
        // DEBUG: Check for personal info fields in the 1099 data
        console.log(`🔍 [1040 GET DEBUG] Checking 1099 personal info fields:`);
        console.log(`  - recipientName: ${form1099DataToMap.recipientName}`);
        console.log(`  - Recipient?.Name: ${form1099DataToMap.Recipient?.Name}`);
        console.log(`  - Recipient.Name: ${form1099DataToMap['Recipient.Name']}`);
        console.log(`  - recipientTIN: ${form1099DataToMap.recipientTIN}`);
        console.log(`  - Recipient?.TIN: ${form1099DataToMap.Recipient?.TIN}`);
        console.log(`  - Recipient.TIN: ${form1099DataToMap['Recipient.TIN']}`);
        console.log(`  - recipientAddress: ${form1099DataToMap.recipientAddress}`);
        console.log(`  - Recipient?.Address: ${form1099DataToMap.Recipient?.Address}`);
        console.log(`  - Recipient.Address: ${form1099DataToMap['Recipient.Address']}`);
        
        // Map 1099 data to 1040 form fields
        const mappedData = Form1099ToForm1040Mapper.map1099ToForm1040(
          form1099DataToMap, 
          form1040Data
        );
        
        if (process.env.NODE_ENV === 'development') {
          console.log(`🔍 [1040 GET] Mapped data from 1099:`, JSON.stringify(mappedData, null, 2));
        }
        
        // DEBUG: Check if personalInfo was updated in mappedData
        if (mappedData.personalInfo) {
          console.log(`✅ [1040 GET DEBUG] personalInfo was updated by 1099:`, JSON.stringify(mappedData.personalInfo, null, 2));
        } else {
          console.log(`❌ [1040 GET DEBUG] personalInfo was NOT updated by 1099 mapping`);
        }
        
        // Merge the mapped data
        form1040Data = { ...form1040Data, ...mappedData };
        
        if (process.env.NODE_ENV === 'development') {
          console.log(`🔍 [1040 GET] Form1040Data after merging 1099:`, JSON.stringify(form1040Data, null, 2));
        }
        
        // DEBUG: Check if personalInfo exists in final form1040Data after 1099
        if (form1040Data.personalInfo) {
          console.log(`✅ [1040 GET DEBUG] personalInfo exists in form1040Data after 1099:`, JSON.stringify(form1040Data.personalInfo, null, 2));
        } else {
          console.log(`❌ [1040 GET DEBUG] personalInfo is MISSING from form1040Data after 1099`);
        }
        
        // Create mapping summary
        const mappingSummary = Form1099ToForm1040Mapper.createMappingSummary(form1099DataToMap);
        
        form1099MappingData.push({
          documentId: form1099Doc.id,
          fileName: form1099Doc.fileName,
          documentType: form1099Doc.documentType,
          mappings: mappingSummary
        });
        
        console.log(`✅ [1040 GET] Successfully processed 1099 document: ${form1099Doc.fileName}`);
      } else {
        console.log(`⚠️ [1040 GET] 1099 document ${form1099Doc.fileName} has no extractedData or invalid format`);
      }
    }

    // CRITICAL FIX: Process accepted income entries from the database
    // This handles 1099 data that was accepted by the user but may not be in document extractedData
    // IMPORTANT: Only process income entries that are NOT already included in W2/1099 document mapping
    console.log(`🔍 [1040 GET] Processing ${taxReturn.incomeEntries.length} accepted income entries`);
    
    // Track which documents have already been processed to avoid double-counting
    const processedDocumentIds = new Set([
      ...w2Documents.map(doc => doc.id),
      ...form1099Documents.map(doc => doc.id)
    ]);
    
    // Determine if we have processed W2 and 1099 documents to avoid double-counting
    const hasProcessedW2Documents = w2Documents.length > 0;
    const hasProcessed1099Documents = form1099Documents.length > 0;
    
    console.log(`🔍 [1040 GET] Document processing flags - W2: ${hasProcessedW2Documents}, 1099: ${hasProcessed1099Documents}`);
    
    for (const incomeEntry of taxReturn.incomeEntries) {
      console.log(`🔍 [1040 GET] Processing income entry: ${incomeEntry.incomeType} - $${incomeEntry.amount}`);
      
      // Skip income entries that come from documents we've already processed
      if (incomeEntry.documentId && processedDocumentIds.has(incomeEntry.documentId)) {
        console.log(`⚠️ [1040 GET] Skipping income entry from already processed document: ${incomeEntry.documentId}`);
        continue;
      }
      
      // DOUBLE-COUNTING FIX: Skip W2_WAGES entries if W2 documents have been processed
      if (incomeEntry.incomeType === 'W2_WAGES' && hasProcessedW2Documents) {
        console.log(`ℹ️ [1040 GET] Skipping W2_WAGES database entry because W2 documents have been processed`);
        continue;
      }
      
      // DOUBLE-COUNTING FIX: Skip OTHER_INCOME entries if 1099 documents have been processed
      if (incomeEntry.incomeType === 'OTHER_INCOME' && hasProcessed1099Documents) {
        console.log(`ℹ️ [1040 GET] Skipping OTHER_INCOME database entry because 1099 documents have been processed`);
        continue;
      }
      
      // DOUBLE-COUNTING FIX: Skip INTEREST entries if 1099 documents have been processed
      if (incomeEntry.incomeType === 'INTEREST' && hasProcessed1099Documents) {
        console.log(`ℹ️ [1040 GET] Skipping INTEREST database entry because 1099 documents have been processed`);
        continue;
      }
      
      // DOUBLE-COUNTING FIX: Skip DIVIDENDS entries if 1099 documents have been processed
      if (incomeEntry.incomeType === 'DIVIDENDS' && hasProcessed1099Documents) {
        console.log(`ℹ️ [1040 GET] Skipping DIVIDENDS database entry because 1099 documents have been processed`);
        continue;
      }
      
      // Only process income entries that are manually entered or from unprocessed documents
      // Map income entries to 1040 form lines based on income type
      switch (incomeEntry.incomeType) {
        case 'W2_WAGES':
          // Only add if this is a manual entry or from an unprocessed document
          if (!incomeEntry.documentId) {
            form1040Data.line1 = (form1040Data.line1 || 0) + (incomeEntry.amount ? incomeEntry.amount.toNumber() : 0);
            form1040Data.line25a = (form1040Data.line25a || 0) + (incomeEntry.federalTaxWithheld ? incomeEntry.federalTaxWithheld.toNumber() : 0);
            console.log(`✅ [1040 GET] Added manual W2 wages to Line 1: $${incomeEntry.amount}, withholding to Line 25a: $${incomeEntry.federalTaxWithheld || 0}`);
          }
          break;
          
        case 'INTEREST':
          form1040Data.line2b = (form1040Data.line2b || 0) + (incomeEntry.amount ? incomeEntry.amount.toNumber() : 0);
          form1040Data.line25a = (form1040Data.line25a || 0) + (incomeEntry.federalTaxWithheld ? incomeEntry.federalTaxWithheld.toNumber() : 0);
          console.log(`✅ [1040 GET] Added interest income to Line 2b: $${incomeEntry.amount}, withholding to Line 25a: $${incomeEntry.federalTaxWithheld || 0}`);
          break;
          
        case 'DIVIDENDS':
          form1040Data.line3b = (form1040Data.line3b || 0) + (incomeEntry.amount ? incomeEntry.amount.toNumber() : 0);
          form1040Data.line25a = (form1040Data.line25a || 0) + (incomeEntry.federalTaxWithheld ? incomeEntry.federalTaxWithheld.toNumber() : 0);
          console.log(`✅ [1040 GET] Added dividend income to Line 3b: $${incomeEntry.amount}, withholding to Line 25a: $${incomeEntry.federalTaxWithheld || 0}`);
          break;
          
        case 'OTHER_INCOME':
          form1040Data.line8 = (form1040Data.line8 || 0) + (incomeEntry.amount ? incomeEntry.amount.toNumber() : 0);
          form1040Data.line25a = (form1040Data.line25a || 0) + (incomeEntry.federalTaxWithheld ? incomeEntry.federalTaxWithheld.toNumber() : 0);
          console.log(`✅ [1040 GET] Added other income to Line 8: $${incomeEntry.amount}, withholding to Line 25a: $${incomeEntry.federalTaxWithheld || 0}`);
          break;
          
        case 'UNEMPLOYMENT':
          form1040Data.line8 = (form1040Data.line8 || 0) + (incomeEntry.amount ? incomeEntry.amount.toNumber() : 0);
          form1040Data.line25a = (form1040Data.line25a || 0) + (incomeEntry.federalTaxWithheld ? incomeEntry.federalTaxWithheld.toNumber() : 0);
          console.log(`✅ [1040 GET] Added unemployment income to Line 8: $${incomeEntry.amount}, withholding to Line 25a: $${incomeEntry.federalTaxWithheld || 0}`);
          break;
          
        case 'RETIREMENT_DISTRIBUTIONS':
          form1040Data.line4b = (form1040Data.line4b || 0) + (incomeEntry.amount ? incomeEntry.amount.toNumber() : 0);
          form1040Data.line25a = (form1040Data.line25a || 0) + (incomeEntry.federalTaxWithheld ? incomeEntry.federalTaxWithheld.toNumber() : 0);
          console.log(`✅ [1040 GET] Added retirement distributions to Line 4b: $${incomeEntry.amount}, withholding to Line 25a: $${incomeEntry.federalTaxWithheld || 0}`);
          break;
          
        case 'SOCIAL_SECURITY':
          form1040Data.line5b = (form1040Data.line5b || 0) + (incomeEntry.amount ? incomeEntry.amount.toNumber() : 0);
          console.log(`✅ [1040 GET] Added social security benefits to Line 5b: $${incomeEntry.amount}`);
          break;
          
        default:
          // Default to other income (Line 8) for unknown types
          form1040Data.line8 = (form1040Data.line8 || 0) + (incomeEntry.amount ? incomeEntry.amount.toNumber() : 0);
          form1040Data.line25a = (form1040Data.line25a || 0) + (incomeEntry.federalTaxWithheld ? incomeEntry.federalTaxWithheld.toNumber() : 0);
          console.log(`✅ [1040 GET] Added unknown income type '${incomeEntry.incomeType}' to Line 8: $${incomeEntry.amount}, withholding to Line 25a: $${incomeEntry.federalTaxWithheld || 0}`);
          break;
      }
    }

    // Recalculate totals after processing all income entries
    form1040Data.line9 = (form1040Data.line1 || 0) + (form1040Data.line2b || 0) + (form1040Data.line3b || 0) + 
                        (form1040Data.line4b || 0) + (form1040Data.line5b || 0) + (form1040Data.line6b || 0) + 
                        (form1040Data.line7 || 0) + (form1040Data.line8 || 0);
    
    form1040Data.line11 = form1040Data.line9 - (form1040Data.line10 || 0);
    
    console.log(`✅ [1040 GET] Recalculated totals - Line 9 (Total Income): $${form1040Data.line9}, Line 11 (AGI): $${form1040Data.line11}, Line 25a (Total Withholdings): $${form1040Data.line25a || 0}`);
    

    // Fill in basic info from tax return if not already populated
    // IMPORTANT: Only fill from taxReturn if we don't have W2 personal info
    const hasW2PersonalInfo = form1040Data.personalInfo && (
      form1040Data.personalInfo.firstName || 
      form1040Data.personalInfo.lastName || 
      form1040Data.personalInfo.ssn || 
      form1040Data.personalInfo.address
    );
    
    console.log(`🔍 [1040 GET DEBUG] hasW2PersonalInfo: ${hasW2PersonalInfo}`);
    console.log(`🔍 [1040 GET DEBUG] form1040Data.firstName: ${form1040Data.firstName}`);
    
    if (!form1040Data.firstName && !hasW2PersonalInfo) {
      console.log("🔍 [1040 GET] No W2 personal info found, filling from taxReturn data");
      form1040Data.firstName = taxReturn.firstName || '';
      form1040Data.lastName = taxReturn.lastName || '';
      form1040Data.ssn = taxReturn.ssn || '';
      form1040Data.spouseFirstName = taxReturn.spouseFirstName || undefined;
      form1040Data.spouseLastName = taxReturn.spouseLastName || undefined;
      form1040Data.spouseSSN = taxReturn.spouseSsn || undefined;
      form1040Data.address = taxReturn.address || '';
      form1040Data.city = taxReturn.city || '';
      form1040Data.state = taxReturn.state || '';
      form1040Data.zipCode = taxReturn.zipCode || '';
      form1040Data.filingStatus = taxReturn.filingStatus as any;
      form1040Data.taxYear = taxReturn.taxYear;
    } else if (hasW2PersonalInfo) {
      console.log("✅ [1040 GET] W2 personal info exists, preserving it and ensuring top-level fields are set");
      // Ensure top-level fields are set from W2 data if they exist
      if (!form1040Data.firstName && form1040Data.personalInfo?.firstName) {
        form1040Data.firstName = form1040Data.personalInfo.firstName;
      }
      if (!form1040Data.lastName && form1040Data.personalInfo?.lastName) {
        form1040Data.lastName = form1040Data.personalInfo.lastName;
      }
      if (!form1040Data.ssn && form1040Data.personalInfo?.ssn) {
        form1040Data.ssn = form1040Data.personalInfo.ssn;
      }
      if (!form1040Data.address && form1040Data.personalInfo?.address) {
        form1040Data.address = form1040Data.personalInfo.address;
      }
      if (!form1040Data.city && form1040Data.personalInfo?.city) {
        form1040Data.city = form1040Data.personalInfo.city;
      }
      if (!form1040Data.state && form1040Data.personalInfo?.state) {
        form1040Data.state = form1040Data.personalInfo.state;
      }
      if (!form1040Data.zipCode && form1040Data.personalInfo?.zipCode) {
        form1040Data.zipCode = form1040Data.personalInfo.zipCode;
      }
      
      // Set other required fields from taxReturn
      form1040Data.filingStatus = form1040Data.filingStatus || taxReturn.filingStatus as any;
      form1040Data.taxYear = form1040Data.taxYear || taxReturn.taxYear;
    } else {
      console.log("🔍 [1040 GET] Top-level firstName exists, filling missing fields from taxReturn");
      // Fill in missing fields from taxReturn without overriding existing data
      form1040Data.lastName = form1040Data.lastName || taxReturn.lastName || '';
      form1040Data.ssn = form1040Data.ssn || taxReturn.ssn || '';
      form1040Data.spouseFirstName = form1040Data.spouseFirstName || taxReturn.spouseFirstName || undefined;
      form1040Data.spouseLastName = form1040Data.spouseLastName || taxReturn.spouseLastName || undefined;
      form1040Data.spouseSSN = form1040Data.spouseSSN || taxReturn.spouseSsn || undefined;
      form1040Data.address = form1040Data.address || taxReturn.address || '';
      form1040Data.city = form1040Data.city || taxReturn.city || '';
      form1040Data.state = form1040Data.state || taxReturn.state || '';
      form1040Data.zipCode = form1040Data.zipCode || taxReturn.zipCode || '';
      form1040Data.filingStatus = form1040Data.filingStatus || taxReturn.filingStatus as any;
      form1040Data.taxYear = form1040Data.taxYear || taxReturn.taxYear;
    }

    // ===== CRITICAL TAX CALCULATION FIX - ADD THIS BLOCK =====
    
    // Import tax calculation functions
    const { calculateTaxLiability, getStandardDeduction } = await import('@/lib/tax-calculations');

    // Get filing status for calculations
    const filingStatus = form1040Data.filingStatus || taxReturn.filingStatus;
    const adjustedGrossIncome = form1040Data.line11 || 0;

    console.log(`🧮 [1040 GET] Starting tax calculations for ${filingStatus} with AGI: $${adjustedGrossIncome}`);

    // Calculate standard deduction (Line 12)
    const standardDeduction = getStandardDeduction(filingStatus);
    form1040Data.line12 = standardDeduction;

    // Calculate taxable income (Line 15)
    const taxableIncome = Math.max(0, adjustedGrossIncome - standardDeduction);
    form1040Data.line15 = taxableIncome;

    // Calculate tax liability (Line 16)
    const taxLiability = calculateTaxLiability(taxableIncome, filingStatus);
    form1040Data.line16 = taxLiability;

    // Calculate total credits (for now, assume 0 unless you have credit calculations)
    const totalCredits = form1040Data.line19 || 0;

    // Calculate total tax after credits (Line 24)
    const totalTax = Math.max(0, taxLiability - totalCredits);
    form1040Data.line24 = totalTax;

    // Total payments and withholdings (Line 32)
    const totalPayments = form1040Data.line25a || 0;
    form1040Data.line32 = totalPayments;

    // Calculate refund or amount owed
    if (totalPayments > totalTax) {
      // Refund (Line 33)
      form1040Data.line33 = totalPayments - totalTax;
      form1040Data.line37 = 0;
    } else {
      // Amount owed (Line 37)
      form1040Data.line33 = 0;
      form1040Data.line37 = totalTax - totalPayments;
    }

    console.log(`✅ [1040 GET] Tax calculations complete:`);
    console.log(`  - Standard Deduction (Line 12): $${standardDeduction}`);
    console.log(`  - Taxable Income (Line 15): $${taxableIncome}`);
    console.log(`  - Tax Liability (Line 16): $${taxLiability}`);
    console.log(`  - Total Tax (Line 24): $${totalTax}`);
    console.log(`  - Total Payments (Line 32): $${totalPayments}`);
    console.log(`  - Refund (Line 33): $${form1040Data.line33}`);
    console.log(`  - Amount Owed (Line 37): $${form1040Data.line37}`);

    // ===== END OF TAX CALCULATION FIX =====

    console.log("✅ [1040 GET] Successfully retrieved 1040 form data");
    if (process.env.NODE_ENV === 'development') {
      console.log("🔍 [1040 GET] Final form1040Data being returned:", JSON.stringify(form1040Data, null, 2));
      console.log("🔍 [1040 GET] W2 mapping data being returned:", JSON.stringify(w2MappingData, null, 2));
    }
    
    // FINAL DEBUG: Ensure personalInfo is preserved in the response
    if (form1040Data.personalInfo) {
      console.log("✅ [1040 GET FINAL] personalInfo is present in final response:", JSON.stringify(form1040Data.personalInfo, null, 2));
    } else {
      console.log("❌ [1040 GET FINAL] personalInfo is MISSING from final response!");
      console.log("🔍 [1040 GET FINAL] Available keys in form1040Data:", Object.keys(form1040Data));
    }
    
    return NextResponse.json({
      form1040Data,
      w2MappingData,
      form1099MappingData,
      taxReturn: {
        id: taxReturn.id,
        taxYear: taxReturn.taxYear,
        filingStatus: taxReturn.filingStatus
      }
    });

  } catch (error) {
    console.error("💥 [1040 GET] Error retrieving form 1040 data:", error);
    return NextResponse.json(
      { error: "Internal server error" },
      { status: 500 }
    );
  }
}

// POST: Save 1040 form data
export async function POST(
  request: Request,
  { params }: { params: { id: string } }
) {
  console.log("🔍 [1040 POST] Starting form 1040 data save for tax return:", params.id);
  
  try {
    const session = await getServerSession(authOptions);
    
    if (!session?.user?.email) {
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
    }

    const user = await prisma.user.findUnique({
      where: { email: session.user.email }
    });

    if (!user) {
      return NextResponse.json({ error: "User not found" }, { status: 404 });
    }

    const body = await request.json();
    const { form1040Data }: { form1040Data: Form1040Data } = body;

    // Verify tax return ownership
    const taxReturn = await prisma.taxReturn.findFirst({
      where: { 
        id: params.id,
        userId: user.id 
      }
    });

    if (!taxReturn) {
      return NextResponse.json({ error: "Tax return not found" }, { status: 404 });
    }

    // Update tax return with 1040 form data
    const updatedTaxReturn = await prisma.taxReturn.update({
      where: { id: params.id },
      data: {
        firstName: form1040Data.firstName,
        lastName: form1040Data.lastName,
        ssn: form1040Data.ssn,
        spouseFirstName: form1040Data.spouseFirstName,
        spouseLastName: form1040Data.spouseLastName,
        spouseSsn: form1040Data.spouseSSN,
        address: form1040Data.address,
        city: form1040Data.city,
        state: form1040Data.state,
        zipCode: form1040Data.zipCode,
        filingStatus: form1040Data.filingStatus as any,
        
        // Tax calculation fields
        totalIncome: form1040Data.line9,
        adjustedGrossIncome: form1040Data.line11,
        standardDeduction: form1040Data.line12,
        taxableIncome: form1040Data.line15,
        taxLiability: form1040Data.line16,
        totalWithholdings: form1040Data.line25a,
        refundAmount: form1040Data.line33,
        amountOwed: form1040Data.line37,
        
        lastSavedAt: new Date()
      }
    });

    // Store full 1040 form data as JSON in a custom field or separate table
    // For now, we'll store it as extractedData in a document record
    
    // First, try to find an existing 1040 document
    const existingForm1040Doc = await prisma.document.findFirst({
      where: {
        taxReturnId: params.id,
        documentType: 'OTHER_TAX_DOCUMENT',
        fileName: { contains: 'Form_1040' }
      }
    });

    let form1040Document;
    if (existingForm1040Doc) {
      // Update existing document
      form1040Document = await prisma.document.update({
        where: { id: existingForm1040Doc.id },
        data: {
          extractedData: form1040Data as any,
          processingStatus: 'COMPLETED',
          fileName: `Form_1040_${form1040Data.taxYear}.json`
        }
      });
    } else {
      // Create new document
      form1040Document = await prisma.document.create({
        data: {
          taxReturnId: params.id,
          fileName: `Form_1040_${form1040Data.taxYear}.json`,
          fileType: 'application/json',
          fileSize: JSON.stringify(form1040Data).length,
          filePath: '',
          documentType: 'OTHER_TAX_DOCUMENT',
          processingStatus: 'COMPLETED',
          extractedData: form1040Data as any
        }
      });
    }

    console.log("✅ [1040 POST] Successfully saved 1040 form data");
    
    return NextResponse.json({
      success: true,
      taxReturn: updatedTaxReturn,
      form1040Document: form1040Document
    });

  } catch (error) {
    console.error("💥 [1040 POST] Error saving form 1040 data:", error);
    return NextResponse.json(
      { error: "Internal server error" },
      { status: 500 }
    );
  }
}
