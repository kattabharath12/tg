
import { DocumentAnalysisClient, AzureKeyCredential } from "@azure/ai-form-recognizer";
import { DocumentType } from "@prisma/client";
import { readFile } from "fs/promises";

export interface AzureDocumentIntelligenceConfig {
  endpoint: string;
  apiKey: string;
}

export interface ExtractedFieldData {
  [key: string]: string | number | DocumentType | number[] | undefined;
  correctedDocumentType?: DocumentType;
  fullText?: string;
}

export class AzureDocumentIntelligenceService {
  private client: DocumentAnalysisClient;
  private config: AzureDocumentIntelligenceConfig;

  constructor(config: AzureDocumentIntelligenceConfig) {
    this.config = config;
    this.client = new DocumentAnalysisClient(
      this.config.endpoint,
      new AzureKeyCredential(this.config.apiKey)
    );
  }

  async extractDataFromDocument(
    documentPathOrBuffer: string | Buffer,
    documentType: string
  ): Promise<ExtractedFieldData> {
    try {
      console.log('🔍 [Azure DI] Processing document with Azure Document Intelligence...');
      console.log('🔍 [Azure DI] Initial document type:', documentType);
      
      // Get document buffer - either from file path or use provided buffer
      const documentBuffer = typeof documentPathOrBuffer === 'string' 
        ? await readFile(documentPathOrBuffer)
        : documentPathOrBuffer;
      
      // Determine the model to use based on document type
      const modelId = this.getModelIdForDocumentType(documentType);
      console.log('🔍 [Azure DI] Using model:', modelId);
      
      let extractedData: ExtractedFieldData;
      let correctedDocumentType: DocumentType | undefined;
      
      try {
        // Analyze the document with specific tax model
        const poller = await this.client.beginAnalyzeDocument(modelId, documentBuffer);
        const result = await poller.pollUntilDone();
        
        console.log('✅ [Azure DI] Document analysis completed with tax model');
        
        // Extract the data based on document type
        extractedData = this.extractTaxDocumentFields(result, documentType);
        
        // Perform OCR-based document type correction if we have OCR text
        if (extractedData.fullText) {
          const ocrBasedType = this.analyzeDocumentTypeFromOCR(extractedData.fullText as string);
          if (ocrBasedType !== 'UNKNOWN' && ocrBasedType !== documentType) {
            console.log(`🔄 [Azure DI] Document type correction: ${documentType} → ${ocrBasedType}`);
            
            // Convert string to DocumentType enum with validation
            if (Object.values(DocumentType).includes(ocrBasedType as DocumentType)) {
              correctedDocumentType = ocrBasedType as DocumentType;
              
              // Re-extract data with the corrected document type
              console.log('🔍 [Azure DI] Re-extracting data with corrected document type...');
              extractedData = this.extractTaxDocumentFields(result, ocrBasedType);
            } else {
              console.log(`⚠️ [Azure DI] Invalid document type detected: ${ocrBasedType}, ignoring correction`);
            }
          }
        }
        
      } catch (modelError: any) {
        console.warn('⚠️ [Azure DI] Tax model failed, attempting fallback to OCR model:', modelError?.message);
        
        // Check if it's a ModelNotFound error
        if (modelError?.message?.includes('ModelNotFound') || 
            modelError?.message?.includes('Resource not found') ||
            modelError?.code === 'NotFound') {
          
          console.log('🔍 [Azure DI] Falling back to prebuilt-read model for OCR extraction...');
          
          // Fallback to general OCR model
          const fallbackPoller = await this.client.beginAnalyzeDocument('prebuilt-read', documentBuffer);
          const fallbackResult = await fallbackPoller.pollUntilDone();
          
          console.log('✅ [Azure DI] Document analysis completed with OCR fallback');
          
          // Extract data using OCR-based approach
          extractedData = this.extractTaxDocumentFieldsFromOCR(fallbackResult, documentType);
          
          // Perform OCR-based document type correction
          if (extractedData.fullText) {
            const ocrBasedType = this.analyzeDocumentTypeFromOCR(extractedData.fullText as string);
            if (ocrBasedType !== 'UNKNOWN' && ocrBasedType !== documentType) {
              console.log(`🔄 [Azure DI] Document type correction (OCR fallback): ${documentType} → ${ocrBasedType}`);
              
              // Convert string to DocumentType enum with validation
              if (Object.values(DocumentType).includes(ocrBasedType as DocumentType)) {
                correctedDocumentType = ocrBasedType as DocumentType;
                
                // Re-extract data with the corrected document type
                console.log('🔍 [Azure DI] Re-extracting data with corrected document type...');
                extractedData = this.extractTaxDocumentFieldsFromOCR(fallbackResult, ocrBasedType);
              } else {
                console.log(`⚠️ [Azure DI] Invalid document type detected: ${ocrBasedType}, ignoring correction`);
              }
            }
          }
        } else {
          // Re-throw if it's not a model availability issue
          throw modelError;
        }
      }
      
      // Add the corrected document type to the result if it was changed
      if (correctedDocumentType) {
        extractedData.correctedDocumentType = correctedDocumentType;
      }
      
      return extractedData;
    } catch (error: any) {
      console.error('❌ [Azure DI] Processing error:', error);
      throw new Error(`Azure Document Intelligence processing failed: ${error?.message || 'Unknown error'}`);
    }
  }

  private getModelIdForDocumentType(documentType: string): string {
    switch (documentType) {
      case 'W2':
        return 'prebuilt-tax.us.w2';
      case 'FORM_1099_INT':
      case 'FORM_1099_DIV':
      case 'FORM_1099_MISC':
      case 'FORM_1099_NEC':
        // All 1099 variants use the unified 1099 model
        return 'prebuilt-tax.us.1099';
      default:
        // Use general document model for other types
        return 'prebuilt-document';
    }
  }

  private extractTaxDocumentFieldsFromOCR(result: any, documentType: string): ExtractedFieldData {
    console.log('🔍 [Azure DI] Extracting tax document fields using OCR fallback...');
    
    const extractedData: ExtractedFieldData = {};
    
    // Extract text content from OCR result
    extractedData.fullText = result.content || '';
    
    // Use OCR-based extraction methods for different document types
    switch (documentType) {
      case 'W2':
        return this.extractW2FieldsFromOCR(extractedData.fullText as string, extractedData);
      case 'FORM_1099_INT':
        return this.extract1099IntFieldsFromOCR(extractedData.fullText as string, extractedData);
      case 'FORM_1099_DIV':
        return this.extract1099DivFieldsFromOCR(extractedData.fullText as string, extractedData);
      case 'FORM_1099_MISC':
        return this.extract1099MiscFieldsFromOCR(extractedData.fullText as string, extractedData);
      case 'FORM_1099_NEC':
        return this.extract1099NecFieldsFromOCR(extractedData.fullText as string, extractedData);
      default:
        console.log('🔍 [Azure DI] Using generic OCR extraction for document type:', documentType);
        return this.extractGenericFieldsFromOCR(extractedData.fullText as string, extractedData);
    }
  }

  private extractTaxDocumentFields(result: any, documentType: string): ExtractedFieldData {
    const extractedData: ExtractedFieldData = {};
    
    // Extract text content
    extractedData.fullText = result.content || '';
    
    // Extract form fields
    if (result.documents && result.documents.length > 0) {
      const document = result.documents[0];
      
      if (document.fields) {
        // Process fields based on document type
        switch (documentType) {
          case 'W2':
            return this.processW2Fields(document.fields, extractedData);
          case 'FORM_1099_INT':
            return this.process1099IntFields(document.fields, extractedData);
          case 'FORM_1099_DIV':
            return this.process1099DivFields(document.fields, extractedData);
          case 'FORM_1099_MISC':
            return this.process1099MiscFields(document.fields, extractedData);
          case 'FORM_1099_NEC':
            return this.process1099NecFields(document.fields, extractedData);
          default:
            return this.processGenericFields(document.fields, extractedData);
        }
      }
    }
    
    // Extract key-value pairs from tables if available
    if (result.keyValuePairs) {
      for (const kvp of result.keyValuePairs) {
        const key = kvp.key?.content?.trim();
        const value = kvp.value?.content?.trim();
        if (key && value) {
          extractedData[key] = value;
        }
      }
    }
    
    return extractedData;
  }

  private processW2Fields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    const w2Data = { ...baseData };
    
    // W2 specific field mappings
    const w2FieldMappings = {
      'Employee.Name': 'employeeName',
      'Employee.SSN': 'employeeSSN',
      'Employee.Address': 'employeeAddress',
      'Employer.Name': 'employerName',
      'Employer.EIN': 'employerEIN',
      'Employer.Address': 'employerAddress',
      'WagesAndTips': 'wages',
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',
      'SocialSecurityWages': 'socialSecurityWages',
      'SocialSecurityTaxWithheld': 'socialSecurityTaxWithheld',
      'MedicareWagesAndTips': 'medicareWages',
      'MedicareTaxWithheld': 'medicareTaxWithheld',
      'SocialSecurityTips': 'socialSecurityTips',
      'AllocatedTips': 'allocatedTips',
      'StateWagesTipsEtc': 'stateWages',
      'StateIncomeTax': 'stateTaxWithheld',
      'LocalWagesTipsEtc': 'localWages',
      'LocalIncomeTax': 'localTaxWithheld'
    };
    
    for (const [azureFieldName, mappedFieldName] of Object.entries(w2FieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        w2Data[mappedFieldName] = typeof value === 'number' ? value : this.parseAmount(value);
      }
    }
    
    // Enhanced personal info extraction with better fallback handling
    console.log('🔍 [Azure DI] Extracting personal information from W2...');
    
    // Employee Name - try multiple field variations
    if (!w2Data.employeeName) {
      const nameFields = ['Employee.Name', 'EmployeeName', 'Employee_Name', 'RecipientName'];
      for (const fieldName of nameFields) {
        if (fields[fieldName]?.value) {
          w2Data.employeeName = fields[fieldName].value;
          console.log('✅ [Azure DI] Found employee name:', w2Data.employeeName);
          break;
        }
      }
    }
    
    // Employee SSN - try multiple field variations
    if (!w2Data.employeeSSN) {
      const ssnFields = ['Employee.SSN', 'EmployeeSSN', 'Employee_SSN', 'RecipientTIN'];
      for (const fieldName of ssnFields) {
        if (fields[fieldName]?.value) {
          w2Data.employeeSSN = fields[fieldName].value;
          console.log('✅ [Azure DI] Found employee SSN:', w2Data.employeeSSN);
          break;
        }
      }
    }
    
    // Employee Address - try multiple field variations
    if (!w2Data.employeeAddress) {
      const addressFields = ['Employee.Address', 'EmployeeAddress', 'Employee_Address', 'RecipientAddress'];
      for (const fieldName of addressFields) {
        if (fields[fieldName]?.value) {
          w2Data.employeeAddress = fields[fieldName].value;
          console.log('✅ [Azure DI] Found employee address:', w2Data.employeeAddress);
          break;
        }
      }
    }
    
    // OCR fallback for personal info if not found in structured fields
    if ((!w2Data.employeeName || !w2Data.employeeSSN || !w2Data.employeeAddress || !w2Data.employerName || !w2Data.employerAddress) && baseData.fullText) {
      console.log('🔍 [Azure DI] Some personal info missing from structured fields, attempting OCR extraction...');
      
      // Pass the already extracted employee name as a target for multi-employee scenarios
      const targetEmployeeName = w2Data.employeeName as string | undefined;
      const personalInfoFromOCR = this.extractPersonalInfoFromOCR(baseData.fullText as string, targetEmployeeName);
      
      if (!w2Data.employeeName && personalInfoFromOCR.name) {
        w2Data.employeeName = personalInfoFromOCR.name;
        console.log('✅ [Azure DI] Extracted employee name from OCR:', w2Data.employeeName);
      }
      
      if (!w2Data.employeeSSN && personalInfoFromOCR.ssn) {
        w2Data.employeeSSN = personalInfoFromOCR.ssn;
        console.log('✅ [Azure DI] Extracted employee SSN from OCR:', w2Data.employeeSSN);
      }
      
      if (!w2Data.employeeAddress && personalInfoFromOCR.address) {
        w2Data.employeeAddress = personalInfoFromOCR.address;
        console.log('✅ [Azure DI] Extracted employee address from OCR:', w2Data.employeeAddress);
      }
      
      if (!w2Data.employerName && personalInfoFromOCR.employerName) {
        w2Data.employerName = personalInfoFromOCR.employerName;
        console.log('✅ [Azure DI] Extracted employer name from OCR:', w2Data.employerName);
      }
      
      if (!w2Data.employerAddress && personalInfoFromOCR.employerAddress) {
        w2Data.employerAddress = personalInfoFromOCR.employerAddress;
        console.log('✅ [Azure DI] Extracted employer address from OCR:', w2Data.employerAddress);
      }
    }

    // Enhanced address parsing - extract city, state, and zipCode from full address
    if (w2Data.employeeAddress && typeof w2Data.employeeAddress === 'string') {
      console.log('🔍 [Azure DI] Parsing address components from:', w2Data.employeeAddress);
      const ocrText = typeof baseData.fullText === 'string' ? baseData.fullText : '';
      const addressParts = this.extractAddressParts(w2Data.employeeAddress, ocrText);
      
      // Add parsed address components to W2 data
      w2Data.employeeAddressStreet = addressParts.street;
      w2Data.employeeCity = addressParts.city;
      w2Data.employeeState = addressParts.state;
      w2Data.employeeZipCode = addressParts.zipCode;
      
      console.log('✅ [Azure DI] Parsed address components:', {
        street: w2Data.employeeAddressStreet,
        city: w2Data.employeeCity,
        state: w2Data.employeeState,
        zipCode: w2Data.employeeZipCode
      });
    }
    
    // OCR fallback for Box 1 wages if not found in structured fields
    if (!w2Data.wages && baseData.fullText) {
      console.log('🔍 [Azure DI] Wages not found in structured fields, attempting OCR extraction...');
      const wagesFromOCR = this.extractWagesFromOCR(baseData.fullText as string);
      if (wagesFromOCR > 0) {
        console.log('✅ [Azure DI] Successfully extracted wages from OCR:', wagesFromOCR);
        w2Data.wages = wagesFromOCR;
      }
    }
    
    return w2Data;
  }

  private process1099IntFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    const data = { ...baseData };
    
    // Comprehensive field mappings for all 1099-INT boxes
    const fieldMappings = {
      // Payer and recipient information
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'AccountNumber': 'accountNumber',
      
      // Box 1-17 mappings for 1099-INT
      'InterestIncome': 'interestIncome',                                    // Box 1
      'EarlyWithdrawalPenalty': 'earlyWithdrawalPenalty',                   // Box 2
      'InterestOnUSTreasuryObligations': 'interestOnUSavingsBonds',         // Box 3
      'InterestOnUSavingsBonds': 'interestOnUSavingsBonds',                 // Box 3 (alternative)
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',                     // Box 4
      'InvestmentExpenses': 'investmentExpenses',                           // Box 5
      'ForeignTaxPaid': 'foreignTaxPaid',                                   // Box 6
      'TaxExemptInterest': 'taxExemptInterest',                            // Box 8
      'SpecifiedPrivateActivityBondInterest': 'specifiedPrivateActivityBondInterest', // Box 9
      'MarketDiscount': 'marketDiscount',                                   // Box 10
      'BondPremium': 'bondPremium',                                         // Box 11
      'BondPremiumOnTreasuryObligations': 'bondPremiumOnTreasuryObligations', // Box 12
      'BondPremiumOnTaxExemptBond': 'bondPremiumOnTaxExemptBond',          // Box 13
      'TaxExemptAndTaxCreditBondCUSIPNo': 'taxExemptAndTaxCreditBondCUSIPNo', // Box 14
      'State': 'state',                                                     // Box 15
      'StateIdentificationNo': 'stateIdentificationNo',                     // Box 16
      'StateTaxWithheld': 'stateTaxWithheld',                              // Box 17
      
      // Alternative field names that Azure might use
      'Box1': 'interestIncome',
      'Box2': 'earlyWithdrawalPenalty',
      'Box3': 'interestOnUSavingsBonds',
      'Box4': 'federalTaxWithheld',
      'Box5': 'investmentExpenses',
      'Box6': 'foreignTaxPaid',
      'Box8': 'taxExemptInterest',
      'Box9': 'specifiedPrivateActivityBondInterest',
      'Box10': 'marketDiscount',
      'Box11': 'bondPremium',
      'Box12': 'bondPremiumOnTreasuryObligations',
      'Box13': 'bondPremiumOnTaxExemptBond',
      'Box14': 'taxExemptAndTaxCreditBondCUSIPNo',
      'Box15': 'state',
      'Box16': 'stateIdentificationNo',
      'Box17': 'stateTaxWithheld'
    };
    
    for (const [azureFieldName, mappedFieldName] of Object.entries(fieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        
        // Handle text fields vs numeric fields
        if (mappedFieldName === 'accountNumber' || 
            mappedFieldName === 'state' || 
            mappedFieldName === 'stateIdentificationNo' ||
            mappedFieldName === 'taxExemptAndTaxCreditBondCUSIPNo') {
          data[mappedFieldName] = String(value).trim();
        } else {
          data[mappedFieldName] = typeof value === 'number' ? value : this.parseAmount(value);
        }
      }
    }
    
    // OCR fallback for personal info if not found in structured fields
    if ((!data.recipientName || !data.recipientTIN || !data.recipientAddress || !data.payerName || !data.payerTIN) && baseData.fullText) {
      console.log('🔍 [Azure DI] Some 1099-INT info missing from structured fields, attempting OCR extraction...');
      const personalInfoFromOCR = this.extractPersonalInfoFromOCR(baseData.fullText as string);
      
      if (!data.recipientName && personalInfoFromOCR.name) {
        data.recipientName = personalInfoFromOCR.name;
        console.log('✅ [Azure DI] Extracted recipient name from OCR:', data.recipientName);
      }
      
      if (!data.recipientTIN && personalInfoFromOCR.tin) {
        data.recipientTIN = personalInfoFromOCR.tin;
        console.log('✅ [Azure DI] Extracted recipient TIN from OCR:', data.recipientTIN);
      }
      
      if (!data.recipientAddress && personalInfoFromOCR.address) {
        data.recipientAddress = personalInfoFromOCR.address;
        console.log('✅ [Azure DI] Extracted recipient address from OCR:', data.recipientAddress);
      }
      
      if (!data.payerName && personalInfoFromOCR.payerName) {
        data.payerName = personalInfoFromOCR.payerName;
        console.log('✅ [Azure DI] Extracted payer name from OCR:', data.payerName);
      }
      
      if (!data.payerTIN && personalInfoFromOCR.payerTIN) {
        data.payerTIN = personalInfoFromOCR.payerTIN;
        console.log('✅ [Azure DI] Extracted payer TIN from OCR:', data.payerTIN);
      }
    }
    
    // CRITICAL FIX: Add field validation and correction using OCR fallback
    if (baseData.fullText) {
      const validatedData = this.validateAndCorrect1099IntFields(data, baseData.fullText as string);
      return validatedData;
    }
    
    return data;
  }

  /**
   * Validates and corrects 1099-INT field mappings using OCR fallback
   * This addresses the issue where Azure DI maps values to incorrect fields
   */
  private validateAndCorrect1099IntFields(
    structuredData: ExtractedFieldData, 
    ocrText: string
  ): ExtractedFieldData {
    console.log('🔍 [Azure DI] Validating 1099-INT field mappings...');
    
    // Extract data using OCR as ground truth
    const ocrData = this.extract1099IntFieldsFromOCR(ocrText, { fullText: ocrText });
    
    const correctedData = { ...structuredData };
    let correctionsMade = 0;
    
    // Define validation rules for critical fields that commonly get mismatched
    const criticalFields = [
      'interestIncome',                    // Box 1 - Most important field
      'earlyWithdrawalPenalty',           // Box 2 - Often gets mapped incorrectly
      'interestOnUSavingsBonds',          // Box 3 - Often receives wrong values
      'federalTaxWithheld',               // Box 4 - Important for tax calculations
      'investmentExpenses',               // Box 5 - Sometimes misaligned
      'foreignTaxPaid',                   // Box 6 - Sometimes misaligned
      'taxExemptInterest',                // Box 8 - Often gets cross-contaminated
      'specifiedPrivateActivityBondInterest', // Box 9 - Complex field
      'marketDiscount',                   // Box 10 - Often misaligned
      'bondPremium'                       // Box 11 - Often misaligned
    ];
    
    for (const field of criticalFields) {
      const structuredValue = this.parseAmount(structuredData[field]) || 0;
      const ocrValue = this.parseAmount(ocrData[field]) || 0;
      
      // If values differ significantly (more than $100), trust OCR
      if (Math.abs(structuredValue - ocrValue) > 100) {
        console.log(`🔧 [Azure DI] Correcting ${field}: $${structuredValue} → $${ocrValue} (OCR)`);
        correctedData[field] = ocrValue;
        correctionsMade++;
      }
      // If structured field is empty/null but OCR found a value, use OCR
      else if ((structuredValue === 0 || !structuredData[field]) && ocrValue > 0) {
        console.log(`🔧 [Azure DI] Filling missing ${field}: $0 → $${ocrValue} (OCR)`);
        correctedData[field] = ocrValue;
        correctionsMade++;
      }
    }
    
    // Special validation for common cross-contamination patterns in 1099-INT
    // Pattern 1: Interest Income value incorrectly mapped to Early Withdrawal Penalty
    if (structuredData.earlyWithdrawalPenalty && !structuredData.interestIncome && 
        ocrData.interestIncome && ocrData.earlyWithdrawalPenalty) {
      const structuredPenalty = this.parseAmount(structuredData.earlyWithdrawalPenalty);
      const ocrInterest = this.parseAmount(ocrData.interestIncome);
      const ocrPenalty = this.parseAmount(ocrData.earlyWithdrawalPenalty);
      
      // If structured penalty amount matches OCR interest amount, it's likely swapped
      if (Math.abs(structuredPenalty - ocrInterest) < 100 && ocrPenalty !== structuredPenalty) {
        console.log(`🔧 [Azure DI] Detected cross-contamination: Interest Income/Early Withdrawal Penalty swap`);
        correctedData.interestIncome = ocrInterest;
        correctedData.earlyWithdrawalPenalty = ocrPenalty;
        correctionsMade += 2;
      }
    }
    
    // Pattern 2: Values shifted between adjacent boxes
    const adjacentBoxPairs = [
      ['interestIncome', 'earlyWithdrawalPenalty'],
      ['earlyWithdrawalPenalty', 'interestOnUSavingsBonds'],
      ['interestOnUSavingsBonds', 'federalTaxWithheld'],
      ['federalTaxWithheld', 'investmentExpenses'],
      ['investmentExpenses', 'foreignTaxPaid'],
      ['foreignTaxPaid', 'taxExemptInterest'],
      ['taxExemptInterest', 'specifiedPrivateActivityBondInterest'],
      ['specifiedPrivateActivityBondInterest', 'marketDiscount'],
      ['marketDiscount', 'bondPremium']
    ];
    
    for (const [field1, field2] of adjacentBoxPairs) {
      const struct1 = this.parseAmount(structuredData[field1]) || 0;
      const struct2 = this.parseAmount(structuredData[field2]) || 0;
      const ocr1 = this.parseAmount(ocrData[field1]) || 0;
      const ocr2 = this.parseAmount(ocrData[field2]) || 0;
      
      // Check if values are swapped between adjacent fields
      if (struct1 > 0 && struct2 > 0 && ocr1 > 0 && ocr2 > 0) {
        if (Math.abs(struct1 - ocr2) < 100 && Math.abs(struct2 - ocr1) < 100) {
          console.log(`🔧 [Azure DI] Detected adjacent field swap: ${field1} ↔ ${field2}`);
          correctedData[field1] = ocr1;
          correctedData[field2] = ocr2;
          correctionsMade += 2;
        }
      }
    }
    
    if (correctionsMade > 0) {
      console.log(`✅ [Azure DI] Made ${correctionsMade} field corrections using OCR validation`);
      
      // Log the corrections for debugging
      console.log('🔍 [Azure DI] Field correction summary:');
      for (const field of criticalFields) {
        const originalValue = this.parseAmount(structuredData[field]) || 0;
        const correctedValue = this.parseAmount(correctedData[field]) || 0;
        if (originalValue !== correctedValue) {
          console.log(`  ${field}: $${originalValue} → $${correctedValue}`);
        }
      }
    } else {
      console.log('✅ [Azure DI] No field corrections needed - structured extraction appears accurate');
    }
    
    return correctedData;
  }

  private process1099DivFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    const data = { ...baseData };
    
    const fieldMappings = {
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'OrdinaryDividends': 'ordinaryDividends',
      'QualifiedDividends': 'qualifiedDividends',
      'TotalCapitalGainDistributions': 'totalCapitalGain',
      'NondividendDistributions': 'nondividendDistributions',
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',
      'Section199ADividends': 'section199ADividends'
    };
    
    for (const [azureFieldName, mappedFieldName] of Object.entries(fieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        data[mappedFieldName] = typeof value === 'number' ? value : this.parseAmount(value);
      }
    }
    
    // OCR fallback for personal info if not found in structured fields
    if ((!data.recipientName || !data.recipientTIN || !data.recipientAddress || !data.payerName || !data.payerTIN) && baseData.fullText) {
      console.log('🔍 [Azure DI] Some 1099-DIV info missing from structured fields, attempting OCR extraction...');
      const personalInfoFromOCR = this.extractPersonalInfoFromOCR(baseData.fullText as string);
      
      if (!data.recipientName && personalInfoFromOCR.name) {
        data.recipientName = personalInfoFromOCR.name;
        console.log('✅ [Azure DI] Extracted recipient name from OCR:', data.recipientName);
      }
      
      if (!data.recipientTIN && personalInfoFromOCR.tin) {
        data.recipientTIN = personalInfoFromOCR.tin;
        console.log('✅ [Azure DI] Extracted recipient TIN from OCR:', data.recipientTIN);
      }
      
      if (!data.recipientAddress && personalInfoFromOCR.address) {
        data.recipientAddress = personalInfoFromOCR.address;
        console.log('✅ [Azure DI] Extracted recipient address from OCR:', data.recipientAddress);
      }
      
      if (!data.payerName && personalInfoFromOCR.payerName) {
        data.payerName = personalInfoFromOCR.payerName;
        console.log('✅ [Azure DI] Extracted payer name from OCR:', data.payerName);
      }
      
      if (!data.payerTIN && personalInfoFromOCR.payerTIN) {
        data.payerTIN = personalInfoFromOCR.payerTIN;
        console.log('✅ [Azure DI] Extracted payer TIN from OCR:', data.payerTIN);
      }
    }
    
    return data;
  }

  private process1099MiscFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    const data = { ...baseData };
    
    // Comprehensive field mappings for all 1099-MISC boxes
    const fieldMappings = {
      // Payer and recipient information
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'AccountNumber': 'accountNumber',
      
      // Box 1-18 mappings
      'Rents': 'rents',                                           // Box 1
      'Royalties': 'royalties',                                   // Box 2
      'OtherIncome': 'otherIncome',                              // Box 3
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',          // Box 4
      'FishingBoatProceeds': 'fishingBoatProceeds',              // Box 5
      'MedicalAndHealthCarePayments': 'medicalHealthPayments',    // Box 6
      'NonemployeeCompensation': 'nonemployeeCompensation',       // Box 7 (deprecated)
      'SubstitutePayments': 'substitutePayments',                 // Box 8
      'CropInsuranceProceeds': 'cropInsuranceProceeds',          // Box 9
      'GrossProceedsPaidToAttorney': 'grossProceedsAttorney',         // Box 10
      'FishPurchasedForResale': 'fishPurchases',                 // Box 11
      'Section409ADeferrals': 'section409ADeferrals',            // Box 12
      'ExcessGoldenParachutePayments': 'excessGoldenParachutePayments', // Box 13
      'NonqualifiedDeferredCompensation': 'nonqualifiedDeferredCompensation', // Box 14
      'Section409AIncome': 'section409AIncome',                  // Box 15a
      'StateTaxWithheld': 'stateTaxWithheld',                    // Box 16
      'StatePayerNumber': 'statePayerNumber',                    // Box 17
      'StateIncome': 'stateIncome',                              // Box 18
      
      // Alternative field names that Azure might use
      'Box1': 'rents',
      'Box2': 'royalties',
      'Box3': 'otherIncome',
      'Box4': 'federalTaxWithheld',
      'Box5': 'fishingBoatProceeds',
      'Box6': 'medicalHealthPayments',
      'Box7': 'nonemployeeCompensation',
      'Box8': 'substitutePayments',
      'Box9': 'cropInsuranceProceeds',
      'Box10': 'grossProceedsAttorney',
      'Box11': 'fishPurchases',
      'Box12': 'section409ADeferrals',
      'Box13': 'excessGoldenParachutePayments',
      'Box14': 'nonqualifiedDeferredCompensation',
      'Box15a': 'section409AIncome',
      'Box16': 'stateTaxWithheld',
      'Box17': 'statePayerNumber',
      'Box18': 'stateIncome'
    };
    
    for (const [azureFieldName, mappedFieldName] of Object.entries(fieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        
        // Handle text fields vs numeric fields
        if (mappedFieldName === 'statePayerNumber' || mappedFieldName === 'accountNumber') {
          data[mappedFieldName] = String(value).trim();
        } else {
          data[mappedFieldName] = typeof value === 'number' ? value : this.parseAmount(value);
        }
      }
    }
    
    // OCR fallback for personal info if not found in structured fields
    if ((!data.recipientName || !data.recipientTIN || !data.recipientAddress || !data.payerName || !data.payerTIN) && baseData.fullText) {
      console.log('🔍 [Azure DI] Some 1099-MISC info missing from structured fields, attempting OCR extraction...');
      const personalInfoFromOCR = this.extractPersonalInfoFromOCR(baseData.fullText as string);
      
      if (!data.recipientName && personalInfoFromOCR.name) {
        data.recipientName = personalInfoFromOCR.name;
        console.log('✅ [Azure DI] Extracted recipient name from OCR:', data.recipientName);
      }
      
      if (!data.recipientTIN && personalInfoFromOCR.tin) {
        data.recipientTIN = personalInfoFromOCR.tin;
        console.log('✅ [Azure DI] Extracted recipient TIN from OCR:', data.recipientTIN);
      }
      
      if (!data.recipientAddress && personalInfoFromOCR.address) {
        data.recipientAddress = personalInfoFromOCR.address;
        console.log('✅ [Azure DI] Extracted recipient address from OCR:', data.recipientAddress);
      }
      
      if (!data.payerName && personalInfoFromOCR.payerName) {
        data.payerName = personalInfoFromOCR.payerName;
        console.log('✅ [Azure DI] Extracted payer name from OCR:', data.payerName);
      }
      
      if (!data.payerTIN && personalInfoFromOCR.payerTIN) {
        data.payerTIN = personalInfoFromOCR.payerTIN;
        console.log('✅ [Azure DI] Extracted payer TIN from OCR:', data.payerTIN);
      }
    }
    
    // CRITICAL FIX: Add field validation and correction using OCR fallback
    if (baseData.fullText) {
      const validatedData = this.validateAndCorrect1099MiscFields(data, baseData.fullText as string);
      return validatedData;
    }
    
    return data;
  }

  /**
   * Validates and corrects 1099-MISC field mappings using OCR fallback
   * This addresses the issue where Azure DI maps values to incorrect fields
   */
  private validateAndCorrect1099MiscFields(
    structuredData: ExtractedFieldData, 
    ocrText: string
  ): ExtractedFieldData {
    console.log('🔍 [Azure DI] Validating 1099-MISC field mappings...');
    
    // Extract data using OCR as ground truth
    const ocrData = this.extract1099MiscFieldsFromOCR(ocrText, { fullText: ocrText });
    
    const correctedData = { ...structuredData };
    let correctionsMade = 0;
    
    // Define validation rules for critical fields that commonly get mismatched
    const criticalFields = [
      'otherIncome',           // Box 3 - Often gets mapped incorrectly
      'fishingBoatProceeds',   // Box 5 - Often receives wrong values
      'medicalHealthPayments', // Box 6 - Often gets cross-contaminated
      'rents',                 // Box 1 - Sometimes misaligned
      'royalties',             // Box 2 - Sometimes misaligned
      'federalTaxWithheld'     // Box 4 - Important for tax calculations
    ];
    
    for (const field of criticalFields) {
      const structuredValue = this.parseAmount(structuredData[field]) || 0;
      const ocrValue = this.parseAmount(ocrData[field]) || 0;
      
      // If values differ significantly (more than $100), trust OCR
      if (Math.abs(structuredValue - ocrValue) > 100) {
        console.log(`🔧 [Azure DI] Correcting ${field}: $${structuredValue} → $${ocrValue} (OCR)`);
        correctedData[field] = ocrValue;
        correctionsMade++;
      }
      // If structured field is empty/null but OCR found a value, use OCR
      else if ((structuredValue === 0 || !structuredData[field]) && ocrValue > 0) {
        console.log(`🔧 [Azure DI] Filling missing ${field}: $0 → $${ocrValue} (OCR)`);
        correctedData[field] = ocrValue;
        correctionsMade++;
      }
    }
    
    // Special validation for common cross-contamination patterns
    // Pattern 1: Other Income value incorrectly mapped to Fishing Boat Proceeds
    if (structuredData.fishingBoatProceeds && !structuredData.otherIncome && 
        ocrData.otherIncome && ocrData.fishingBoatProceeds) {
      const structuredFishing = this.parseAmount(structuredData.fishingBoatProceeds);
      const ocrOther = this.parseAmount(ocrData.otherIncome);
      const ocrFishing = this.parseAmount(ocrData.fishingBoatProceeds);
      
      // If structured fishing amount matches OCR other income amount, it's likely swapped
      if (Math.abs(structuredFishing - ocrOther) < 100 && ocrFishing !== structuredFishing) {
        console.log(`🔧 [Azure DI] Detected cross-contamination: Other Income/Fishing Boat Proceeds swap`);
        correctedData.otherIncome = ocrOther;
        correctedData.fishingBoatProceeds = ocrFishing;
        correctionsMade += 2;
      }
    }
    
    // Pattern 2: Values shifted between adjacent boxes
    const adjacentBoxPairs = [
      ['rents', 'royalties'],
      ['royalties', 'otherIncome'],
      ['otherIncome', 'federalTaxWithheld'],
      ['federalTaxWithheld', 'fishingBoatProceeds'],
      ['fishingBoatProceeds', 'medicalHealthPayments']
    ];
    
    for (const [field1, field2] of adjacentBoxPairs) {
      const struct1 = this.parseAmount(structuredData[field1]) || 0;
      const struct2 = this.parseAmount(structuredData[field2]) || 0;
      const ocr1 = this.parseAmount(ocrData[field1]) || 0;
      const ocr2 = this.parseAmount(ocrData[field2]) || 0;
      
      // Check if values are swapped between adjacent fields
      if (struct1 > 0 && struct2 > 0 && ocr1 > 0 && ocr2 > 0) {
        if (Math.abs(struct1 - ocr2) < 100 && Math.abs(struct2 - ocr1) < 100) {
          console.log(`🔧 [Azure DI] Detected adjacent field swap: ${field1} ↔ ${field2}`);
          correctedData[field1] = ocr1;
          correctedData[field2] = ocr2;
          correctionsMade += 2;
        }
      }
    }
    
    if (correctionsMade > 0) {
      console.log(`✅ [Azure DI] Made ${correctionsMade} field corrections using OCR validation`);
      
      // Log the corrections for debugging
      console.log('🔍 [Azure DI] Field correction summary:');
      for (const field of criticalFields) {
        const originalValue = this.parseAmount(structuredData[field]) || 0;
        const correctedValue = this.parseAmount(correctedData[field]) || 0;
        if (originalValue !== correctedValue) {
          console.log(`  ${field}: $${originalValue} → $${correctedValue}`);
        }
      }
    } else {
      console.log('✅ [Azure DI] No field corrections needed - structured extraction appears accurate');
    }
    
    return correctedData;
  }

  private process1099NecFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    const data = { ...baseData };
    
    const fieldMappings = {
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'NonemployeeCompensation': 'nonemployeeCompensation',
      'FederalIncomeTaxWithheld': 'federalTaxWithheld'
    };
    
    for (const [azureFieldName, mappedFieldName] of Object.entries(fieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        data[mappedFieldName] = typeof value === 'number' ? value : this.parseAmount(value);
      }
    }
    
    // OCR fallback for personal info if not found in structured fields
    if ((!data.recipientName || !data.recipientTIN || !data.recipientAddress || !data.payerName || !data.payerTIN) && baseData.fullText) {
      console.log('🔍 [Azure DI] Some 1099-NEC info missing from structured fields, attempting OCR extraction...');
      const personalInfoFromOCR = this.extractPersonalInfoFromOCR(baseData.fullText as string);
      
      if (!data.recipientName && personalInfoFromOCR.name) {
        data.recipientName = personalInfoFromOCR.name;
        console.log('✅ [Azure DI] Extracted recipient name from OCR:', data.recipientName);
      }
      
      if (!data.recipientTIN && personalInfoFromOCR.tin) {
        data.recipientTIN = personalInfoFromOCR.tin;
        console.log('✅ [Azure DI] Extracted recipient TIN from OCR:', data.recipientTIN);
      }
      
      if (!data.recipientAddress && personalInfoFromOCR.address) {
        data.recipientAddress = personalInfoFromOCR.address;
        console.log('✅ [Azure DI] Extracted recipient address from OCR:', data.recipientAddress);
      }
      
      if (!data.payerName && personalInfoFromOCR.payerName) {
        data.payerName = personalInfoFromOCR.payerName;
        console.log('✅ [Azure DI] Extracted payer name from OCR:', data.payerName);
      }
      
      if (!data.payerTIN && personalInfoFromOCR.payerTIN) {
        data.payerTIN = personalInfoFromOCR.payerTIN;
        console.log('✅ [Azure DI] Extracted payer TIN from OCR:', data.payerTIN);
      }
    }
    
    return data;
  }

  private processGenericFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    const data = { ...baseData };
    
    // Process all available fields
    for (const [fieldName, fieldData] of Object.entries(fields)) {
      if (fieldData && typeof fieldData === 'object' && 'value' in fieldData) {
        const value = (fieldData as any).value;
        if (value !== undefined && value !== null && value !== '') {
          data[fieldName] = typeof value === 'number' ? value : this.parseAmount(value);
        }
      }
    }
    
    return data;
  }

  public analyzeDocumentTypeFromOCR(ocrText: string): string {
    console.log('🔍 [Azure DI] Analyzing document type from OCR content...');
    
    const formType = this.detectFormType(ocrText);
    
    if (formType === 'W2') {
      console.log('✅ [Azure DI] Confirmed W2 document type');
      return 'W2';
    } else if (formType === '1099') {
      const specific1099Type = this.detectSpecific1099Type(ocrText);
      console.log(`✅ [Azure DI] Detected specific 1099 type: ${specific1099Type}`);
      return specific1099Type;
    }
    
    console.log('⚠️ [Azure DI] Could not determine document type from OCR');
    return 'UNKNOWN';
  }

  public detectSpecific1099Type(ocrText: string): string {
    console.log('🔍 [Azure DI] Detecting specific 1099 subtype from OCR text...');
    
    const text = ocrText.toLowerCase();
    
    // Check for specific 1099 form types with high-confidence indicators
    const formTypePatterns = [
      {
        type: 'FORM_1099_DIV',
        indicators: [
          'form 1099-div',
          'dividends and distributions',
          'ordinary dividends',
          'qualified dividends',
          'total capital gain distributions',
          'capital gain distributions'
        ]
      },
      {
        type: 'FORM_1099_INT',
        indicators: [
          'form 1099-int',
          'interest income',
          'early withdrawal penalty',
          'interest on u.s. treasury obligations',
          'investment expenses',
          'interest on u.s. savings bonds',
          'tax-exempt interest'
        ]
      },
      {
        type: 'FORM_1099_MISC',
        indicators: [
          'form 1099-misc',
          'miscellaneous income',
          'nonemployee compensation',
          'rents',
          'royalties',
          'fishing boat proceeds'
        ]
      },
      {
        type: 'FORM_1099_NEC',
        indicators: [
          'form 1099-nec',
          'nonemployee compensation',
          'nec'
        ]
      }
    ];
    
    // Score each form type based on indicator matches
    let bestMatch = { type: 'FORM_1099_MISC', score: 0 }; // Default to MISC
    
    for (const formPattern of formTypePatterns) {
      let score = 0;
      for (const indicator of formPattern.indicators) {
        if (text.includes(indicator)) {
          score += 1;
          console.log(`✅ [Azure DI] Found indicator "${indicator}" for ${formPattern.type}`);
        }
      }
      
      if (score > bestMatch.score) {
        bestMatch = { type: formPattern.type, score };
      }
    }
    
    console.log(`✅ [Azure DI] Best match: ${bestMatch.type} (score: ${bestMatch.score})`);
    return bestMatch.type;
  }

  private detectFormType(ocrText: string): string {
    const text = ocrText.toLowerCase();
    
    // W2 indicators
    const w2Indicators = [
      'form w-2',
      'wage and tax statement',
      'wages, tips, other compensation',
      'federal income tax withheld',
      'social security wages',
      'medicare wages'
    ];
    
    // 1099 indicators
    const form1099Indicators = [
      'form 1099',
      '1099-',
      'payer',
      'recipient',
      'tin'
    ];
    
    // Count matches for each form type
    let w2Score = 0;
    let form1099Score = 0;
    
    for (const indicator of w2Indicators) {
      if (text.includes(indicator)) {
        w2Score++;
      }
    }
    
    for (const indicator of form1099Indicators) {
      if (text.includes(indicator)) {
        form1099Score++;
      }
    }
    
    console.log(`🔍 [Azure DI] Form type scores - W2: ${w2Score}, 1099: ${form1099Score}`);
    
    if (w2Score > form1099Score) {
      return 'W2';
    } else if (form1099Score > 0) {
      return '1099';
    }
    
    return 'UNKNOWN';
  }

  // === 1099-INT OCR EXTRACTION ===
  /**
   * Extracts 1099-INT fields from OCR text using comprehensive regex patterns
   * Handles all 17 boxes plus personal information fields
   */
  private extract1099IntFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    console.log('🔍 [Azure DI OCR] Extracting 1099-INT fields from OCR text...');
    
    const data = { ...baseData };
    
    // Extract personal information first
    const personalInfo = this.extract1099InfoFromOCR(ocrText);
    
    // Add personal info to data
    if (personalInfo.name) data.recipientName = personalInfo.name;
    if (personalInfo.tin) data.recipientTIN = personalInfo.tin;
    if (personalInfo.address) data.recipientAddress = personalInfo.address;
    if (personalInfo.payerName) data.payerName = personalInfo.payerName;
    if (personalInfo.payerTIN) data.payerTIN = personalInfo.payerTIN;
    if (personalInfo.payerAddress) data.payerAddress = personalInfo.payerAddress;
    
    // Extract account number
    const accountNumberMatch = ocrText.match(/Account number[:\s]*([A-Za-z0-9\-]+)/i);
    if (accountNumberMatch && accountNumberMatch[1]) {
      data.accountNumber = accountNumberMatch[1].trim();
      console.log('✅ [Azure DI OCR] Found account number:', data.accountNumber);
    }
    
    // === BOX EXTRACTION PATTERNS ===
    
    // Box 1: Interest Income
    const box1Patterns = [
      /(?:1\s*Interest income|Box 1|Interest income)[:\s]*\$?([0-9,]+\.?\d*)/i,
      /Interest income[:\s]*\$?([0-9,]+\.?\d*)/i
    ];
    
    for (const pattern of box1Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.interestIncome = amount;
          console.log('✅ [Azure DI OCR] Found Box 1 - Interest Income:', amount);
          break;
        }
      }
    }
    
    // Box 2: Early Withdrawal Penalty
    const box2Patterns = [
      /(?:2\s*Early withdrawal penalty|Box 2|Early withdrawal penalty)[:\s]*\$?([0-9,]+\.?\d*)/i,
      /Early withdrawal penalty[:\s]*\$?([0-9,]+\.?\d*)/i
    ];
    
    for (const pattern of box2Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.earlyWithdrawalPenalty = amount;
          console.log('✅ [Azure DI OCR] Found Box 2 - Early Withdrawal Penalty:', amount);
          break;
        }
      }
    }
    
    // Box 3: Interest on U.S. Savings Bonds and Treasury Obligations
    const box3Patterns = [
      /(?:3\s*Interest on U\.S\. Savings Bonds|Box 3|Interest on U\.S\. Treasury obligations)[:\s]*\$?([0-9,]+\.?\d*)/i,
      /Interest on U\.S\. Savings Bonds[:\s]*\$?([0-9,]+\.?\d*)/i,
      /Interest on U\.S\. Treasury obligations[:\s]*\$?([0-9,]+\.?\d*)/i
    ];
    
    for (const pattern of box3Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.interestOnUSavingsBonds = amount;
          console.log('✅ [Azure DI OCR] Found Box 3 - Interest on U.S. Savings Bonds:', amount);
          break;
        }
      }
    }
    
    // Box 4: Federal Income Tax Withheld
    const box4Patterns = [
      /(?:4\s*Federal income tax withheld|Box 4|Federal income tax withheld)[:\s]*\$?([0-9,]+\.?\d*)/i,
      /Federal income tax withheld[:\s]*\$?([0-9,]+\.?\d*)/i
    ];
    
    for (const pattern of box4Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.federalTaxWithheld = amount;
          console.log('✅ [Azure DI OCR] Found Box 4 - Federal Tax Withheld:', amount);
          break;
        }
      }
    }
    
    // Box 5: Investment Expenses
    const box5Patterns = [
      /(?:5\s*Investment expenses|Box 5|Investment expenses)[:\s]*\$?([0-9,]+\.?\d*)/i,
      /Investment expenses[:\s]*\$?([0-9,]+\.?\d*)/i
    ];
    
    for (const pattern of box5Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.investmentExpenses = amount;
          console.log('✅ [Azure DI OCR] Found Box 5 - Investment Expenses:', amount);
          break;
        }
      }
    }
    
    // Box 6: Foreign Tax Paid
    const box6Patterns = [
      /(?:6\s*Foreign tax paid|Box 6|Foreign tax paid)[:\s]*\$?([0-9,]+\.?\d*)/i,
      /Foreign tax paid[:\s]*\$?([0-9,]+\.?\d*)/i
    ];
    
    for (const pattern of box6Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.foreignTaxPaid = amount;
          console.log('✅ [Azure DI OCR] Found Box 6 - Foreign Tax Paid:', amount);
          break;
        }
      }
    }
    
    // Box 8: Tax-Exempt Interest
    const box8Patterns = [
      /(?:8\s*Tax-exempt interest|Box 8|Tax-exempt interest)[:\s]*\$?([0-9,]+\.?\d*)/i,
      /Tax-exempt interest[:\s]*\$?([0-9,]+\.?\d*)/i
    ];
    
    for (const pattern of box8Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.taxExemptInterest = amount;
          console.log('✅ [Azure DI OCR] Found Box 8 - Tax-Exempt Interest:', amount);
          break;
        }
      }
    }
    
    // Box 9: Specified Private Activity Bond Interest
    const box9Patterns = [
      /(?:9\s*Specified private activity bond interest|Box 9|Specified private activity bond interest)[:\s]*\$?([0-9,]+\.?\d*)/i,
      /Specified private activity bond interest[:\s]*\$?([0-9,]+\.?\d*)/i
    ];
    
    for (const pattern of box9Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.specifiedPrivateActivityBondInterest = amount;
          console.log('✅ [Azure DI OCR] Found Box 9 - Specified Private Activity Bond Interest:', amount);
          break;
        }
      }
    }
    
    // Box 10: Market Discount
    const box10Patterns = [
      /(?:10\s*Market discount|Box 10|Market discount)[:\s]*\$?([0-9,]+\.?\d*)/i,
      /Market discount[:\s]*\$?([0-9,]+\.?\d*)/i
    ];
    
    for (const pattern of box10Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.marketDiscount = amount;
          console.log('✅ [Azure DI OCR] Found Box 10 - Market Discount:', amount);
          break;
        }
      }
    }
    
    // Box 11: Bond Premium
    const box11Patterns = [
      /(?:11\s*Bond premium|Box 11|Bond premium)[:\s]*\$?([0-9,]+\.?\d*)/i,
      /Bond premium[:\s]*\$?([0-9,]+\.?\d*)/i
    ];
    
    for (const pattern of box11Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.bondPremium = amount;
          console.log('✅ [Azure DI OCR] Found Box 11 - Bond Premium:', amount);
          break;
        }
      }
    }
    
    // Box 12: Bond Premium on Treasury Obligations
    const box12Patterns = [
      /(?:12\s*Bond premium on Treasury obligations|Box 12|Bond premium on Treasury obligations)[:\s]*\$?([0-9,]+\.?\d*)/i,
      /Bond premium on Treasury obligations[:\s]*\$?([0-9,]+\.?\d*)/i
    ];
    
    for (const pattern of box12Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.bondPremiumOnTreasuryObligations = amount;
          console.log('✅ [Azure DI OCR] Found Box 12 - Bond Premium on Treasury Obligations:', amount);
          break;
        }
      }
    }
    
    // Box 13: Bond Premium on Tax-Exempt Bond
    const box13Patterns = [
      /(?:13\s*Bond premium on tax-exempt bond|Box 13|Bond premium on tax-exempt bond)[:\s]*\$?([0-9,]+\.?\d*)/i,
      /Bond premium on tax-exempt bond[:\s]*\$?([0-9,]+\.?\d*)/i
    ];
    
    for (const pattern of box13Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.bondPremiumOnTaxExemptBond = amount;
          console.log('✅ [Azure DI OCR] Found Box 13 - Bond Premium on Tax-Exempt Bond:', amount);
          break;
        }
      }
    }
    
    // Box 14: Tax-Exempt and Tax Credit Bond CUSIP No.
    const box14Patterns = [
      /(?:14\s*Tax-exempt and tax credit bond CUSIP no\.|Box 14|Tax-exempt and tax credit bond CUSIP)[:\s]*([A-Za-z0-9]+)/i,
      /Tax-exempt and tax credit bond CUSIP[:\s]*([A-Za-z0-9]+)/i
    ];
    
    for (const pattern of box14Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const cusip = match[1].trim();
        if (cusip.length > 0) {
          data.taxExemptAndTaxCreditBondCUSIPNo = cusip;
          console.log('✅ [Azure DI OCR] Found Box 14 - CUSIP No.:', cusip);
          break;
        }
      }
    }
    
    // Box 15: State
    const box15Patterns = [
      /(?:15\s*State|Box 15)[:\s]*([A-Z]{2})/i,
      /State[:\s]*([A-Z]{2})/i
    ];
    
    for (const pattern of box15Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const state = match[1].trim();
        if (state.length === 2) {
          data.state = state;
          console.log('✅ [Azure DI OCR] Found Box 15 - State:', state);
          break;
        }
      }
    }
    
    // Box 16: State Identification No.
    const box16Patterns = [
      /(?:16\s*State identification no\.|Box 16|State identification)[:\s]*([A-Za-z0-9\-]+)/i,
      /State identification[:\s]*([A-Za-z0-9\-]+)/i
    ];
    
    for (const pattern of box16Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const stateId = match[1].trim();
        if (stateId.length > 0) {
          data.stateIdentificationNo = stateId;
          console.log('✅ [Azure DI OCR] Found Box 16 - State Identification No.:', stateId);
          break;
        }
      }
    }
    
    // Box 17: State Tax Withheld
    const box17Patterns = [
      /(?:17\s*State tax withheld|Box 17|State tax withheld)[:\s]*\$?([0-9,]+\.?\d*)/i,
      /State tax withheld[:\s]*\$?([0-9,]+\.?\d*)/i
    ];
    
    for (const pattern of box17Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.stateTaxWithheld = amount;
          console.log('✅ [Azure DI OCR] Found Box 17 - State Tax Withheld:', amount);
          break;
        }
      }
    }
    
    console.log('✅ [Azure DI OCR] Completed 1099-INT field extraction');
    return data;
  }

  // === 1099 PATTERNS ===
  /**
   * Extracts personal information from 1099 OCR text using comprehensive regex patterns
   * Specifically designed for 1099 form OCR text patterns with enhanced fallback mechanisms
   */
  private extract1099InfoFromOCR(ocrText: string): {
    name?: string;
    tin?: string;
    address?: string;
    payerName?: string;
    payerTIN?: string;
    payerAddress?: string;
  } {
    console.log('🔍 [Azure DI OCR] Searching for 1099 info in OCR text...');
    
    const info1099: { 
      name?: string; 
      tin?: string; 
      address?: string;
      payerName?: string;
      payerTIN?: string;
      payerAddress?: string;
    } = {};
    
    // === RECIPIENT NAME PATTERNS ===
    const recipientNamePatterns = [
      // RECIPIENT_NAME_MULTILINE: Extract name that appears after "RECIPIENT'S name" label
      {
        name: 'RECIPIENT_NAME_MULTILINE',
        pattern: /(?:RECIPIENT'S?\s+name|Recipient'?s?\s+name)\s*\n([A-Za-z\s]+?)(?:\n|$)/i,
        example: "RECIPIENT'S name\nJordan Blake"
      },
      // RECIPIENT_NAME_BASIC: Basic recipient name extraction
      {
        name: 'RECIPIENT_NAME_BASIC',
        pattern: /(?:RECIPIENT'S?\s+NAME|Recipient'?s?\s+name)[:\s]+([A-Za-z\s]+?)(?:\s+\d|\n|RECIPIENT'S?\s+|Recipient'?s?\s+|TIN|address|street|$)/i,
        example: "RECIPIENT'S NAME JOHN DOE"
      },
      {
        name: 'RECIPIENT_NAME_COLON',
        pattern: /(?:RECIPIENT'S?\s+name|Recipient'?s?\s+name):\s*([A-Za-z\s]+?)(?:\n|RECIPIENT'S?\s+|Recipient'?s?\s+|TIN|address|street|$)/i,
        example: "RECIPIENT'S name: JOHN DOE"
      }
    ];
    
    // Try recipient name patterns
    for (const patternInfo of recipientNamePatterns) {
      const match = ocrText.match(patternInfo.pattern);
      if (match && match[1]) {
        const name = match[1].trim();
        if (name.length > 2 && /^[A-Za-z\s]+$/.test(name)) {
          info1099.name = name;
          console.log(`✅ [Azure DI OCR] Found recipient name using ${patternInfo.name}:`, name);
          break;
        }
      }
    }
    
    // === RECIPIENT TIN PATTERNS ===
    const recipientTinPatterns = [
      {
        name: 'RECIPIENT_TIN_BASIC',
        pattern: /(?:RECIPIENT'S?\s+TIN|Recipient'?s?\s+TIN)[:\s]+(\d{2,3}[-\s]?\d{2}[-\s]?\d{4})/i,
        example: "RECIPIENT'S TIN 123-45-6789"
      },
      {
        name: 'RECIPIENT_TIN_MULTILINE',
        pattern: /(?:RECIPIENT'S?\s+TIN|Recipient'?s?\s+TIN)\s*\n(\d{2,3}[-\s]?\d{2}[-\s]?\d{4})/i,
        example: "RECIPIENT'S TIN\n123-45-6789"
      }
    ];
    
    // Try recipient TIN patterns
    for (const patternInfo of recipientTinPatterns) {
      const match = ocrText.match(patternInfo.pattern);
      if (match && match[1]) {
        const tin = match[1].trim();
        if (tin.length >= 9) {
          info1099.tin = tin;
          console.log(`✅ [Azure DI OCR] Found recipient TIN using ${patternInfo.name}:`, tin);
          break;
        }
      }
    }
    
    // === RECIPIENT ADDRESS PATTERNS ===
    const recipientAddressPatterns = [
      {
        name: 'RECIPIENT_ADDRESS_STREET_CITY_STRUCTURED',
        pattern: /Street address \(including apt\. no\.\)\s*\n([^\n]+)\s*\nCity or town, state or province, country, and ZIP or foreign postal code\s*\n([^\n]+)/i,
        example: "Street address (including apt. no.)\n456 MAIN STREET\nCity or town, state or province, country, and ZIP or foreign postal code\nHOMETOWN, ST 67890"
      },
      {
        name: 'RECIPIENT_ADDRESS_MULTILINE',
        pattern: /(?:RECIPIENT'S?\s+address|Recipient'?s?\s+address)\s*\n([^\n]+(?:\n[^\n]+)*?)(?:\n\s*\n|PAYER'S?\s+|Payer'?s?\s+|$)/i,
        example: "RECIPIENT'S address\n123 Main St\nAnytown, ST 12345"
      },
      {
        name: 'RECIPIENT_ADDRESS_BASIC',
        pattern: /(?:RECIPIENT'S?\s+address|Recipient'?s?\s+address)[:\s]+([^\n]+(?:\n[^\n]+)*?)(?:\n\s*\n|PAYER'S?\s+|Payer'?s?\s+|$)/i,
        example: "RECIPIENT'S address: 123 Main St, Anytown, ST 12345"
      },
      {
        name: 'RECIPIENT_ADDRESS_STREET_CITY_PRECISE',
        pattern: /RECIPIENT'S name\s*\n[^\n]+\s*\nStreet address[^\n]*\n([^\n]+)\s*\nCity[^\n]*\n([^\n]+)/i,
        example: "RECIPIENT'S name\nJordan Blake\nStreet address (including apt. no.)\n456 MAIN STREET\nCity or town, state or province, country, and ZIP or foreign postal code\nHOMETOWN, ST 67890"
      },
      {
        name: 'RECIPIENT_ADDRESS_AFTER_TIN',
        pattern: /RECIPIENT'S TIN:[^\n]*\n\s*\n([^\n]+)\s*\n([^\n]+)/i,
        example: "RECIPIENT'S TIN: XXX-XX-4567\n\n456 MAIN STREET\nHOMETOWN, ST 67890"
      },
      {
        name: 'RECIPIENT_ADDRESS_SIMPLE_AFTER_NAME',
        pattern: /RECIPIENT'S name\s*\n([^\n]+)\s*\n\s*([^\n]+)\s*\n\s*([^\n]+)/i,
        example: "RECIPIENT'S name\nJordan Blake\n456 MAIN STREET\nHOMETOWN, ST 67890"
      }
    ];
    
    // Try recipient address patterns
    for (const patternInfo of recipientAddressPatterns) {
      const match = ocrText.match(patternInfo.pattern);
      if (match && match[1]) {
        let address = '';
        
        // Handle patterns that capture street and city separately
        if (patternInfo.name === 'RECIPIENT_ADDRESS_STREET_CITY_STRUCTURED') {
          // match[1] is street, match[2] is city/state/zip
          if (match[2]) {
            address = `${match[1].trim()} ${match[2].trim()}`;
          } else {
            address = match[1].trim();
          }
        } else if (patternInfo.name === 'RECIPIENT_ADDRESS_STREET_CITY_PRECISE') {
          // match[1] is street, match[2] is city/state/zip
          if (match[2] && !match[2].toLowerCase().includes('city or town')) {
            address = `${match[1].trim()} ${match[2].trim()}`;
          } else {
            address = match[1].trim();
          }
        } else if (patternInfo.name === 'RECIPIENT_ADDRESS_AFTER_TIN') {
          // match[1] is street, match[2] is city/state/zip
          if (match[2]) {
            address = `${match[1].trim()} ${match[2].trim()}`;
          } else {
            address = match[1].trim();
          }
        } else if (patternInfo.name === 'RECIPIENT_ADDRESS_SIMPLE_AFTER_NAME') {
          // match[1] is name (skip), match[2] is street, match[3] is city/state/zip
          if (match[3] && match[2] && !match[2].toLowerCase().includes('street address')) {
            address = `${match[2].trim()} ${match[3].trim()}`;
          } else if (match[2] && !match[2].toLowerCase().includes('street address')) {
            address = match[2].trim();
          }
        } else {
          // For basic patterns, just use the captured text
          address = match[1].trim().replace(/\n+/g, ' ');
        }
        
        // Validate the address doesn't contain form labels
        if (address.length > 5 && 
            !address.toLowerCase().includes('street address') &&
            !address.toLowerCase().includes('including apt') &&
            !address.toLowerCase().includes('city or town')) {
          info1099.address = address;
          console.log(`✅ [Azure DI OCR] Found recipient address using ${patternInfo.name}:`, address);
          break;
        }
      }
    }
    
    // === PAYER NAME PATTERNS ===
    const payerNamePatterns = [
      {
        name: 'PAYER_NAME_AFTER_LABEL',
        pattern: /(?:PAYER'S?\s+name,\s+street\s+address[^\n]*\n)([A-Za-z\s&.,'-]+?)(?:\n|$)/i,
        example: "PAYER'S name, street address, city or town, state or province, country, ZIP or foreign postal code, and telephone no.\nABC COMPANY INC"
      },
      {
        name: 'PAYER_NAME_MULTILINE',
        pattern: /(?:PAYER'S?\s+name|Payer'?s?\s+name)\s*\n([A-Za-z\s&.,'-]+?)(?:\n|$)/i,
        example: "PAYER'S name\nAcme Corporation"
      },
      {
        name: 'PAYER_NAME_BASIC',
        pattern: /(?:PAYER'S?\s+name|Payer'?s?\s+name)[:\s]+([A-Za-z\s&.,'-]+?)(?:\s+\d|\n|PAYER'S?\s+|Payer'?s?\s+|TIN|address|street|$)/i,
        example: "PAYER'S NAME ACME CORPORATION"
      }
    ];
    
    // Try payer name patterns
    for (const patternInfo of payerNamePatterns) {
      const match = ocrText.match(patternInfo.pattern);
      if (match && match[1]) {
        const name = match[1].trim();
        if (name.length > 2 && !name.toLowerCase().includes('street address')) {
          info1099.payerName = name;
          console.log(`✅ [Azure DI OCR] Found payer name using ${patternInfo.name}:`, name);
          break;
        }
      }
    }
    
    // === PAYER TIN PATTERNS ===
    const payerTinPatterns = [
      {
        name: 'PAYER_TIN_BASIC',
        pattern: /(?:PAYER'S?\s+TIN|Payer'?s?\s+TIN)[:\s]+(\d{2}[-\s]?\d{7})/i,
        example: "PAYER'S TIN 12-3456789"
      },
      {
        name: 'PAYER_TIN_MULTILINE',
        pattern: /(?:PAYER'S?\s+TIN|Payer'?s?\s+TIN)\s*\n(\d{2}[-\s]?\d{7})/i,
        example: "PAYER'S TIN\n12-3456789"
      }
    ];
    
    // Try payer TIN patterns
    for (const patternInfo of payerTinPatterns) {
      const match = ocrText.match(patternInfo.pattern);
      if (match && match[1]) {
        const tin = match[1].trim();
        if (tin.length >= 9) {
          info1099.payerTIN = tin;
          console.log(`✅ [Azure DI OCR] Found payer TIN using ${patternInfo.name}:`, tin);
          break;
        }
      }
    }
    
    // === PAYER ADDRESS PATTERNS - FIXED ===
    const payerAddressPatterns = [
      {
        name: 'PAYER_ADDRESS_COMPLETE_BLOCK',
        // Extract the complete payer info block: Company name + street + city/state/zip
        pattern: /(?:PAYER'S?\s+name[^\n]*\n)([A-Za-z\s&.,'-]+(?:LLC|Inc|Corp|Company|Co\.?)?)\s*\n([0-9]+\s+[A-Za-z\s]+(?:Drive|Street|St|Ave|Avenue|Blvd|Boulevard|Road|Rd|Lane|Ln)?\.?)\s*\n([A-Za-z\s]+,\s*[A-Z]{2}\s+\d{5}(?:-\d{4})?)/i,
        example: "PAYER'S name, street address...\nAlphaTech Solutions LLC\n920 Tech Drive\nAustin, TX 73301"
      },
      {
        name: 'PAYER_ADDRESS_SIMPLE_BLOCK',
        // Simpler pattern for company + address block
        pattern: /([A-Za-z\s&.,'-]+(?:LLC|Inc|Corp|Company|Co\.?))\s*\n([0-9]+\s+[A-Za-z\s]+(?:Drive|Street|St|Ave|Avenue|Blvd|Boulevard|Road|Rd|Lane|Ln)?\.?)\s*\n([A-Za-z\s]+,\s*[A-Z]{2}\s+\d{5}(?:-\d{4})?)/i,
        example: "AlphaTech Solutions LLC\n920 Tech Drive\nAustin, TX 73301"
      },
      {
        name: 'PAYER_ADDRESS_AFTER_LABEL',
        // Extract everything after the payer label until PAYER'S TIN
        pattern: /PAYER'S?\s+name[^\n]*\n([^\n]+)\s*\n([^\n]+)\s*\n([^\n]+)(?:\s*\n.*?)?(?=PAYER'S?\s+TIN|$)/i,
        example: "PAYER'S name, street address, city...\nAlphaTech Solutions LLC\n920 Tech Drive\nAustin, TX 73301"
      },
      {
        name: 'PAYER_ADDRESS_FALLBACK',
        // Fallback: Look for company name pattern followed by address components
        pattern: /([A-Za-z\s&.,'-]+(?:LLC|Inc|Corp|Company|Co\.?))[^\n]*\n([0-9]+[^\n]+)[^\n]*\n([A-Za-z\s]+,\s*[A-Z]{2}\s+\d{5})/i,
        example: "AlphaTech Solutions LLC\n920 Tech Drive\nAustin, TX 73301"
      }
    ];
    
    // Try payer address patterns
    for (const patternInfo of payerAddressPatterns) {
      const match = ocrText.match(patternInfo.pattern);
      if (match && match[1] && match[2] && match[3]) {
        const companyName = match[1].trim();
        const street = match[2].trim();
        const cityStateZip = match[3].trim();
        
        // Validate the components
        if (companyName.length > 2 && 
            !companyName.toLowerCase().includes('street address') &&
            /\d/.test(street) && // Street should contain numbers
            /[A-Z]{2}\s+\d{5}/.test(cityStateZip)) { // City should have state and zip
          
          const fullAddress = `${companyName} ${street} ${cityStateZip}`;
          info1099.payerAddress = fullAddress;
          console.log(`✅ [Azure DI OCR] Found payer address using ${patternInfo.name}:`, fullAddress);
          break;
        }
      }
    }
    
    return info1099;
  }

  // === W2 PATTERNS ===
  /**
   * ENHANCED: Extracts personal information from W2 OCR text using comprehensive regex patterns
   * Specifically designed for W2 form OCR text patterns with enhanced fallback mechanisms
   * NOW INCLUDES: Multi-employee record handling for documents with multiple W2s
   * @param ocrText - The OCR text to extract information from
   * @param targetEmployeeName - Optional target employee name to match against in multi-employee scenarios
   */
  private extractPersonalInfoFromOCR(ocrText: string, targetEmployeeName?: string): {
    name?: string;
    ssn?: string;
    tin?: string;
    address?: string;
    employerName?: string;
    employerAddress?: string;
    payerName?: string;
    payerTIN?: string;
    payerAddress?: string;
  } {
    console.log('🔍 [Azure DI OCR] Searching for personal info in OCR text...');
    
    const personalInfo: { 
      name?: string; 
      ssn?: string; 
      tin?: string;
      address?: string;
      employerName?: string;
      employerAddress?: string;
      payerName?: string;
      payerTIN?: string;
      payerAddress?: string;
    } = {};
    
    // Detect if this is a W2 or 1099 form to use appropriate extraction
    const formType = this.detectFormType(ocrText);
    
    if (formType === 'W2') {
      // Use W2-specific extraction
      const w2Info = this.extractW2InfoFromOCR(ocrText, targetEmployeeName);
      personalInfo.name = w2Info.name;
      personalInfo.ssn = w2Info.ssn;
      personalInfo.address = w2Info.address;
      personalInfo.employerName = w2Info.employerName;
      personalInfo.employerAddress = w2Info.employerAddress;
    } else if (formType === '1099') {
      // Use 1099-specific extraction
      const form1099Info = this.extract1099InfoFromOCR(ocrText);
      personalInfo.name = form1099Info.name;
      personalInfo.tin = form1099Info.tin;
      personalInfo.address = form1099Info.address;
      personalInfo.payerName = form1099Info.payerName;
      personalInfo.payerTIN = form1099Info.payerTIN;
      personalInfo.payerAddress = form1099Info.payerAddress;
    }
    
    return personalInfo;
  }

  /**
   * ENHANCED: Extracts W2-specific personal information from OCR text
   * NOW INCLUDES: Multi-employee record handling for documents with multiple W2s
   * @param ocrText - The OCR text to extract information from
   * @param targetEmployeeName - Optional target employee name to match against in multi-employee scenarios
   */
  private extractW2InfoFromOCR(ocrText: string, targetEmployeeName?: string): {
    name?: string;
    ssn?: string;
    address?: string;
    employerName?: string;
    employerAddress?: string;
  } {
    console.log('🔍 [Azure DI OCR] Searching for W2 info in OCR text...');
    
    const w2Info: { 
      name?: string; 
      ssn?: string; 
      address?: string;
      employerName?: string;
      employerAddress?: string;
    } = {};
    
    // === EMPLOYEE NAME PATTERNS ===
    const employeeNamePatterns = [
      // EMPLOYEE_NAME_MULTILINE: Extract name that appears after "Employee's name" label
      {
        name: 'EMPLOYEE_NAME_MULTILINE',
        pattern: /(?:Employee'?s?\s+name|EMPLOYEE'S?\s+NAME)\s*\n([A-Za-z\s]+?)(?:\n|$)/i,
        example: "Employee's name\nJohn Doe"
      },
      // EMPLOYEE_NAME_BASIC: Basic employee name extraction
      {
        name: 'EMPLOYEE_NAME_BASIC',
        pattern: /(?:Employee'?s?\s+name|EMPLOYEE'S?\s+NAME)[:\s]+([A-Za-z\s]+?)(?:\s+\d|\n|Employee'?s?\s+|EMPLOYEE'S?\s+|SSN|address|street|$)/i,
        example: "Employee's name JOHN DOE"
      },
      {
        name: 'EMPLOYEE_NAME_COLON',
        pattern: /(?:Employee'?s?\s+name|EMPLOYEE'S?\s+NAME):\s*([A-Za-z\s]+?)(?:\n|Employee'?s?\s+|EMPLOYEE'S?\s+|SSN|address|street|$)/i,
        example: "Employee's name: JOHN DOE"
      }
    ];
    
    // Try employee name patterns
    for (const patternInfo of employeeNamePatterns) {
      const match = ocrText.match(patternInfo.pattern);
      if (match && match[1]) {
        const name = match[1].trim();
        if (name.length > 2 && /^[A-Za-z\s]+$/.test(name)) {
          // If we have a target employee name, check if this matches
          if (targetEmployeeName) {
            const similarity = this.calculateNameSimilarity(name, targetEmployeeName);
            if (similarity > 0.7) { // 70% similarity threshold
              w2Info.name = name;
              console.log(`✅ [Azure DI OCR] Found matching employee name using ${patternInfo.name}:`, name);
              break;
            } else {
              console.log(`⚠️ [Azure DI OCR] Found employee name but similarity too low (${similarity}):`, name);
              continue;
            }
          } else {
            w2Info.name = name;
            console.log(`✅ [Azure DI OCR] Found employee name using ${patternInfo.name}:`, name);
            break;
          }
        }
      }
    }
    
    // === EMPLOYEE SSN PATTERNS ===
    const employeeSsnPatterns = [
      {
        name: 'EMPLOYEE_SSN_BASIC',
        pattern: /(?:Employee'?s?\s+SSN|EMPLOYEE'S?\s+SSN)[:\s]+(\d{3}[-\s]?\d{2}[-\s]?\d{4})/i,
        example: "Employee's SSN 123-45-6789"
      },
      {
        name: 'EMPLOYEE_SSN_MULTILINE',
        pattern: /(?:Employee'?s?\s+SSN|EMPLOYEE'S?\s+SSN)\s*\n(\d{3}[-\s]?\d{2}[-\s]?\d{4})/i,
        example: "Employee's SSN\n123-45-6789"
      }
    ];
    
    // Try employee SSN patterns
    for (const patternInfo of employeeSsnPatterns) {
      const match = ocrText.match(patternInfo.pattern);
      if (match && match[1]) {
        const ssn = match[1].trim();
        if (ssn.length >= 9) {
          w2Info.ssn = ssn;
          console.log(`✅ [Azure DI OCR] Found employee SSN using ${patternInfo.name}:`, ssn);
          break;
        }
      }
    }
    
    // === EMPLOYEE ADDRESS PATTERNS ===
    const employeeAddressPatterns = [
      {
        name: 'EMPLOYEE_ADDRESS_MULTILINE',
        pattern: /(?:Employee'?s?\s+address|EMPLOYEE'S?\s+ADDRESS)\s*\n([^\n]+(?:\n[^\n]+)*?)(?:\n\s*\n|Employer'?s?\s+|EMPLOYER'S?\s+|$)/i,
        example: "Employee's address\n123 Main St\nAnytown, ST 12345"
      },
      {
        name: 'EMPLOYEE_ADDRESS_BASIC',
        pattern: /(?:Employee'?s?\s+address|EMPLOYEE'S?\s+ADDRESS)[:\s]+([^\n]+(?:\n[^\n]+)*?)(?:\n\s*\n|Employer'?s?\s+|EMPLOYER'S?\s+|$)/i,
        example: "Employee's address: 123 Main St, Anytown, ST 12345"
      }
    ];
    
    // Try employee address patterns
    for (const patternInfo of employeeAddressPatterns) {
      const match = ocrText.match(patternInfo.pattern);
      if (match && match[1]) {
        const address = match[1].trim().replace(/\n+/g, ' ');
        if (address.length > 5) {
          w2Info.address = address;
          console.log(`✅ [Azure DI OCR] Found employee address using ${patternInfo.name}:`, address);
          break;
        }
      }
    }
    
    // === EMPLOYER NAME PATTERNS ===
    const employerNamePatterns = [
      {
        name: 'EMPLOYER_NAME_MULTILINE',
        pattern: /(?:Employer'?s?\s+name|EMPLOYER'S?\s+NAME)\s*\n([A-Za-z\s&.,'-]+?)(?:\n|$)/i,
        example: "Employer's name\nAcme Corporation"
      },
      {
        name: 'EMPLOYER_NAME_BASIC',
        pattern: /(?:Employer'?s?\s+name|EMPLOYER'S?\s+NAME)[:\s]+([A-Za-z\s&.,'-]+?)(?:\s+\d|\n|Employer'?s?\s+|EMPLOYER'S?\s+|EIN|address|street|$)/i,
        example: "Employer's name ACME CORPORATION"
      }
    ];
    
    // Try employer name patterns
    for (const patternInfo of employerNamePatterns) {
      const match = ocrText.match(patternInfo.pattern);
      if (match && match[1]) {
        const name = match[1].trim();
        if (name.length > 2) {
          w2Info.employerName = name;
          console.log(`✅ [Azure DI OCR] Found employer name using ${patternInfo.name}:`, name);
          break;
        }
      }
    }
    
    // === EMPLOYER ADDRESS PATTERNS ===
    const employerAddressPatterns = [
      {
        name: 'EMPLOYER_ADDRESS_MULTILINE',
        pattern: /(?:Employer'?s?\s+address|EMPLOYER'S?\s+ADDRESS)\s*\n([^\n]+(?:\n[^\n]+)*?)(?:\n\s*\n|$)/i,
        example: "Employer's address\n456 Business Ave\nBusiness City, ST 67890"
      },
      {
        name: 'EMPLOYER_ADDRESS_BASIC',
        pattern: /(?:Employer'?s?\s+address|EMPLOYER'S?\s+ADDRESS)[:\s]+([^\n]+(?:\n[^\n]+)*?)(?:\n\s*\n|$)/i,
        example: "Employer's address: 456 Business Ave, Business City, ST 67890"
      }
    ];
    
    // Try employer address patterns
    for (const patternInfo of employerAddressPatterns) {
      const match = ocrText.match(patternInfo.pattern);
      if (match && match[1]) {
        const address = match[1].trim().replace(/\n+/g, ' ');
        if (address.length > 5) {
          w2Info.employerAddress = address;
          console.log(`✅ [Azure DI OCR] Found employer address using ${patternInfo.name}:`, address);
          break;
        }
      }
    }
    
    return w2Info;
  }

  /**
   * Calculates similarity between two names for multi-employee matching
   * Uses Levenshtein distance normalized by string length
   */
  private calculateNameSimilarity(name1: string, name2: string): number {
    const s1 = name1.toLowerCase().trim();
    const s2 = name2.toLowerCase().trim();
    
    if (s1 === s2) return 1.0;
    
    const maxLength = Math.max(s1.length, s2.length);
    if (maxLength === 0) return 1.0;
    
    const distance = this.levenshteinDistance(s1, s2);
    return (maxLength - distance) / maxLength;
  }

  /**
   * Calculates Levenshtein distance between two strings
   */
  private levenshteinDistance(str1: string, str2: string): number {
    const matrix = [];
    
    for (let i = 0; i <= str2.length; i++) {
      matrix[i] = [i];
    }
    
    for (let j = 0; j <= str1.length; j++) {
      matrix[0][j] = j;
    }
    
    for (let i = 1; i <= str2.length; i++) {
      for (let j = 1; j <= str1.length; j++) {
        if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
          matrix[i][j] = matrix[i - 1][j - 1];
        } else {
          matrix[i][j] = Math.min(
            matrix[i - 1][j - 1] + 1,
            matrix[i][j - 1] + 1,
            matrix[i - 1][j] + 1
          );
        }
      }
    }
    
    return matrix[str2.length][str1.length];
  }

  /**
   * Extracts W2 fields from OCR text using comprehensive regex patterns
   */
  private extractW2FieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    console.log('🔍 [Azure DI OCR] Extracting W2 fields from OCR text...');
    
    const data = { ...baseData };
    
    // Extract personal information first
    const personalInfo = this.extractW2InfoFromOCR(ocrText);
    
    // Add personal info to data
    if (personalInfo.name) data.employeeName = personalInfo.name;
    if (personalInfo.ssn) data.employeeSSN = personalInfo.ssn;
    if (personalInfo.address) data.employeeAddress = personalInfo.address;
    if (personalInfo.employerName) data.employerName = personalInfo.employerName;
    if (personalInfo.employerAddress) data.employerAddress = personalInfo.employerAddress;
    
    // Extract wages from Box 1
    const wagesFromOCR = this.extractWagesFromOCR(ocrText);
    if (wagesFromOCR > 0) {
      data.wages = wagesFromOCR;
    }
    
    return data;
  }

  /**
   * Extracts wages from W2 OCR text using multiple patterns
   */
  private extractWagesFromOCR(ocrText: string): number {
    console.log('🔍 [Azure DI OCR] Extracting wages from OCR text...');
    
    // Box 1 wage patterns
    const wagePatterns = [
      // Pattern 1: "1 Wages, tips, other compensation" followed by amount
      /1\s+Wages,\s+tips,\s+other\s+compensation[:\s]*\$?([0-9,]+\.?\d*)/i,
      // Pattern 2: "Box 1" followed by amount
      /Box\s+1[:\s]*\$?([0-9,]+\.?\d*)/i,
      // Pattern 3: Just "1" followed by amount (more generic)
      /(?:^|\n)\s*1\s+([0-9,]+\.?\d*)/m,
      // Pattern 4: "Wages" followed by amount
      /Wages[:\s]*\$?([0-9,]+\.?\d*)/i
    ];
    
    for (const pattern of wagePatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          console.log('✅ [Azure DI OCR] Found wages:', amount);
          return amount;
        }
      }
    }
    
    console.log('⚠️ [Azure DI OCR] No wages found in OCR text');
    return 0;
  }

  /**
   * Extracts 1099-MISC fields from OCR text using comprehensive regex patterns
   */
  private extract1099MiscFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    console.log('🔍 [Azure DI OCR] Extracting 1099-MISC fields from OCR text...');
    
    const data = { ...baseData };
    
    // Extract personal information first
    const personalInfo = this.extract1099InfoFromOCR(ocrText);
    
    // Add personal info to data
    if (personalInfo.name) data.recipientName = personalInfo.name;
    if (personalInfo.tin) data.recipientTIN = personalInfo.tin;
    if (personalInfo.address) data.recipientAddress = personalInfo.address;
    if (personalInfo.payerName) data.payerName = personalInfo.payerName;
    if (personalInfo.payerTIN) data.payerTIN = personalInfo.payerTIN;
    if (personalInfo.payerAddress) data.payerAddress = personalInfo.payerAddress;
    
    // Extract account number
    const accountNumberMatch = ocrText.match(/Account number[:\s]*([A-Za-z0-9\-]+)/i);
    if (accountNumberMatch && accountNumberMatch[1]) {
      data.accountNumber = accountNumberMatch[1].trim();
      console.log('✅ [Azure DI OCR] Found account number:', data.accountNumber);
    }
    
    // === BOX EXTRACTION PATTERNS ===
    
    // Box 1: Rents
    const box1Patterns = [
      /(?:1\s*Rents|Box 1|Rents)[:\s]*\$?([0-9,]+\.?\d*)/i,
      /Rents[:\s]*\$?([0-9,]+\.?\d*)/i
    ];
    
    for (const pattern of box1Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.rents = amount;
          console.log('✅ [Azure DI OCR] Found Box 1 - Rents:', amount);
          break;
        }
      }
    }
    
    // Box 2: Royalties
    const box2Patterns = [
      /(?:2\s*Royalties|Box 2|Royalties)[:\s]*\$?([0-9,]+\.?\d*)/i,
      /Royalties[:\s]*\$?([0-9,]+\.?\d*)/i
    ];
    
    for (const pattern of box2Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.royalties = amount;
          console.log('✅ [Azure DI OCR] Found Box 2 - Royalties:', amount);
          break;
        }
      }
    }
    
    // Box 3: Other Income
    const box3Patterns = [
      /(?:3\s*Other income|Box 3|Other income)[:\s]*\$?([0-9,]+\.?\d*)/i,
      /Other income[:\s]*\$?([0-9,]+\.?\d*)/i
    ];
    
    for (const pattern of box3Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.otherIncome = amount;
          console.log('✅ [Azure DI OCR] Found Box 3 - Other Income:', amount);
          break;
        }
      }
    }
    
    // Box 4: Federal Income Tax Withheld
    const box4Patterns = [
      /(?:4\s*Federal income tax withheld|Box 4|Federal income tax withheld)[:\s]*\$?([0-9,]+\.?\d*)/i,
      /Federal income tax withheld[:\s]*\$?([0-9,]+\.?\d*)/i
    ];
    
    for (const pattern of box4Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.federalTaxWithheld = amount;
          console.log('✅ [Azure DI OCR] Found Box 4 - Federal Tax Withheld:', amount);
          break;
        }
      }
    }
    
    // Box 5: Fishing Boat Proceeds
    const box5Patterns = [
      /(?:5\s*Fishing boat proceeds|Box 5|Fishing boat proceeds)[:\s]*\$?([0-9,]+\.?\d*)/i,
      /Fishing boat proceeds[:\s]*\$?([0-9,]+\.?\d*)/i
    ];
    
    for (const pattern of box5Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.fishingBoatProceeds = amount;
          console.log('✅ [Azure DI OCR] Found Box 5 - Fishing Boat Proceeds:', amount);
          break;
        }
      }
    }
    
    // Box 6: Medical and Health Care Payments
    const box6Patterns = [
      /(?:6\s*Medical and health care payments|Box 6|Medical and health care payments)[:\s]*\$?([0-9,]+\.?\d*)/i,
      /Medical and health care payments[:\s]*\$?([0-9,]+\.?\d*)/i
    ];
    
    for (const pattern of box6Patterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.medicalHealthPayments = amount;
          console.log('✅ [Azure DI OCR] Found Box 6 - Medical and Health Care Payments:', amount);
          break;
        }
      }
    }
    
    console.log('✅ [Azure DI OCR] Completed 1099-MISC field extraction');
    return data;
  }

  /**
   * Extracts 1099-DIV fields from OCR text using comprehensive regex patterns
   */
  private extract1099DivFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    console.log('🔍 [Azure DI OCR] Extracting 1099-DIV fields from OCR text...');
    
    const data = { ...baseData };
    
    // Extract personal information first
    const personalInfo = this.extract1099InfoFromOCR(ocrText);
    
    // Add personal info to data
    if (personalInfo.name) data.recipientName = personalInfo.name;
    if (personalInfo.tin) data.recipientTIN = personalInfo.tin;
    if (personalInfo.address) data.recipientAddress = personalInfo.address;
    if (personalInfo.payerName) data.payerName = personalInfo.payerName;
    if (personalInfo.payerTIN) data.payerTIN = personalInfo.payerTIN;
    if (personalInfo.payerAddress) data.payerAddress = personalInfo.payerAddress;
    
    // Box extraction patterns for 1099-DIV would go here
    // Similar to 1099-MISC but with dividend-specific fields
    
    return data;
  }

  /**
   * Extracts 1099-NEC fields from OCR text using comprehensive regex patterns
   */
  private extract1099NecFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    console.log('🔍 [Azure DI OCR] Extracting 1099-NEC fields from OCR text...');
    
    const data = { ...baseData };
    
    // Extract personal information first
    const personalInfo = this.extract1099InfoFromOCR(ocrText);
    
    // Add personal info to data
    if (personalInfo.name) data.recipientName = personalInfo.name;
    if (personalInfo.tin) data.recipientTIN = personalInfo.tin;
    if (personalInfo.address) data.recipientAddress = personalInfo.address;
    if (personalInfo.payerName) data.payerName = personalInfo.payerName;
    if (personalInfo.payerTIN) data.payerTIN = personalInfo.payerTIN;
    if (personalInfo.payerAddress) data.payerAddress = personalInfo.payerAddress;
    
    // Box extraction patterns for 1099-NEC would go here
    // Similar to other 1099 forms but with NEC-specific fields
    
    return data;
  }

  /**
   * Extracts generic fields from OCR text
   */
  private extractGenericFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    console.log('🔍 [Azure DI OCR] Extracting generic fields from OCR text...');
    
    const data = { ...baseData };
    
    // Generic extraction patterns would go here
    // This is a fallback for unknown document types
    
    return data;
  }

  /**
   * Parses address components from a full address string
   */
  private extractAddressParts(fullAddress: string, ocrText: string): {
    street?: string;
    city?: string;
    state?: string;
    zipCode?: string;
  } {
    const addressParts: {
      street?: string;
      city?: string;
      state?: string;
      zipCode?: string;
    } = {};
    
    // Try to parse the address components
    const addressMatch = fullAddress.match(/^(.+?),?\s*([A-Za-z\s]+),?\s*([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$/);
    
    if (addressMatch) {
      addressParts.street = addressMatch[1].trim();
      addressParts.city = addressMatch[2].trim();
      addressParts.state = addressMatch[3].trim();
      addressParts.zipCode = addressMatch[4].trim();
    } else {
      // Fallback: try to extract components from OCR text
      const zipMatch = ocrText.match(/\b(\d{5}(?:-\d{4})?)\b/);
      if (zipMatch) {
        addressParts.zipCode = zipMatch[1];
      }
      
      const stateMatch = ocrText.match(/\b([A-Z]{2})\s+\d{5}/);
      if (stateMatch) {
        addressParts.state = stateMatch[1];
      }
    }
    
    return addressParts;
  }

  /**
   * Parses amount strings into numbers
   */
  private parseAmount(value: any): number {
    if (typeof value === 'number') {
      return value;
    }
    
    if (typeof value === 'string') {
      // Remove currency symbols, commas, and whitespace
      const cleanValue = value.replace(/[$,\s]/g, '');
      const parsed = parseFloat(cleanValue);
      return isNaN(parsed) ? 0 : parsed;
    }
    
    return 0;
  }
}
