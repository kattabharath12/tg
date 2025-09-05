


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
              console.log(`🔄 [Azure DI] Document type correction: ${documentType} → ${ocrBasedType}`);
              
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
          // Re-throw other errors
          throw modelError;
        }
      }
      
      // Add corrected document type to the extracted data if it was corrected
      if (correctedDocumentType) {
        extractedData.correctedDocumentType = correctedDocumentType;
      }
      
      console.log('✅ [Azure DI] Final extracted data:', JSON.stringify(extractedData, null, 2));
      return extractedData;
      
    } catch (error: any) {
      console.error('❌ [Azure DI] Error extracting data from document:', error);
      throw new Error(`Document analysis failed: ${error.message}`);
    }
  }

  private getModelIdForDocumentType(documentType: string): string {
    // Map document types to Azure Document Intelligence model IDs
    const modelMappings: { [key: string]: string } = {
      'W2': 'prebuilt-tax.us.w2',
      'W2_CORRECTED': 'prebuilt-tax.us.w2',
      'FORM_1099_INT': 'prebuilt-tax.us.1099int',
      'FORM_1099_DIV': 'prebuilt-tax.us.1099div',
      'FORM_1099_MISC': 'prebuilt-tax.us.1099misc',
      'FORM_1099_NEC': 'prebuilt-tax.us.1099nec',
      'FORM_1099_R': 'prebuilt-tax.us.1099r',
      'FORM_1099_G': 'prebuilt-tax.us.1099g',
      'FORM_1099_K': 'prebuilt-tax.us.1099k',
      'FORM_1098': 'prebuilt-tax.us.1098',
      'FORM_1098_E': 'prebuilt-tax.us.1098e',
      'FORM_1098_T': 'prebuilt-tax.us.1098t',
      'FORM_5498': 'prebuilt-tax.us.5498'
    };
    
    return modelMappings[documentType] || 'prebuilt-read';
  }

  private extractTaxDocumentFields(result: any, documentType: string): ExtractedFieldData {
    console.log('🔍 [Azure DI] Extracting fields for document type:', documentType);
    
    // Extract full text from the document
    const fullText = result.content || '';
    
    const extractedData: ExtractedFieldData = {
      fullText: fullText
    };
    
    // Process documents based on type
    if (result.documents && result.documents.length > 0) {
      const document = result.documents[0];
      
      switch (documentType) {
        case 'W2':
        case 'W2_CORRECTED':
          return this.processW2Fields(document.fields, extractedData);
        case 'FORM_1099_INT':
          return this.process1099IntFields(document.fields, extractedData);
        case 'FORM_1099_DIV':
          return this.process1099DivFields(document.fields, extractedData);
        case 'FORM_1099_MISC':
          return this.process1099MiscFields(document.fields, extractedData);
        case 'FORM_1099_NEC':
          return this.process1099NecFields(document.fields, extractedData);
        case 'FORM_1099_R':
          return this.process1099RFields(document.fields, extractedData);
        case 'FORM_1099_G':
          return this.process1099GFields(document.fields, extractedData);
        case 'FORM_1099_K':
          return this.process1099KFields(document.fields, extractedData);
        case 'FORM_1098':
          return this.process1098Fields(document.fields, extractedData);
        case 'FORM_1098_E':
          return this.process1098EFields(document.fields, extractedData);
        case 'FORM_1098_T':
          return this.process1098TFields(document.fields, extractedData);
        case 'FORM_5498':
          return this.process5498Fields(document.fields, extractedData);
        default:
          console.log('⚠️ [Azure DI] Unknown document type, returning basic extracted data');
          return extractedData;
      }
    }
    
    return extractedData;
  }

  private extractTaxDocumentFieldsFromOCR(result: any, documentType: string): ExtractedFieldData {
    console.log('🔍 [Azure DI OCR] Extracting fields from OCR for document type:', documentType);
    
    // Extract full text from the document
    const fullText = result.content || '';
    
    const extractedData: ExtractedFieldData = {
      fullText: fullText
    };
    
    // Process documents based on type using OCR
    switch (documentType) {
      case 'W2':
      case 'W2_CORRECTED':
        return this.extractW2FieldsFromOCR(fullText, extractedData);
      case 'FORM_1099_INT':
        return this.extract1099IntFieldsFromOCR(fullText, extractedData);
      case 'FORM_1099_DIV':
        return this.extract1099DivFieldsFromOCR(fullText, extractedData);
      case 'FORM_1099_MISC':
        return this.extract1099MiscFieldsFromOCR(fullText, extractedData);
      case 'FORM_1099_NEC':
        return this.extract1099NecFieldsFromOCR(fullText, extractedData);
      case 'FORM_1099_R':
        return this.extract1099RFieldsFromOCR(fullText, extractedData);
      case 'FORM_1099_G':
        return this.extract1099GFieldsFromOCR(fullText, extractedData);
      case 'FORM_1099_K':
        return this.extract1099KFieldsFromOCR(fullText, extractedData);
      case 'FORM_1098':
        return this.extract1098FieldsFromOCR(fullText, extractedData);
      case 'FORM_1098_E':
        return this.extract1098EFieldsFromOCR(fullText, extractedData);
      case 'FORM_1098_T':
        return this.extract1098TFieldsFromOCR(fullText, extractedData);
      case 'FORM_5498':
        return this.extract5498FieldsFromOCR(fullText, extractedData);
      default:
        console.log('⚠️ [Azure DI OCR] Unknown document type, returning basic extracted data');
        return extractedData;
    }
  }

  private process1099IntFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    const data = { ...baseData };
    
    // Enhanced field mappings for all 1099-INT boxes and fields
    const fieldMappings = {
      // Payer and recipient information
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'AccountNumber': 'accountNumber',
      
      // Box 1-17 mappings
      'InterestIncome': 'interestIncome',                           // Box 1
      'EarlyWithdrawalPenalty': 'earlyWithdrawalPenalty',          // Box 2
      'InterestOnUSTreasuryObligations': 'interestOnUSavingsBonds', // Box 3
      'InterestOnUSavingsBonds': 'interestOnUSavingsBonds',        // Box 3 (alternative name)
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',            // Box 4
      'InvestmentExpenses': 'investmentExpenses',                   // Box 5
      'ForeignTaxPaid': 'foreignTaxPaid',                          // Box 6
      'TaxExemptInterest': 'taxExemptInterest',                    // Box 8
      'SpecifiedPrivateActivityBondInterest': 'specifiedPrivateActivityBondInterest', // Box 9
      'MarketDiscount': 'marketDiscount',                          // Box 10
      'BondPremium': 'bondPremium',                                // Box 11
      'BondPremiumOnTreasuryObligations': 'bondPremiumOnTreasuryObligations', // Box 12
      'BondPremiumOnTaxExemptBond': 'bondPremiumOnTaxExemptBond',  // Box 13
      'TaxExemptAndTaxCreditBondCUSIPNo': 'taxExemptAndTaxCreditBondCUSIPNo', // Box 14
      'State': 'state',                                            // Box 15
      'StateIdentificationNo': 'stateIdentificationNo',            // Box 16
      'StateTaxWithheld': 'stateTaxWithheld',                      // Box 17
      
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
        if (['accountNumber', 'state', 'stateIdentificationNo', 'taxExemptAndTaxCreditBondCUSIPNo'].includes(mappedFieldName)) {
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
      'interestIncome',
      'earlyWithdrawalPenalty', 
      'interestOnUSavingsBonds',
      'federalTaxWithheld',
      'investmentExpenses',
      'foreignTaxPaid',
      'taxExemptInterest',
      'specifiedPrivateActivityBondInterest',
      'marketDiscount',
      'bondPremium',
      'bondPremiumOnTreasuryObligations',
      'bondPremiumOnTaxExemptBond',
      'stateTaxWithheld'
    ];
    
    for (const field of criticalFields) {
      const structuredValue = structuredData[field];
      const ocrValue = ocrData[field];
      
      // If OCR found a value but structured extraction didn't, use OCR value
      if (!structuredValue && ocrValue) {
        correctedData[field] = ocrValue;
        correctionsMade++;
        console.log(`🔧 [Azure DI] Corrected ${field}: null → ${ocrValue} (from OCR)`);
      }
      // If both have values but they differ significantly, prefer OCR for critical fields
      else if (structuredValue && ocrValue && typeof structuredValue === 'number' && typeof ocrValue === 'number') {
        const difference = Math.abs(structuredValue - ocrValue);
        const percentDifference = difference / Math.max(structuredValue, ocrValue);
        
        if (percentDifference > 0.1) { // More than 10% difference
          correctedData[field] = ocrValue;
          correctionsMade++;
          console.log(`🔧 [Azure DI] Corrected ${field}: ${structuredValue} → ${ocrValue} (significant difference)`);
        }
      }
    }
    
    if (correctionsMade > 0) {
      console.log(`✅ [Azure DI] Made ${correctionsMade} corrections to 1099-INT field mappings`);
    } else {
      console.log('✅ [Azure DI] No corrections needed for 1099-INT field mappings');
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
      'rents', 'royalties', 'otherIncome', 'federalTaxWithheld', 'fishingBoatProceeds',
      'medicalHealthPayments', 'nonemployeeCompensation', 'substitutePayments', 
      'cropInsuranceProceeds', 'grossProceedsAttorney', 'fishPurchases', 
      'section409ADeferrals', 'excessGoldenParachutePayments', 'nonqualifiedDeferredCompensation',
      'section409AIncome', 'stateTaxWithheld', 'stateIncome'
    ];
    
    for (const field of criticalFields) {
      const structuredValue = structuredData[field];
      const ocrValue = ocrData[field];
      
      // If OCR found a value but structured extraction didn't, use OCR value
      if (!structuredValue && ocrValue) {
        correctedData[field] = ocrValue;
        correctionsMade++;
        console.log(`🔧 [Azure DI] Corrected ${field}: null → ${ocrValue} (from OCR)`);
      }
      // If both have values but they differ significantly, prefer OCR for critical fields
      else if (structuredValue && ocrValue && typeof structuredValue === 'number' && typeof ocrValue === 'number') {
        const difference = Math.abs(structuredValue - ocrValue);
        const percentDifference = difference / Math.max(structuredValue, ocrValue);
        
        if (percentDifference > 0.1) { // More than 10% difference
          correctedData[field] = ocrValue;
          correctionsMade++;
          console.log(`🔧 [Azure DI] Corrected ${field}: ${structuredValue} → ${ocrValue} (significant difference)`);
        }
      }
    }
    
    if (correctionsMade > 0) {
      console.log(`✅ [Azure DI] Made ${correctionsMade} corrections to 1099-MISC field mappings`);
    } else {
      console.log('✅ [Azure DI] No corrections needed for 1099-MISC field mappings');
    }
    
    return correctedData;
  }

  // Placeholder methods for other document types - implement as needed
  private processW2Fields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for W2 processing (existing logic)
    return baseData;
  }

  private process1099NecFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1099-NEC processing
    return baseData;
  }

  private process1099RFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1099-R processing
    return baseData;
  }

  private process1099GFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1099-G processing
    return baseData;
  }

  private process1099KFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1099-K processing
    return baseData;
  }

  private process1098Fields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1098 processing
    return baseData;
  }

  private process1098EFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1098-E processing
    return baseData;
  }

  private process1098TFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1098-T processing
    return baseData;
  }

  private process5498Fields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 5498 processing
    return baseData;
  }

  private extract1099IntFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    console.log('🔍 [Azure DI OCR] Extracting 1099-INT fields from OCR text...');
    
    const data = { ...baseData };
    
    // Extract personal information using 1099-specific patterns
    const personalInfo = this.extractPersonalInfoFromOCR(ocrText);
    if (personalInfo.name) data.recipientName = personalInfo.name;
    if (personalInfo.tin) data.recipientTIN = personalInfo.tin;
    if (personalInfo.address) data.recipientAddress = personalInfo.address;
    if (personalInfo.payerName) data.payerName = personalInfo.payerName;
    if (personalInfo.payerTIN) data.payerTIN = personalInfo.payerTIN;
    
    // Extract account number
    const accountNumberPatterns = [
      /Account\s+number[:\s]*([A-Z0-9\-]+)/i,
      /Account[:\s]*([A-Z0-9\-]+)/i,
      /Acct[:\s]*([A-Z0-9\-]+)/i
    ];
    
    for (const pattern of accountNumberPatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        data.accountNumber = match[1].trim();
        console.log(`✅ [Azure DI OCR] Found account number: ${data.accountNumber}`);
        break;
      }
    }
    
    // Enhanced 1099-INT specific amount patterns for all boxes
    const amountPatterns = {
      interestIncome: [
        /1\s+Interest\s+income\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*1\s+\$?([0-9,]+\.?\d{0,2})/m,
        /Box\s+1[:\s]*\$?([0-9,]+\.?\d{0,2})/i
      ],
      earlyWithdrawalPenalty: [
        /2\s+Early\s+withdrawal\s+penalty\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*2\s+\$?([0-9,]+\.?\d{0,2})/m,
        /Box\s+2[:\s]*\$?([0-9,]+\.?\d{0,2})/i
      ],
      interestOnUSavingsBonds: [
        /3\s+Interest\s+on\s+U\.?S\.?\s+Savings\s+Bonds\s+and\s+Treasury\s+obligations\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*3\s+\$?([0-9,]+\.?\d{0,2})/m,
        /Box\s+3[:\s]*\$?([0-9,]+\.?\d{0,2})/i
      ],
      federalTaxWithheld: [
        /4\s+Federal\s+income\s+tax\s+withheld\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*4\s+\$?([0-9,]+\.?\d{0,2})/m,
        /Box\s+4[:\s]*\$?([0-9,]+\.?\d{0,2})/i
      ],
      investmentExpenses: [
        /5\s+Investment\s+expenses\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*5\s+\$?([0-9,]+\.?\d{0,2})/m,
        /Box\s+5[:\s]*\$?([0-9,]+\.?\d{0,2})/i
      ],
      foreignTaxPaid: [
        /6\s+Foreign\s+tax\s+paid\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*6\s+\$?([0-9,]+\.?\d{0,2})/m,
        /Box\s+6[:\s]*\$?([0-9,]+\.?\d{0,2})/i
      ],
      taxExemptInterest: [
        /8\s+Tax-exempt\s+interest\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*8\s+\$?([0-9,]+\.?\d{0,2})/m,
        /Box\s+8[:\s]*\$?([0-9,]+\.?\d{0,2})/i
      ],
      specifiedPrivateActivityBondInterest: [
        /9\s+Specified\s+private\s+activity\s+bond\s+interest\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*9\s+\$?([0-9,]+\.?\d{0,2})/m,
        /Box\s+9[:\s]*\$?([0-9,]+\.?\d{0,2})/i
      ],
      marketDiscount: [
        /10\s+Market\s+discount\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*10\s+\$?([0-9,]+\.?\d{0,2})/m,
        /Box\s+10[:\s]*\$?([0-9,]+\.?\d{0,2})/i
      ],
      bondPremium: [
        /11\s+Bond\s+premium\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*11\s+\$?([0-9,]+\.?\d{0,2})/m,
        /Box\s+11[:\s]*\$?([0-9,]+\.?\d{0,2})/i
      ],
      bondPremiumOnTreasuryObligations: [
        /12\s+Bond\s+premium\s+on\s+Treasury\s+obligations\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*12\s+\$?([0-9,]+\.?\d{0,2})/m,
        /Box\s+12[:\s]*\$?([0-9,]+\.?\d{0,2})/i
      ],
      bondPremiumOnTaxExemptBond: [
        /13\s+Bond\s+premium\s+on\s+tax-exempt\s+bond\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*13\s+\$?([0-9,]+\.?\d{0,2})/m,
        /Box\s+13[:\s]*\$?([0-9,]+\.?\d{0,2})/i
      ],
      taxExemptAndTaxCreditBondCUSIPNo: [
        /14\s+Tax-exempt\s+and\s+tax\s+credit\s+bond\s+CUSIP\s+no\.\s*[\n\s]*([A-Z0-9]+)/i,
        /(?:^|\n)\s*14\s+([A-Z0-9]+)/m,
        /Box\s+14[:\s]*([A-Z0-9]+)/i
      ],
      stateTaxWithheld: [
        /17\s+State\s+tax\s+withheld\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*17\s+\$?([0-9,]+\.?\d{0,2})/m,
        /Box\s+17[:\s]*\$?([0-9,]+\.?\d{0,2})/i
      ]
    };
    
    // Extract state and state identification number
    const statePatterns = [
      /15\s+State\s*[\n\s]*([A-Z]{2})/i,
      /(?:^|\n)\s*15\s+([A-Z]{2})/m,
      /Box\s+15[:\s]*([A-Z]{2})/i
    ];
    
    for (const pattern of statePatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        data.state = match[1].trim().toUpperCase();
        console.log(`✅ [Azure DI OCR] Found state: ${data.state}`);
        break;
      }
    }
    
    const stateIdPatterns = [
      /16\s+State\s+identification\s+no\.\s*[\n\s]*([0-9]+)/i,
      /(?:^|\n)\s*16\s+([0-9]+)/m,
      /Box\s+16[:\s]*([0-9]+)/i
    ];
    
    for (const pattern of stateIdPatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        data.stateIdentificationNo = match[1].trim();
        console.log(`✅ [Azure DI OCR] Found state identification no: ${data.stateIdentificationNo}`);
        break;
      }
    }
    
    // Extract all amount fields
    for (const [fieldName, patterns] of Object.entries(amountPatterns)) {
      for (const pattern of patterns) {
        const match = ocrText.match(pattern);
        if (match && match[1]) {
          if (fieldName === 'taxExemptAndTaxCreditBondCUSIPNo') {
            // Handle CUSIP number as text
            data[fieldName] = match[1].trim();
            console.log(`✅ [Azure DI OCR] Found ${fieldName}: ${data[fieldName]}`);
          } else {
            // Handle monetary amounts
            const amountStr = match[1].replace(/,/g, '');
            const amount = parseFloat(amountStr);
            
            if (!isNaN(amount) && amount >= 0) {
              data[fieldName] = amount;
              console.log(`✅ [Azure DI OCR] Found ${fieldName}: $${amount}`);
            }
          }
          break;
        }
      }
    }
    
    return data;
  }

  private extract1099DivFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    console.log('🔍 [Azure DI OCR] Extracting 1099-DIV fields from OCR text...');
    
    const data = { ...baseData };
    
    // Extract personal information using 1099-specific patterns
    const personalInfo = this.extractPersonalInfoFromOCR(ocrText);
    if (personalInfo.name) data.recipientName = personalInfo.name;
    if (personalInfo.tin) data.recipientTIN = personalInfo.tin;
    if (personalInfo.address) data.recipientAddress = personalInfo.address;
    if (personalInfo.payerName) data.payerName = personalInfo.payerName;
    if (personalInfo.payerTIN) data.payerTIN = personalInfo.payerTIN;
    
    // Extract 1099-DIV specific amounts
    const amountPatterns = {
      ordinaryDividends: [
        /1a\s+Ordinary\s+dividends\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*1a\s+\$?([0-9,]+\.?\d{0,2})/m
      ],
      qualifiedDividends: [
        /1b\s+Qualified\s+dividends\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*1b\s+\$?([0-9,]+\.?\d{0,2})/m
      ],
      totalCapitalGain: [
        /2a\s+Total\s+capital\s+gain\s+distributions\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*2a\s+\$?([0-9,]+\.?\d{0,2})/m
      ],
      federalTaxWithheld: [
        /4\s+Federal\s+income\s+tax\s+withheld\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*4\s+\$?([0-9,]+\.?\d{0,2})/m
      ]
    };
    
    for (const [fieldName, patterns] of Object.entries(amountPatterns)) {
      for (const pattern of patterns) {
        const match = ocrText.match(pattern);
        if (match && match[1]) {
          const amountStr = match[1].replace(/,/g, '');
          const amount = parseFloat(amountStr);
          
          if (!isNaN(amount) && amount >= 0) {
            data[fieldName] = amount;
            console.log(`✅ [Azure DI OCR] Found ${fieldName}: $${amount}`);
            break;
          }
        }
      }
    }
    
    return data;
  }

  private extract1099MiscFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    console.log('🔍 [Azure DI OCR] Extracting 1099-MISC fields from OCR text...');
    
    const data = { ...baseData };
    
    // Extract personal information using 1099-specific patterns
    const personalInfo = this.extractPersonalInfoFromOCR(ocrText);
    if (personalInfo.name) data.recipientName = personalInfo.name;
    if (personalInfo.tin) data.recipientTIN = personalInfo.tin;
    if (personalInfo.address) data.recipientAddress = personalInfo.address;
    if (personalInfo.payerName) data.payerName = personalInfo.payerName;
    if (personalInfo.payerTIN) data.payerTIN = personalInfo.payerTIN;
    
    // Extract 1099-MISC specific amounts for all boxes
    const amountPatterns = {
      rents: [
        /1\s+Rents\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*1\s+\$?([0-9,]+\.?\d{0,2})/m
      ],
      royalties: [
        /2\s+Royalties\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*2\s+\$?([0-9,]+\.?\d{0,2})/m
      ],
      otherIncome: [
        /3\s+Other\s+income\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*3\s+\$?([0-9,]+\.?\d{0,2})/m
      ],
      federalTaxWithheld: [
        /4\s+Federal\s+income\s+tax\s+withheld\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*4\s+\$?([0-9,]+\.?\d{0,2})/m
      ],
      fishingBoatProceeds: [
        /5\s+Fishing\s+boat\s+proceeds\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*5\s+\$?([0-9,]+\.?\d{0,2})/m
      ],
      medicalHealthPayments: [
        /6\s+Medical\s+and\s+health\s+care\s+payments\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*6\s+\$?([0-9,]+\.?\d{0,2})/m
      ],
      nonemployeeCompensation: [
        /7\s+Nonemployee\s+compensation\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*7\s+\$?([0-9,]+\.?\d{0,2})/m
      ],
      substitutePayments: [
        /8\s+Substitute\s+payments\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*8\s+\$?([0-9,]+\.?\d{0,2})/m
      ],
      cropInsuranceProceeds: [
        /9\s+Crop\s+insurance\s+proceeds\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*9\s+\$?([0-9,]+\.?\d{0,2})/m
      ],
      grossProceedsAttorney: [
        /10\s+Gross\s+proceeds\s+paid\s+to\s+an\s+attorney\s*[\n\s]*\$?([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*10\s+\$?([0-9,]+\.?\d{0,2})/m
      ]
    };
    
    for (const [fieldName, patterns] of Object.entries(amountPatterns)) {
      for (const pattern of patterns) {
        const match = ocrText.match(pattern);
        if (match && match[1]) {
          const amountStr = match[1].replace(/,/g, '');
          const amount = parseFloat(amountStr);
          
          if (!isNaN(amount) && amount >= 0) {
            data[fieldName] = amount;
            console.log(`✅ [Azure DI OCR] Found ${fieldName}: $${amount}`);
            break;
          }
        }
      }
    }
    
    return data;
  }

  // Placeholder methods for other OCR extractions
  private extractW2FieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for W2 OCR extraction
    return baseData;
  }

  private extract1099NecFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1099-NEC OCR extraction
    return baseData;
  }

  private extract1099RFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1099-R OCR extraction
    return baseData;
  }

  private extract1099GFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1099-G OCR extraction
    return baseData;
  }

  private extract1099KFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1099-K OCR extraction
    return baseData;
  }

  private extract1098FieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1098 OCR extraction
    return baseData;
  }

  private extract1098EFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1098-E OCR extraction
    return baseData;
  }

  private extract1098TFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1098-T OCR extraction
    return baseData;
  }

  private extract5498FieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 5498 OCR extraction
    return baseData;
  }

  private extractPersonalInfoFromOCR(ocrText: string): any {
    // Extract personal information patterns
    const personalInfo: any = {};
    
    // Extract recipient name (usually appears after "RECIPIENT'S name")
    const namePatterns = [
      /RECIPIENT'S\s+name[:\s]*([A-Za-z\s,\.]+)/i,
      /Recipient[:\s]*([A-Za-z\s,\.]+)/i
    ];
    
    for (const pattern of namePatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        personalInfo.name = match[1].trim();
        break;
      }
    }
    
    // Extract recipient TIN/SSN
    const tinPatterns = [
      /RECIPIENT'S\s+TIN[:\s]*([0-9\-]+)/i,
      /SSN[:\s]*([0-9\-]+)/i,
      /TIN[:\s]*([0-9\-]+)/i
    ];
    
    for (const pattern of tinPatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        personalInfo.tin = match[1].trim();
        break;
      }
    }
    
    // Extract payer name
    const payerNamePatterns = [
      /PAYER'S\s+name[:\s]*([A-Za-z\s,\.]+)/i,
      /Payer[:\s]*([A-Za-z\s,\.]+)/i
    ];
    
    for (const pattern of payerNamePatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        personalInfo.payerName = match[1].trim();
        break;
      }
    }
    
    // Extract payer TIN
    const payerTinPatterns = [
      /PAYER'S\s+TIN[:\s]*([0-9\-]+)/i,
      /Payer\s+TIN[:\s]*([0-9\-]+)/i
    ];
    
    for (const pattern of payerTinPatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        personalInfo.payerTIN = match[1].trim();
        break;
      }
    }
    
    return personalInfo;
  }

  private analyzeDocumentTypeFromOCR(ocrText: string): string {
    // Analyze OCR text to determine document type
    const text = ocrText.toLowerCase();
    
    if (text.includes('1099-int') || text.includes('form 1099-int')) {
      return 'FORM_1099_INT';
    } else if (text.includes('1099-div') || text.includes('form 1099-div')) {
      return 'FORM_1099_DIV';
    } else if (text.includes('1099-misc') || text.includes('form 1099-misc')) {
      return 'FORM_1099_MISC';
    } else if (text.includes('1099-nec') || text.includes('form 1099-nec')) {
      return 'FORM_1099_NEC';
    } else if (text.includes('w-2') || text.includes('form w-2')) {
      return 'W2';
    }
    
    return 'UNKNOWN';
  }

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
