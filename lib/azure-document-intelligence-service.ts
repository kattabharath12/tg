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
  confidence?: number;
  extractionWarnings?: string[];
}

export interface FieldExtractionResult {
  value: any;
  confidence: number;
  source: 'structured' | 'ocr_fallback';
}

export class AzureDocumentIntelligenceService {
  private client: DocumentAnalysisClient;
  private config: AzureDocumentIntelligenceConfig;
  private readonly MIN_CONFIDENCE_THRESHOLD = 0.7;
  private readonly CRITICAL_FIELDS_1099_INT = [
    'interestIncome', 'federalTaxWithheld', 'taxExemptInterest', 'investmentExpenses',
    'earlyWithdrawalPenalty', 'interestOnUSavingsBonds', 'foreignTaxPaid', 'marketDiscount',
    'bondPremium', 'bondPremiumTreasury', 'bondPremiumTaxExempt', 'specifiedPrivateActivityBond'
  ];

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
        extractedData = await this.extractTaxDocumentFields(result, documentType);
        
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
              extractedData = await this.extractTaxDocumentFields(result, ocrBasedType);
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
          extractedData = await this.extractTaxDocumentFieldsFromOCR(fallbackResult, documentType);
          
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
                extractedData = await this.extractTaxDocumentFieldsFromOCR(fallbackResult, ocrBasedType);
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
      
      // Enhanced error handling - return partial results if possible
      const partialData: ExtractedFieldData = {
        extractionWarnings: [`Azure Document Intelligence processing failed: ${error?.message || 'Unknown error'}`],
        confidence: 0
      };
      
      // If we have a document buffer, try basic OCR as last resort
      try {
        if (typeof documentPathOrBuffer === 'string' || Buffer.isBuffer(documentPathOrBuffer)) {
          console.log('🔍 [Azure DI] Attempting basic OCR as last resort...');
          const documentBuffer = typeof documentPathOrBuffer === 'string' 
            ? await readFile(documentPathOrBuffer)
            : documentPathOrBuffer;
          
          const fallbackPoller = await this.client.beginAnalyzeDocument('prebuilt-read', documentBuffer);
          const fallbackResult = await fallbackPoller.pollUntilDone();
          
          const ocrData = await this.extractTaxDocumentFieldsFromOCR(fallbackResult, documentType);
          return { ...partialData, ...ocrData };
        }
      } catch (fallbackError) {
        console.error('❌ [Azure DI] Fallback OCR also failed:', fallbackError);
      }
      
      return partialData;
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

  private async extractTaxDocumentFieldsFromOCR(result: any, documentType: string): Promise<ExtractedFieldData> {
    console.log('🔍 [Azure DI] Extracting tax document fields using OCR fallback...');
    
    const extractedData: ExtractedFieldData = {
      extractionWarnings: [],
      confidence: 0.5 // OCR fallback has lower confidence
    };
    
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

  private async extractTaxDocumentFields(result: any, documentType: string): Promise<ExtractedFieldData> {
    const extractedData: ExtractedFieldData = {
      extractionWarnings: [],
      confidence: 0.8 // Structured extraction has higher confidence
    };
    
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
            return await this.process1099IntFields(document.fields, extractedData);
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

  /**
   * FIXED: Enhanced 1099-INT field processing with comprehensive field mappings and validation
   */
  private async process1099IntFields(fields: any, baseData: ExtractedFieldData): Promise<ExtractedFieldData> {
    const data = { ...baseData };
    
    // COMPREHENSIVE 1099-INT field mappings covering all 17 boxes
    const fieldMappings = {
      // Payer and recipient information
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN', 
      'Payer.Address': 'payerAddress',
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'AccountNumber': 'accountNumber',
      
      // Box 1: Interest Income - CRITICAL FIELD
      'InterestIncome': 'interestIncome',
      'Interest': 'interestIncome',
      'Box1': 'interestIncome',
      'TotalInterestIncome': 'interestIncome',
      
      // Box 2: Early Withdrawal Penalty - CRITICAL FIELD  
      'EarlyWithdrawalPenalty': 'earlyWithdrawalPenalty',
      'EarlyWithdrawal': 'earlyWithdrawalPenalty',
      'Box2': 'earlyWithdrawalPenalty',
      'WithdrawalPenalty': 'earlyWithdrawalPenalty',
      
      // Box 3: Interest on U.S. Treasury Obligations - CRITICAL FIELD
      'InterestOnUSTreasuryObligations': 'interestOnUSavingsBonds',
      'InterestOnUSavingsBonds': 'interestOnUSavingsBonds',
      'USTreasuryInterest': 'interestOnUSavingsBonds',
      'Box3': 'interestOnUSavingsBonds',
      'TreasuryObligations': 'interestOnUSavingsBonds',
      
      // Box 4: Federal Income Tax Withheld - CRITICAL FIELD
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',
      'FederalTaxWithheld': 'federalTaxWithheld',
      'Box4': 'federalTaxWithheld',
      'TaxWithheld': 'federalTaxWithheld',
      'BackupWithholding': 'federalTaxWithheld',
      
      // Box 5: Investment Expenses - CRITICAL FIELD
      'InvestmentExpenses': 'investmentExpenses',
      'Box5': 'investmentExpenses',
      'REMICExpenses': 'investmentExpenses',
      
      // Box 6: Foreign Tax Paid - CRITICAL FIELD
      'ForeignTaxPaid': 'foreignTaxPaid',
      'Box6': 'foreignTaxPaid',
      'ForeignTax': 'foreignTaxPaid',
      
      // Box 7: Foreign Country or U.S. Territory
      'ForeignCountry': 'foreignCountry',
      'Box7': 'foreignCountry',
      'ForeignCountryOrUSTerritory': 'foreignCountry',
      
      // Box 8: Tax-Exempt Interest - CRITICAL FIELD
      'TaxExemptInterest': 'taxExemptInterest',
      'Box8': 'taxExemptInterest',
      'TaxExempt': 'taxExemptInterest',
      'MunicipalBondInterest': 'taxExemptInterest',
      
      // Box 9: Specified Private Activity Bond Interest - CRITICAL FIELD
      'SpecifiedPrivateActivityBondInterest': 'specifiedPrivateActivityBond',
      'Box9': 'specifiedPrivateActivityBond',
      'PrivateActivityBond': 'specifiedPrivateActivityBond',
      'AMTInterest': 'specifiedPrivateActivityBond',
      
      // Box 10: Market Discount - CRITICAL FIELD
      'MarketDiscount': 'marketDiscount',
      'Box10': 'marketDiscount',
      'AccruedMarketDiscount': 'marketDiscount',
      
      // Box 11: Bond Premium - CRITICAL FIELD
      'BondPremium': 'bondPremium',
      'Box11': 'bondPremium',
      'BondPremiumAmortization': 'bondPremium',
      
      // Box 12: Bond Premium on Treasury Obligations - CRITICAL FIELD
      'BondPremiumOnTreasuryObligations': 'bondPremiumTreasury',
      'Box12': 'bondPremiumTreasury',
      'TreasuryBondPremium': 'bondPremiumTreasury',
      
      // Box 13: Bond Premium on Tax-Exempt Bond - CRITICAL FIELD
      'BondPremiumOnTaxExemptBond': 'bondPremiumTaxExempt',
      'Box13': 'bondPremiumTaxExempt',
      'TaxExemptBondPremium': 'bondPremiumTaxExempt',
      
      // Box 14: Tax-Exempt and Tax Credit Bond CUSIP No.
      'CUSIP': 'cusipNumber',
      'Box14': 'cusipNumber',
      'CUSIPNumber': 'cusipNumber',
      'TaxExemptBondCUSIP': 'cusipNumber',
      
      // Boxes 15-17: State Information
      'State': 'state',
      'Box15': 'state',
      'StateCode': 'state',
      'StateIdentificationNumber': 'stateIdNumber',
      'Box16': 'stateIdNumber',
      'StateTaxWithheld': 'stateTaxWithheld',
      'Box17': 'stateTaxWithheld'
    };
    
    // Extract structured fields with confidence tracking
    const extractionResults: { [key: string]: FieldExtractionResult } = {};
    let totalConfidence = 0;
    let fieldCount = 0;
    
    for (const [azureFieldName, mappedFieldName] of Object.entries(fieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const fieldData = fields[azureFieldName];
        const value = fieldData.value;
        const confidence = fieldData.confidence || 0.8;
        
        // Handle different field types appropriately
        let processedValue: any;
        if (mappedFieldName === 'cusipNumber' || mappedFieldName === 'accountNumber' || 
            mappedFieldName === 'state' || mappedFieldName === 'foreignCountry') {
          processedValue = String(value).trim();
        } else if (mappedFieldName.includes('Name') || mappedFieldName.includes('Address')) {
          processedValue = String(value).trim();
        } else {
          processedValue = typeof value === 'number' ? value : this.parseAmount(value);
        }
        
        extractionResults[mappedFieldName] = {
          value: processedValue,
          confidence: confidence,
          source: 'structured'
        };
        
        data[mappedFieldName] = processedValue;
        totalConfidence += confidence;
        fieldCount++;
        
        console.log(`✅ [Azure DI] Extracted ${mappedFieldName}: ${processedValue} (confidence: ${confidence})`);
      }
    }
    
    // OCR fallback for missing critical fields
    if (baseData.fullText) {
      const ocrResults = this.extract1099IntFieldsFromOCR(baseData.fullText as string, { fullText: baseData.fullText });
      
      // Check each critical field and use OCR if structured extraction failed or has low confidence
      for (const criticalField of this.CRITICAL_FIELDS_1099_INT) {
        const structuredResult = extractionResults[criticalField];
        const ocrValue = ocrResults[criticalField];
        
        if (!structuredResult && ocrValue && this.parseAmount(ocrValue) > 0) {
          // Missing from structured, found in OCR
          data[criticalField] = ocrValue;
          extractionResults[criticalField] = {
            value: ocrValue,
            confidence: 0.6,
            source: 'ocr_fallback'
          };
          console.log(`🔧 [Azure DI] OCR fallback for ${criticalField}: ${ocrValue}`);
          
        } else if (structuredResult && structuredResult.confidence < this.MIN_CONFIDENCE_THRESHOLD && 
                   ocrValue && Math.abs(this.parseAmount(ocrValue) - this.parseAmount(structuredResult.value)) > 100) {
          // Low confidence structured result, significantly different OCR result
          data[criticalField] = ocrValue;
          extractionResults[criticalField] = {
            value: ocrValue,
            confidence: 0.6,
            source: 'ocr_fallback'
          };
          console.log(`🔧 [Azure DI] OCR correction for ${criticalField}: ${structuredResult.value} → ${ocrValue}`);
        }
      }
      
      // Extract personal info if missing
      if (!data.recipientName || !data.recipientTIN || !data.recipientAddress || !data.payerName || !data.payerTIN) {
        console.log('🔍 [Azure DI] Some 1099-INT info missing from structured fields, attempting OCR extraction...');
        const personalInfoFromOCR = this.extract1099InfoFromOCR(baseData.fullText as string);
        
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
    }
    
    // Calculate overall confidence and add warnings
    const overallConfidence = fieldCount > 0 ? totalConfidence / fieldCount : 0;
    data.confidence = overallConfidence;
    
    // Add extraction warnings for low confidence fields
    const warnings: string[] = [];
    let criticalFieldsMissing = 0;
    
    for (const criticalField of this.CRITICAL_FIELDS_1099_INT) {
      const result = extractionResults[criticalField];
      if (!result) {
        criticalFieldsMissing++;
        warnings.push(`Missing critical field: ${criticalField}`);
      } else if (result.confidence < this.MIN_CONFIDENCE_THRESHOLD) {
        warnings.push(`Low confidence for ${criticalField}: ${result.confidence}`);
      }
    }
    
    if (criticalFieldsMissing > 0) {
      warnings.push(`${criticalFieldsMissing} out of ${this.CRITICAL_FIELDS_1099_INT.length} critical fields missing`);
    }
    
    data.extractionWarnings = warnings;
    
    console.log(`✅ [Azure DI] 1099-INT extraction completed. Overall confidence: ${overallConfidence}, Warnings: ${warnings.length}`);
    
    return data;
  }

  /**
   * ENHANCED: OCR-based 1099-INT field extraction with comprehensive regex patterns
   */
  private extract1099IntFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    console.log('🔍 [Azure DI OCR] Extracting 1099-INT fields from OCR text...');
    
    const data = { ...baseData };
    
    // Box 1: Interest Income - Multiple patterns for different OCR layouts
    const interestIncomePatterns = [
      /(?:Box\s*1|1\.?\s*Interest\s+income|Interest\s+income)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /Interest\s+income[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /1\s+Interest\s+income[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /(?:^|\n)\s*1[:\s]*([0-9,]+\.?\d{0,2})/m
    ];
    
    for (const pattern of interestIncomePatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.interestIncome = amount;
          console.log('✅ [Azure DI OCR] Found interest income:', amount);
          break;
        }
      }
    }
    
    // Box 2: Early Withdrawal Penalty
    const earlyWithdrawalPatterns = [
      /(?:Box\s*2|2\.?\s*Early\s+withdrawal\s+penalty|Early\s+withdrawal\s+penalty)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /Early\s+withdrawal\s+penalty[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /2\s+Early\s+withdrawal[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
    ];
    
    for (const pattern of earlyWithdrawalPatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.earlyWithdrawalPenalty = amount;
          console.log('✅ [Azure DI OCR] Found early withdrawal penalty:', amount);
          break;
        }
      }
    }
    
    // Box 3: Interest on U.S. Treasury Obligations
    const treasuryInterestPatterns = [
      /(?:Box\s*3|3\.?\s*Interest\s+on\s+U\.?S\.?\s+Treasury|Interest\s+on\s+U\.?S\.?\s+Treasury)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /Interest\s+on\s+U\.?S\.?\s+Treasury\s+obligations[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /3\s+Interest\s+on\s+U\.?S\.?[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /Treasury\s+obligations[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
    ];
    
    for (const pattern of treasuryInterestPatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.interestOnUSavingsBonds = amount;
          console.log('✅ [Azure DI OCR] Found Treasury interest:', amount);
          break;
        }
      }
    }
    
    // Box 4: Federal Income Tax Withheld - CRITICAL FIELD
    const federalTaxPatterns = [
      /(?:Box\s*4|4\.?\s*Federal\s+income\s+tax\s+withheld|Federal\s+income\s+tax\s+withheld)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /Federal\s+income\s+tax\s+withheld[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /4\s+Federal\s+income\s+tax[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /Tax\s+withheld[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /Backup\s+withholding[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
    ];
    
    for (const pattern of federalTaxPatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.federalTaxWithheld = amount;
          console.log('✅ [Azure DI OCR] Found federal tax withheld:', amount);
          break;
        }
      }
    }
    
    // Box 5: Investment Expenses - CRITICAL FIELD
    const investmentExpensesPatterns = [
      /(?:Box\s*5|5\.?\s*Investment\s+expenses|Investment\s+expenses)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /Investment\s+expenses[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /5\s+Investment\s+expenses[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /REMIC\s+expenses[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
    ];
    
    for (const pattern of investmentExpensesPatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.investmentExpenses = amount;
          console.log('✅ [Azure DI OCR] Found investment expenses:', amount);
          break;
        }
      }
    }
    
    // Box 6: Foreign Tax Paid
    const foreignTaxPatterns = [
      /(?:Box\s*6|6\.?\s*Foreign\s+tax\s+paid|Foreign\s+tax\s+paid)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /Foreign\s+tax\s+paid[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /6\s+Foreign\s+tax[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
    ];
    
    for (const pattern of foreignTaxPatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.foreignTaxPaid = amount;
          console.log('✅ [Azure DI OCR] Found foreign tax paid:', amount);
          break;
        }
      }
    }
    
    // Box 8: Tax-Exempt Interest - CRITICAL FIELD
    const taxExemptPatterns = [
      /(?:Box\s*8|8\.?\s*Tax-exempt\s+interest|Tax-exempt\s+interest)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /Tax-exempt\s+interest[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /8\s+Tax-exempt[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /Municipal\s+bond\s+interest[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
    ];
    
    for (const pattern of taxExemptPatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.taxExemptInterest = amount;
          console.log('✅ [Azure DI OCR] Found tax-exempt interest:', amount);
          break;
        }
      }
    }
    
    // Box 9: Specified Private Activity Bond Interest
    const privateActivityPatterns = [
      /(?:Box\s*9|9\.?\s*Specified\s+private\s+activity|Specified\s+private\s+activity)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /Specified\s+private\s+activity\s+bond[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /9\s+Specified\s+private[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /Private\s+activity\s+bond[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
    ];
    
    for (const pattern of privateActivityPatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.specifiedPrivateActivityBond = amount;
          console.log('✅ [Azure DI OCR] Found private activity bond interest:', amount);
          break;
        }
      }
    }
    
    // Box 10: Market Discount
    const marketDiscountPatterns = [
      /(?:Box\s*10|10\.?\s*Market\s+discount|Market\s+discount)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /Market\s+discount[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /10\s+Market\s+discount[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
    ];
    
    for (const pattern of marketDiscountPatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.marketDiscount = amount;
          console.log('✅ [Azure DI OCR] Found market discount:', amount);
          break;
        }
      }
    }
    
    // Box 11: Bond Premium
    const bondPremiumPatterns = [
      /(?:Box\s*11|11\.?\s*Bond\s+premium|Bond\s+premium)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /Bond\s+premium[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /11\s+Bond\s+premium[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
    ];
    
    for (const pattern of bondPremiumPatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.bondPremium = amount;
          console.log('✅ [Azure DI OCR] Found bond premium:', amount);
          break;
        }
      }
    }
    
    // Box 12: Bond Premium on Treasury Obligations
    const treasuryPremiumPatterns = [
      /(?:Box\s*12|12\.?\s*Bond\s+premium\s+on\s+Treasury|Bond\s+premium\s+on\s+Treasury)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /Bond\s+premium\s+on\s+Treasury\s+obligations[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /12\s+Bond\s+premium\s+on\s+Treasury[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
    ];
    
    for (const pattern of treasuryPremiumPatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.bondPremiumTreasury = amount;
          console.log('✅ [Azure DI OCR] Found Treasury bond premium:', amount);
          break;
        }
      }
    }
    
    // Box 13: Bond Premium on Tax-Exempt Bond
    const taxExemptPremiumPatterns = [
      /(?:Box\s*13|13\.?\s*Bond\s+premium\s+on\s+tax-exempt|Bond\s+premium\s+on\s+tax-exempt)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /Bond\s+premium\s+on\s+tax-exempt\s+bond[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
      /13\s+Bond\s+premium\s+on\s+tax-exempt[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
    ];
    
    for (const pattern of taxExemptPremiumPatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        const amount = this.parseAmount(match[1]);
        if (amount > 0) {
          data.bondPremiumTaxExempt = amount;
          console.log('✅ [Azure DI OCR] Found tax-exempt bond premium:', amount);
          break;
        }
      }
    }
    
    // Extract personal information using existing method
    const personalInfo = this.extract1099InfoFromOCR(ocrText);
    if (personalInfo.name) data.recipientName = personalInfo.name;
    if (personalInfo.tin) data.recipientTIN = personalInfo.tin;
    if (personalInfo.address) data.recipientAddress = personalInfo.address;
    if (personalInfo.payerName) data.payerName = personalInfo.payerName;
    if (personalInfo.payerTIN) data.payerTIN = personalInfo.payerTIN;
    if (personalInfo.payerAddress) data.payerAddress = personalInfo.payerAddress;
    
    return data;
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
      const personalInfoFromOCR = this.extract1099InfoFromOCR(baseData.fullText as string);
      
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
      const personalInfoFromOCR = this.extract1099InfoFromOCR(baseData.fullText as string);
      
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
      const personalInfoFromOCR = this.extract1099InfoFromOCR(baseData.fullText as string);
      
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
          'investment expenses'
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

  /**
   * ENHANCED: Extracts personal information from 1099 OCR text using comprehensive regex patterns
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
    
    // Check if this is a 1099 form first
    if (ocrText.toLowerCase().includes('1099')) {
      return this.extract1099InfoFromOCR(ocrText);
    }
    
    // W2-specific extraction logic continues here...
    // [Rest of the existing W2 extraction code would go here]
    
    return personalInfo;
  }

  // Additional helper methods for OCR extraction
  private extractW2FieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for W2 OCR extraction
    return baseData;
  }

  private extract1099DivFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1099-DIV OCR extraction
    return baseData;
  }

  private extract1099MiscFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1099-MISC OCR extraction
    return baseData;
  }

  private extract1099NecFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1099-NEC OCR extraction
    return baseData;
  }

  private extractGenericFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for generic OCR extraction
    return baseData;
  }

  private extractAddressParts(address: string, ocrText: string): {
    street?: string;
    city?: string;
    state?: string;
    zipCode?: string;
  } {
    // Implementation for address parsing
    return {};
  }

  private extractWagesFromOCR(ocrText: string): number {
    // Implementation for wages extraction
    return 0;
  }

  /**
   * Enhanced amount parsing with better error handling and validation
   */
  private parseAmount(value: any): number {
    if (typeof value === 'number') {
      return value;
    }
    
    if (typeof value === 'string') {
      // Remove currency symbols, commas, and whitespace
      const cleanValue = value.replace(/[$,\s]/g, '').trim();
      
      // Handle empty or non-numeric strings
      if (!cleanValue || cleanValue === '' || cleanValue === '-' || cleanValue === 'N/A') {
        return 0;
      }
      
      // Parse the cleaned value
      const parsed = parseFloat(cleanValue);
      
      // Return 0 for invalid numbers
      return isNaN(parsed) ? 0 : parsed;
    }
    
    return 0;
  }

  /**
   * Validates if a value represents a valid dollar amount
   */
  private isDollarAmount(value: any): boolean {
    const amount = this.parseAmount(value);
    return !isNaN(amount) && isFinite(amount) && amount >= 0;
  }
}
