import { DocumentAnalysisClient, AzureKeyCredential } from "@azure/ai-form-recognizer";
import { DocumentType } from "@prisma/client";

// Configuration interface
export interface AzureDocumentIntelligenceConfig {
  endpoint: string;
  apiKey: string;
}

// Enhanced field extraction result with confidence scoring
export interface FieldExtractionResult {
  value: any;
  confidence: number;
  source: 'structured' | 'ocr_fallback' | 'enhanced_ocr' | 'cross_validated' | 'auto_corrected';
  originalValue?: any;
  correctionApplied?: string;
}

// Enhanced extracted field data interface
export interface ExtractedFieldData {
  [key: string]: string | number | DocumentType | number[] | string[] | undefined;
  fullText?: string;
  confidence?: number;
  extractionWarnings?: string[];
  correctedDocumentType?: DocumentType;
  processingMetrics?: {
    structuredFieldsExtracted: number;
    ocrFallbacksUsed: number;
    autoCorrectionsApplied: number;
    validationWarnings: number;
    totalProcessingTime: number;
  };
}

export class AzureDocumentIntelligenceServiceEnhanced {
  private client: DocumentAnalysisClient;
  private readonly MIN_CONFIDENCE_THRESHOLD = 0.7;
  private readonly CRITICAL_FIELDS_1099_INT = [
    'interestIncome', 'federalTaxWithheld', 'taxExemptInterest', 
    'investmentExpenses', 'earlyWithdrawalPenalty', 'interestOnUSavingsBonds',
    'foreignTaxPaid', 'marketDiscount', 'bondPremium', 'bondPremiumTreasury',
    'bondPremiumTaxExempt', 'specifiedPrivateActivityBond'
  ];

  constructor(config: AzureDocumentIntelligenceConfig) {
    this.client = new DocumentAnalysisClient(
      config.endpoint,
      new AzureKeyCredential(config.apiKey)
    );
  }

  async extractDataFromDocument(filePath: string, documentType: DocumentType): Promise<ExtractedFieldData> {
    const startTime = Date.now();
    console.log(`🔍 [Azure DI Enhanced] Starting extraction for ${documentType}`);
    
    try {
      const documentBuffer = require('fs').readFileSync(filePath);
      const modelId = this.getModelIdForDocumentType(documentType);
      
      console.log(`🔍 [Azure DI Enhanced] Using model: ${modelId}`);
      
      const poller = await this.client.beginAnalyzeDocument(modelId, documentBuffer);
      const result = await poller.pollUntilDone();
      
      // Extract structured data
      const structuredResults = await this.extractStructuredFields(result, documentType);
      
      // Enhanced OCR extraction
      const enhancedOcrResults = await this.extractEnhancedOCRFields(result, documentType);
      
      // Cross-validate and merge results
      const mergedResults = await this.crossValidateAndMerge(
        structuredResults, 
        enhancedOcrResults, 
        documentType
      );
      
      // Apply automatic error corrections
      const correctedResults = await this.applyAutoCorrections(mergedResults, documentType);
      
      // Calculate final confidence and metrics
      const finalResults = await this.calculateFinalMetrics(
        correctedResults, 
        documentType, 
        startTime
      );
      
      console.log(`✅ [Azure DI Enhanced] Extraction completed with ${finalResults.confidence}% confidence`);
      return finalResults;
      
    } catch (error: any) {
      console.error('❌ [Azure DI Enhanced] Processing error:', error);
      return this.handleExtractionError(error, documentType, startTime);
    }
  }

  private getModelIdForDocumentType(documentType: DocumentType): string {
    switch (documentType) {
      case 'W2':
        return 'prebuilt-tax.us.w2';
      case 'FORM_1099_INT':
      case 'FORM_1099_MISC':
      case 'FORM_1099_NEC':
      case 'FORM_1099_DIV':
        return 'prebuilt-tax.us.1099';
      default:
        return 'prebuilt-read';
    }
  }

  private async extractStructuredFields(result: any, documentType: DocumentType): Promise<Record<string, FieldExtractionResult>> {
    const extractionResults: Record<string, FieldExtractionResult> = {};
    
    if (!result.documents || result.documents.length === 0) {
      console.log('⚠️ [Azure DI Enhanced] No structured documents found');
      return extractionResults;
    }

    const document = result.documents[0];
    const fieldMappings = this.getFieldMappingsForDocumentType(documentType);
    
    for (const [azureFieldName, localFieldName] of Object.entries(fieldMappings)) {
      const field = document.fields?.[azureFieldName];
      if (field && field.value !== undefined && field.value !== null) {
        extractionResults[localFieldName] = {
          value: this.parseFieldValue(field.value),
          confidence: field.confidence || 0.5,
          source: 'structured'
        };
      }
    }
    
    console.log(`🔍 [Azure DI Enhanced] Structured extraction: ${Object.keys(extractionResults).length} fields`);
    return extractionResults;
  }

  private async extractEnhancedOCRFields(result: any, documentType: DocumentType): Promise<Record<string, FieldExtractionResult>> {
    const ocrText = result.content || '';
    if (!ocrText) {
      console.log('⚠️ [Azure DI Enhanced] No OCR text available');
      return {};
    }

    // Enhanced OCR preprocessing
    const preprocessedText = this.preprocessOCRText(ocrText);
    
    switch (documentType) {
      case 'FORM_1099_INT':
        return this.extractEnhanced1099IntFields(preprocessedText);
      case 'FORM_1099_MISC':
        return this.extractEnhanced1099MiscFields(preprocessedText);
      case 'FORM_1099_NEC':
        return this.extractEnhanced1099NecFields(preprocessedText);
      case 'FORM_1099_DIV':
        return this.extractEnhanced1099DivFields(preprocessedText);
      case 'W2':
        return this.extractEnhancedW2Fields(preprocessedText);
      default:
        return {};
    }
  }

  private preprocessOCRText(ocrText: string): string {
    let processed = ocrText;
    
    // Character error corrections (common OCR mistakes)
    const characterCorrections = {
      // Numbers
      'O': '0', 'o': '0', 'D': '0', 'Q': '0',
      'l': '1', 'I': '1', '|': '1',
      'Z': '2', 'z': '2',
      'E': '3',
      'A': '4', 'h': '4',
      'S': '5', 's': '5',
      'G': '6', 'b': '6',
      'T': '7',
      'B': '8',
      'g': '9', 'q': '9',
      
      // Letters
      'rn': 'm', 'vv': 'w', 'ii': 'n'
    };
    
    // Apply character corrections in monetary contexts
    processed = processed.replace(/\$[\s]*([O|o|D|Q|l|I|\||Z|z|E|A|h|S|s|G|b|T|B|g|q|,|\d|\.|\s]+)/g, (match) => {
      let corrected = match;
      for (const [wrong, right] of Object.entries(characterCorrections)) {
        corrected = corrected.replace(new RegExp(wrong, 'g'), right);
      }
      return corrected;
    });
    
    // Word corrections
    const wordCorrections = {
      'lnterest': 'Interest',
      'Eariy': 'Early',
      'Forelgn': 'Foreign',
      'Wlthheld': 'Withheld',
      'Wlthdrawal': 'Withdrawal',
      'Penaliy': 'Penalty',
      'Treasu ry': 'Treasury',
      'Obllgatlons': 'Obligations',
      'lnvestment': 'Investment',
      'Dlscount': 'Discount',
      'Premlum': 'Premium'
    };
    
    for (const [wrong, right] of Object.entries(wordCorrections)) {
      processed = processed.replace(new RegExp(wrong, 'gi'), right);
    }
    
    // Clean up spacing around dollar signs and numbers
    processed = processed.replace(/\$\s+/g, '$');
    processed = processed.replace(/(\d)\s+(\d)/g, '$1$2');
    processed = processed.replace(/,\s+(\d)/g, ',$1');
    
    return processed;
  }

  private extractEnhanced1099IntFields(ocrText: string): Record<string, FieldExtractionResult> {
    const results: Record<string, FieldExtractionResult> = {};
    
    // Enhanced patterns for each 1099-INT field
    const fieldPatterns = {
      // Box 1 - Interest Income
      interestIncome: [
        /(?:Box\s*1|1\.?\s*Interest\s+income|Interest\s+income)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /1\s+Interest\s+income[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Interest\s+income[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /(?:^|\n)\s*1\s+([0-9,]+\.?\d{0,2})/m
      ],
      
      // Box 2 - Early Withdrawal Penalty
      earlyWithdrawalPenalty: [
        /(?:Box\s*2|2\.?\s*Early\s+withdrawal\s+penalty|Early\s+withdrawal\s+penalty)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /2\s+Early\s+withdrawal\s+penalty[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Early\s+withdrawal\s+penalty[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /withdrawal\s+penalty[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
      ],
      
      // Box 3 - Interest on US Savings Bonds and Treasury Obligations
      interestOnUSavingsBonds: [
        /(?:Box\s*3|3\.?\s*Interest\s+on\s+U\.?S\.?\s+Savings\s+Bonds|Interest\s+on\s+U\.?S\.?\s+Savings\s+Bonds)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /3\s+Interest\s+on\s+U\.?S\.?\s+Savings\s+Bonds[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Interest\s+on\s+U\.?S\.?\s+Savings\s+Bonds\s+and\s+Treasury\s+obligations[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /U\.?S\.?\s+Savings\s+Bonds[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Treasury\s+obligations[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
      ],
      
      // Box 4 - Federal Income Tax Withheld
      federalTaxWithheld: [
        /(?:Box\s*4|4\.?\s*Federal\s+income\s+tax\s+withheld|Federal\s+income\s+tax\s+withheld)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /4\s+Federal\s+income\s+tax\s+withheld[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Federal\s+income\s+tax\s+withheld[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Tax\s+withheld[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Backup\s+withholding[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
      ],
      
      // Box 5 - Investment Expenses
      investmentExpenses: [
        /(?:Box\s*5|5\.?\s*Investment\s+expenses|Investment\s+expenses)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /5\s+Investment\s+expenses[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Investment\s+expenses[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
      ],
      
      // Box 6 - Foreign Tax Paid
      foreignTaxPaid: [
        /(?:Box\s*6|6\.?\s*Foreign\s+tax\s+paid|Foreign\s+tax\s+paid)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /6\s+Foreign\s+tax\s+paid[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Foreign\s+tax\s+paid[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Foreign\s+tax[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
      ],
      
      // Box 8 - Tax-Exempt Interest
      taxExemptInterest: [
        /(?:Box\s*8|8\.?\s*Tax.?exempt\s+interest|Tax.?exempt\s+interest)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /8\s+Tax.?exempt\s+interest[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Tax.?exempt\s+interest[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Tax\s+exempt[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
      ],
      
      // Box 10 - Market Discount
      marketDiscount: [
        /(?:Box\s*10|10\.?\s*Market\s+discount|Market\s+discount)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /10\s+Market\s+discount[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Market\s+discount[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
      ],
      
      // Box 11 - Bond Premium
      bondPremium: [
        /(?:Box\s*11|11\.?\s*Bond\s+premium|Bond\s+premium)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /11\s+Bond\s+premium[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Bond\s+premium[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
      ],
      
      // Box 12 - Bond Premium on Treasury Obligations
      bondPremiumTreasury: [
        /(?:Box\s*12|12\.?\s*Bond\s+premium\s+on\s+Treasury|Bond\s+premium\s+on\s+Treasury)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /12\s+Bond\s+premium\s+on\s+Treasury\s+obligations[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Bond\s+premium\s+on\s+Treasury\s+obligations[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
      ],
      
      // Box 13 - Bond Premium on Tax-Exempt Bond
      bondPremiumTaxExempt: [
        /(?:Box\s*13|13\.?\s*Bond\s+premium\s+on\s+tax.?exempt|Bond\s+premium\s+on\s+tax.?exempt)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /13\s+Bond\s+premium\s+on\s+tax.?exempt\s+bond[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Bond\s+premium\s+on\s+tax.?exempt\s+bond[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i
      ]
    };
    
    // Extract each field using multiple patterns
    for (const [fieldName, patterns] of Object.entries(fieldPatterns)) {
      let bestMatch: FieldExtractionResult | null = null;
      let highestConfidence = 0;
      
      for (const pattern of patterns) {
        const match = ocrText.match(pattern);
        if (match && match[1]) {
          const amount = this.parseEnhancedAmount(match[1]);
          if (amount > 0) {
            const confidence = this.calculatePatternConfidence(pattern, match, ocrText);
            if (confidence > highestConfidence) {
              highestConfidence = confidence;
              bestMatch = {
                value: amount,
                confidence: confidence,
                source: 'enhanced_ocr'
              };
            }
          }
        }
      }
      
      if (bestMatch) {
        results[fieldName] = bestMatch;
      }
    }
    
    // Extract personal information with enhanced patterns
    const personalInfoPatterns = {
      payerName: [
        /PAYER'S\s+name[^\\n]*\\n([^\\n]+)/i,
        /Payer[^\\n]*\\n([A-Za-z][^\\n]+)/i
      ],
      recipientName: [
        /RECIPIENT'S\s+name[^\\n]*\\n([^\\n]+)/i,
        /Recipient[^\\n]*\\n([A-Za-z][^\\n]+)/i
      ],
      payerTIN: [
        /PAYER'S\s+TIN[^\\n]*\\n([0-9-]+)/i,
        /TIN[^\\n]*\\n([0-9-]+)/i
      ]
    };
    
    for (const [fieldName, patterns] of Object.entries(personalInfoPatterns)) {
      for (const pattern of patterns) {
        const match = ocrText.match(pattern);
        if (match && match[1] && match[1].trim()) {
          results[fieldName] = {
            value: match[1].trim(),
            confidence: 0.85,
            source: 'enhanced_ocr'
          };
          break;
        }
      }
    }
    
    console.log(`🔍 [Azure DI Enhanced] Enhanced OCR extraction: ${Object.keys(results).length} fields`);
    return results;
  }

  private parseEnhancedAmount(amountStr: string): number {
    if (!amountStr) return 0;
    
    // Enhanced amount parsing with OCR error correction
    let cleaned = amountStr.toString()
      .replace(/[^\d.,]/g, '') // Remove non-numeric characters except . and ,
      .replace(/^[.,]+/, '') // Remove leading dots/commas
      .replace(/[.,]+$/, ''); // Remove trailing dots/commas
    
    // Handle common OCR errors in numbers
    cleaned = cleaned
      .replace(/O/g, '0')
      .replace(/o/g, '0')
      .replace(/l/g, '1')
      .replace(/I/g, '1')
      .replace(/S/g, '5')
      .replace(/s/g, '5')
      .replace(/Z/g, '2')
      .replace(/B/g, '8')
      .replace(/G/g, '6');
    
    // Handle comma as thousands separator
    if (cleaned.includes(',')) {
      const parts = cleaned.split(',');
      if (parts.length === 2 && parts[1].length <= 2) {
        // Comma as decimal separator (European format)
        cleaned = parts[0] + '.' + parts[1];
      } else {
        // Comma as thousands separator
        cleaned = cleaned.replace(/,/g, '');
      }
    }
    
    const parsed = parseFloat(cleaned);
    return isNaN(parsed) ? 0 : Math.round(parsed * 100) / 100;
  }

  private calculatePatternConfidence(pattern: RegExp, match: RegExpMatchArray, fullText: string): number {
    let confidence = 0.7; // Base confidence
    
    // Increase confidence based on pattern specificity
    const patternStr = pattern.toString();
    if (patternStr.includes('Box')) confidence += 0.1;
    if (patternStr.includes('\\d')) confidence += 0.05;
    if (patternStr.includes('\\$')) confidence += 0.05;
    
    // Increase confidence based on match quality
    const matchedText = match[0];
    if (matchedText.includes('$')) confidence += 0.05;
    if (matchedText.includes('Box')) confidence += 0.1;
    if (/\d{1,3}(,\d{3})*(\.\d{2})?/.test(match[1])) confidence += 0.1;
    
    return Math.min(confidence, 0.95);
  }

  private async crossValidateAndMerge(
    structuredResults: Record<string, FieldExtractionResult>,
    ocrResults: Record<string, FieldExtractionResult>,
    documentType: DocumentType
  ): Promise<Record<string, FieldExtractionResult>> {
    const mergedResults: Record<string, FieldExtractionResult> = {};
    
    // Get all unique field names
    const allFields = new Set([
      ...Object.keys(structuredResults),
      ...Object.keys(ocrResults)
    ]);
    
    for (const fieldName of allFields) {
      const structuredResult = structuredResults[fieldName];
      const ocrResult = ocrResults[fieldName];
      
      if (structuredResult && ocrResult) {
        // Both methods found the field - cross-validate
        const validation = this.crossValidateField(structuredResult, ocrResult, fieldName);
        mergedResults[fieldName] = validation;
      } else if (structuredResult && structuredResult.confidence >= this.MIN_CONFIDENCE_THRESHOLD) {
        // Only structured found it with good confidence
        mergedResults[fieldName] = structuredResult;
      } else if (ocrResult && ocrResult.confidence >= this.MIN_CONFIDENCE_THRESHOLD) {
        // Only OCR found it with good confidence
        mergedResults[fieldName] = ocrResult;
      } else if (structuredResult) {
        // Use structured even with low confidence
        mergedResults[fieldName] = {
          ...structuredResult,
          source: 'structured'
        };
      } else if (ocrResult) {
        // Use OCR even with low confidence
        mergedResults[fieldName] = {
          ...ocrResult,
          source: 'enhanced_ocr'
        };
      }
    }
    
    console.log(`🔍 [Azure DI Enhanced] Cross-validation: ${Object.keys(mergedResults).length} fields merged`);
    return mergedResults;
  }

  private crossValidateField(
    structuredResult: FieldExtractionResult,
    ocrResult: FieldExtractionResult,
    fieldName: string
  ): FieldExtractionResult {
    // For numeric fields, compare values
    if (typeof structuredResult.value === 'number' && typeof ocrResult.value === 'number') {
      const structuredValue = structuredResult.value;
      const ocrValue = ocrResult.value;
      const percentDiff = Math.abs(structuredValue - ocrValue) / Math.max(structuredValue, ocrValue, 1);
      
      if (percentDiff < 0.05) {
        // Values are very close - use higher confidence
        return structuredResult.confidence > ocrResult.confidence ? 
          { ...structuredResult, source: 'cross_validated' } :
          { ...ocrResult, source: 'cross_validated' };
      } else if (percentDiff < 0.2) {
        // Values are somewhat close - flag for review but use higher confidence
        const result = structuredResult.confidence > ocrResult.confidence ? structuredResult : ocrResult;
        return {
          ...result,
          source: 'cross_validated',
          confidence: Math.max(result.confidence - 0.1, 0.5)
        };
      } else {
        // Values are very different - use OCR if it seems more reasonable
        if (this.isReasonableValue(ocrValue, fieldName) && !this.isReasonableValue(structuredValue, fieldName)) {
          return {
            ...ocrResult,
            source: 'cross_validated',
            originalValue: structuredValue,
            correctionApplied: `Chose OCR value ${ocrValue} over structured value ${structuredValue}`
          };
        } else {
          return {
            ...structuredResult,
            source: 'cross_validated',
            confidence: Math.max(structuredResult.confidence - 0.2, 0.3)
          };
        }
      }
    }
    
    // For string fields, use the one with higher confidence
    return structuredResult.confidence > ocrResult.confidence ? 
      { ...structuredResult, source: 'cross_validated' } :
      { ...ocrResult, source: 'cross_validated' };
  }

  private isReasonableValue(value: number, fieldName: string): boolean {
    if (value <= 0) return false;
    
    // Field-specific reasonableness checks
    switch (fieldName) {
      case 'interestIncome':
        return value >= 0.01 && value <= 1000000; // $0.01 to $1M
      case 'federalTaxWithheld':
        return value >= 0.01 && value <= 500000; // $0.01 to $500K
      case 'earlyWithdrawalPenalty':
        return value >= 10 && value <= 50000; // $10 to $50K (penalties are usually significant)
      case 'foreignTaxPaid':
        return value >= 1 && value <= 100000; // $1 to $100K
      case 'taxExemptInterest':
        return value >= 0.01 && value <= 500000; // $0.01 to $500K
      case 'investmentExpenses':
        return value >= 1 && value <= 100000; // $1 to $100K
      case 'marketDiscount':
      case 'bondPremium':
      case 'bondPremiumTreasury':
      case 'bondPremiumTaxExempt':
        return value >= 1 && value <= 50000; // $1 to $50K
      default:
        return value >= 0.01 && value <= 1000000;
    }
  }

  private async applyAutoCorrections(
    results: Record<string, FieldExtractionResult>,
    documentType: DocumentType
  ): Promise<Record<string, FieldExtractionResult>> {
    if (documentType !== 'FORM_1099_INT') return results;
    
    const correctedResults = { ...results };
    let correctionsApplied = 0;
    
    // Auto-correction rules for 1099-INT
    const interestIncome = results.interestIncome?.value as number || 0;
    
    // Rule 1: Early withdrawal penalty should be reasonable relative to interest
    if (results.earlyWithdrawalPenalty && interestIncome > 1000) {
      const penalty = results.earlyWithdrawalPenalty.value as number;
      if (penalty < 10 && penalty > 0) {
        // Likely OCR error - penalty too small
        const correctedPenalty = penalty * 1000; // Common OCR error: missing zeros
        if (correctedPenalty <= interestIncome * 0.1) { // Penalty shouldn't exceed 10% of interest
          correctedResults.earlyWithdrawalPenalty = {
            ...results.earlyWithdrawalPenalty,
            value: correctedPenalty,
            source: 'auto_corrected',
            originalValue: penalty,
            correctionApplied: `Auto-corrected penalty from $${penalty} to $${correctedPenalty} (likely missing zeros)`
          };
          correctionsApplied++;
        }
      }
    }
    
    // Rule 2: Foreign tax paid should be reasonable
    if (results.foreignTaxPaid && interestIncome > 1000) {
      const foreignTax = results.foreignTaxPaid.value as number;
      if (foreignTax < 20 && foreignTax > 0) {
        // Likely OCR error - foreign tax too small
        const correctedForeignTax = foreignTax * 100; // Common OCR error: decimal point misplacement
        if (correctedForeignTax <= interestIncome * 0.3) { // Foreign tax shouldn't exceed 30% of interest
          correctedResults.foreignTaxPaid = {
            ...results.foreignTaxPaid,
            value: correctedForeignTax,
            source: 'auto_corrected',
            originalValue: foreignTax,
            correctionApplied: `Auto-corrected foreign tax from $${foreignTax} to $${correctedForeignTax} (likely decimal error)`
          };
          correctionsApplied++;
        }
      }
    }
    
    // Rule 3: Cross-field validation for Treasury interest
    if (results.interestOnUSavingsBonds && results.interestIncome) {
      const treasuryInterest = results.interestOnUSavingsBonds.value as number;
      const totalInterest = results.interestIncome.value as number;
      
      if (treasuryInterest > totalInterest) {
        // Treasury interest can't exceed total interest
        correctedResults.interestOnUSavingsBonds = {
          ...results.interestOnUSavingsBonds,
          value: Math.min(treasuryInterest, totalInterest * 0.8),
          source: 'auto_corrected',
          originalValue: treasuryInterest,
          correctionApplied: `Auto-corrected Treasury interest to not exceed total interest`
        };
        correctionsApplied++;
      }
    }
    
    if (correctionsApplied > 0) {
      console.log(`🔧 [Azure DI Enhanced] Applied ${correctionsApplied} auto-corrections`);
    }
    
    return correctedResults;
  }

  private async calculateFinalMetrics(
    results: Record<string, FieldExtractionResult>,
    documentType: DocumentType,
    startTime: number
  ): Promise<ExtractedFieldData> {
    const finalData: ExtractedFieldData = {};
    const warnings: string[] = [];
    let totalConfidence = 0;
    let fieldCount = 0;
    let structuredFields = 0;
    let ocrFallbacks = 0;
    let autoCorrections = 0;
    
    // Convert results to final format
    for (const [fieldName, result] of Object.entries(results)) {
      finalData[fieldName] = result.value;
      totalConfidence += result.confidence;
      fieldCount++;
      
      // Track metrics
      switch (result.source) {
        case 'structured':
          structuredFields++;
          break;
        case 'enhanced_ocr':
        case 'ocr_fallback':
          ocrFallbacks++;
          break;
        case 'auto_corrected':
          autoCorrections++;
          if (result.correctionApplied) {
            warnings.push(result.correctionApplied);
          }
          break;
      }
      
      // Add warnings for low confidence fields
      if (result.confidence < this.MIN_CONFIDENCE_THRESHOLD) {
        warnings.push(`Low confidence (${Math.round(result.confidence * 100)}%) for field: ${fieldName}`);
      }
    }
    
    // Check for missing critical fields
    if (documentType === 'FORM_1099_INT') {
      for (const criticalField of this.CRITICAL_FIELDS_1099_INT) {
        if (!finalData[criticalField]) {
          warnings.push(`Missing critical field: ${criticalField}`);
        }
      }
    }
    
    // Calculate final confidence
    const avgConfidence = fieldCount > 0 ? totalConfidence / fieldCount : 0;
    finalData.confidence = Math.round(avgConfidence * 100) / 100;
    finalData.extractionWarnings = warnings;
    
    // Add processing metrics
    finalData.processingMetrics = {
      structuredFieldsExtracted: structuredFields,
      ocrFallbacksUsed: ocrFallbacks,
      autoCorrectionsApplied: autoCorrections,
      validationWarnings: warnings.length,
      totalProcessingTime: Date.now() - startTime
    };
    
    return finalData;
  }

  // Additional helper methods for other document types
  private extractEnhanced1099MiscFields(ocrText: string): Record<string, FieldExtractionResult> {
    // Implementation for 1099-MISC (similar structure to 1099-INT)
    return {};
  }

  private extractEnhanced1099NecFields(ocrText: string): Record<string, FieldExtractionResult> {
    // Implementation for 1099-NEC
    return {};
  }

  private extractEnhanced1099DivFields(ocrText: string): Record<string, FieldExtractionResult> {
    // Implementation for 1099-DIV
    return {};
  }

  private extractEnhancedW2Fields(ocrText: string): Record<string, FieldExtractionResult> {
    // Implementation for W-2
    return {};
  }

  private getFieldMappingsForDocumentType(documentType: DocumentType): Record<string, string> {
    switch (documentType) {
      case 'FORM_1099_INT':
        return {
          'Payer.Name': 'payerName',
          'Payer.TIN': 'payerTIN',
          'Payer.Address': 'payerAddress',
          'Recipient.Name': 'recipientName',
          'Recipient.TIN': 'recipientTIN',
          'Recipient.Address': 'recipientAddress',
          'InterestIncome': 'interestIncome',
          'EarlyWithdrawalPenalty': 'earlyWithdrawalPenalty',
          'InterestOnUSTreasuryObligations': 'interestOnUSavingsBonds',
          'FederalIncomeTaxWithheld': 'federalTaxWithheld',
          'InvestmentExpenses': 'investmentExpenses',
          'ForeignTaxPaid': 'foreignTaxPaid',
          'TaxExemptInterest': 'taxExemptInterest',
          'SpecifiedPrivateActivityBondInterest': 'specifiedPrivateActivityBond',
          'MarketDiscount': 'marketDiscount',
          'BondPremium': 'bondPremium',
          'BondPremiumOnTreasuryObligations': 'bondPremiumTreasury',
          'BondPremiumOnTaxExemptBond': 'bondPremiumTaxExempt'
        };
      default:
        return {};
    }
  }

  private parseFieldValue(value: any): any {
    if (value === null || value === undefined) return null;
    if (typeof value === 'object' && value.amount !== undefined) {
      return this.parseEnhancedAmount(value.amount.toString());
    }
    if (typeof value === 'string' && /^\d+\.?\d*$/.test(value)) {
      return this.parseEnhancedAmount(value);
    }
    return value;
  }

  private handleExtractionError(error: any, documentType: DocumentType, startTime: number): ExtractedFieldData {
    console.error('❌ [Azure DI Enhanced] Extraction failed:', error);
    
    return {
      confidence: 0,
      extractionWarnings: [`Extraction failed: ${error?.message || 'Unknown error'}`],
      processingMetrics: {
        structuredFieldsExtracted: 0,
        ocrFallbacksUsed: 0,
        autoCorrectionsApplied: 0,
        validationWarnings: 1,
        totalProcessingTime: Date.now() - startTime
      }
    };
  }
}

// Factory function for backward compatibility
export function getAzureDocumentIntelligenceService(): AzureDocumentIntelligenceServiceEnhanced {
  const config: AzureDocumentIntelligenceConfig = {
    endpoint: process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT || '',
    apiKey: process.env.AZURE_DOCUMENT_INTELLIGENCE_API_KEY || ''
  };
  
  if (!config.endpoint || !config.apiKey) {
    throw new Error('Azure Document Intelligence configuration is missing. Please set AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT and AZURE_DOCUMENT_INTELLIGENCE_API_KEY environment variables.');
  }
  
  return new AzureDocumentIntelligenceServiceEnhanced(config);
}
