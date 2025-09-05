import { DocumentAnalysisClient, AzureKeyCredential } from "@azure/ai-form-recognizer";
import { DocumentType } from "@prisma/client";
import { readFile } from "fs/promises";

export interface AzureDocumentIntelligenceConfig {
  endpoint: string;
  apiKey: string;
}

export interface ExtractedFieldData {
  [key: string]: string | number | DocumentType | number[] | string[] | { [key: string]: number } | undefined;
  correctedDocumentType?: DocumentType;
  fullText?: string;
  confidence?: number;
  extractionWarnings?: string[];
  fieldConfidences?: { [key: string]: number };
}

export interface FieldExtractionResult {
  value: any;
  confidence: number;
  source: 'structured' | 'ocr_primary' | 'ocr_fallback' | 'regex_enhanced';
  method?: string;
}

export class AzureDocumentIntelligenceService {
  private client: DocumentAnalysisClient;
  private config: AzureDocumentIntelligenceConfig;
  
  // Comprehensive 1099-INT field definitions with all required fields
  private readonly FORM_1099_INT_FIELDS = {
    // Box 1: Interest Income (Payer must complete)
    interestIncome: {
      boxNumber: 1,
      label: 'Interest Income',
      required: true,
      patterns: [
        // Primary patterns - exact box matching
        /(?:Box\s*1|1\.?\s*Interest\s+income)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Interest\s+income[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /1\s+Interest\s+income[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        // OCR error patterns (O->0, l->1, I->1)
        /(?:Box\s*[1lI]|[1lI]\.?\s*Interest\s+income)[:\s]*\$?\s*([0-9,Ol]+\.?\d{0,2})/i,
        // Position-based patterns for table layouts
        /(?:^|\n)\s*1[:\s]*([0-9,]+\.?\d{0,2})/m,
        // Context-based patterns
        /Total\s+interest\s+income[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        // Fallback patterns for poor OCR
        /(?:lnterest|Interest).*?income.*?([0-9,]+\.?\d{0,2})/i
      ]
    },
    
    // Box 2: Early Withdrawal Penalty (CRITICAL - was showing $3 instead of $2,000)
    earlyWithdrawalPenalty: {
      boxNumber: 2,
      label: 'Early Withdrawal Penalty',
      required: false,
      patterns: [
        // Primary patterns with enhanced matching
        /(?:Box\s*2|2\.?\s*Early\s+withdrawal\s+penalty)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Early\s+withdrawal\s+penalty[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /2\s+Early\s+withdrawal[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        // OCR error patterns
        /(?:Box\s*[2Z]|[2Z]\.?\s*Early\s+withdrawal)[:\s]*\$?\s*([0-9,OZ]+\.?\d{0,2})/i,
        // Enhanced patterns for common OCR errors
        /Early.*?withdrawal.*?penalty.*?([0-9,]+\.?\d{0,2})/i,
        /(?:Eariy|Early).*?withdrawal.*?([0-9,]+\.?\d{0,2})/i,
        // Position-based patterns
        /(?:^|\n)\s*2[:\s]*([0-9,]+\.?\d{0,2})/m,
        // Table-based patterns with better context
        /withdrawal.*?penalty.*?([0-9,]+\.\d{2})/i,
        // Penalty-specific patterns
        /penalty.*?early.*?withdrawal.*?([0-9,]+\.?\d{0,2})/i
      ]
    },
    
    // Box 3: Interest on U.S. Treasury Obligations (MISSING FIELD - $5,000)
    interestOnUSavingsBonds: {
      boxNumber: 3,
      label: 'Interest on U.S. Treasury Obligations',
      required: false,
      patterns: [
        // Primary patterns
        /(?:Box\s*3|3\.?\s*Interest\s+on\s+U\.?S\.?\s+Treasury)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Interest\s+on\s+U\.?S\.?\s+Treasury\s+obligations[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /3\s+Interest\s+on\s+U\.?S\.?[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        // OCR error patterns
        /(?:Box\s*[3E]|[3E]\.?\s*Interest\s+on\s+U\.?S\.?)[:\s]*\$?\s*([0-9,OE]+\.?\d{0,2})/i,
        // Enhanced patterns for Treasury obligations
        /Treasury\s+obligations[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /U\.?S\.?\s+Treasury.*?([0-9,]+\.?\d{0,2})/i,
        /Interest.*?Treasury.*?obligations.*?([0-9,]+\.?\d{0,2})/i,
        // Position-based patterns
        /(?:^|\n)\s*3[:\s]*([0-9,]+\.?\d{0,2})/m,
        // Fallback patterns
        /(?:Treasury|Treasuiy).*?obligations.*?([0-9,]+\.?\d{0,2})/i,
        // Savings bonds patterns
        /(?:U\.?S\.?\s+)?savings\s+bonds.*?([0-9,]+\.?\d{0,2})/i
      ]
    },
    
    // Box 4: Federal Income Tax Withheld
    federalTaxWithheld: {
      boxNumber: 4,
      label: 'Federal Income Tax Withheld',
      required: false,
      patterns: [
        // Primary patterns
        /(?:Box\s*4|4\.?\s*Federal\s+income\s+tax\s+withheld)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Federal\s+income\s+tax\s+withheld[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /4\s+Federal\s+income\s+tax[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        // OCR error patterns
        /(?:Box\s*[4A]|[4A]\.?\s*Federal\s+income)[:\s]*\$?\s*([0-9,OA]+\.?\d{0,2})/i,
        // Enhanced patterns
        /Tax\s+withheld[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Backup\s+withholding[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Federal.*?tax.*?withheld.*?([0-9,]+\.?\d{0,2})/i,
        // Position-based patterns
        /(?:^|\n)\s*4[:\s]*([0-9,]+\.?\d{0,2})/m
      ]
    },
    
    // Box 5: Investment Expenses
    investmentExpenses: {
      boxNumber: 5,
      label: 'Investment Expenses',
      required: false,
      patterns: [
        // Primary patterns
        /(?:Box\s*5|5\.?\s*Investment\s+expenses)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Investment\s+expenses[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /5\s+Investment\s+expenses[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        // OCR error patterns
        /(?:Box\s*[5S]|[5S]\.?\s*Investment)[:\s]*\$?\s*([0-9,OS]+\.?\d{0,2})/i,
        // Enhanced patterns
        /REMIC\s+expenses[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Investment.*?expenses.*?([0-9,]+\.?\d{0,2})/i,
        // Position-based patterns
        /(?:^|\n)\s*5[:\s]*([0-9,]+\.?\d{0,2})/m
      ]
    },
    
    // Box 6: Foreign Tax Paid (CRITICAL - was showing $8 instead of higher amount)
    foreignTaxPaid: {
      boxNumber: 6,
      label: 'Foreign Tax Paid',
      required: false,
      patterns: [
        // Primary patterns with enhanced matching
        /(?:Box\s*6|6\.?\s*Foreign\s+tax\s+paid)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Foreign\s+tax\s+paid[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /6\s+Foreign\s+tax[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        // OCR error patterns
        /(?:Box\s*[6G]|[6G]\.?\s*Foreign\s+tax)[:\s]*\$?\s*([0-9,OG]+\.?\d{0,2})/i,
        // Enhanced patterns for better accuracy
        /Foreign.*?tax.*?paid.*?([0-9,]+\.?\d{0,2})/i,
        /(?:Foreign|Forelgn).*?tax.*?([0-9,]+\.?\d{0,2})/i,
        // Position-based patterns
        /(?:^|\n)\s*6[:\s]*([0-9,]+\.?\d{0,2})/m,
        // Table-based patterns
        /tax.*?paid.*?([0-9,]+\.\d{2})/i,
        // International tax patterns
        /foreign.*?withholding.*?([0-9,]+\.?\d{0,2})/i
      ]
    },
    
    // Box 7: Foreign Country or U.S. Territory
    foreignCountry: {
      boxNumber: 7,
      label: 'Foreign Country or U.S. Territory',
      required: false,
      patterns: [
        /(?:Box\s*7|7\.?\s*Foreign\s+country)[:\s]*([A-Za-z\s]+)/i,
        /Foreign\s+country[:\s]*([A-Za-z\s]+)/i,
        /(?:^|\n)\s*7[:\s]*([A-Za-z\s]+)/m
      ]
    },
    
    // Box 8: Tax-Exempt Interest
    taxExemptInterest: {
      boxNumber: 8,
      label: 'Tax-Exempt Interest',
      required: false,
      patterns: [
        // Primary patterns
        /(?:Box\s*8|8\.?\s*Tax-exempt\s+interest)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Tax-exempt\s+interest[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /8\s+Tax-exempt[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        // OCR error patterns
        /(?:Box\s*[8B]|[8B]\.?\s*Tax-exempt)[:\s]*\$?\s*([0-9,OB]+\.?\d{0,2})/i,
        // Enhanced patterns
        /Municipal\s+bond\s+interest[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Tax.*?exempt.*?interest.*?([0-9,]+\.?\d{0,2})/i,
        // Position-based patterns
        /(?:^|\n)\s*8[:\s]*([0-9,]+\.?\d{0,2})/m
      ]
    },
    
    // Box 9: Specified Private Activity Bond Interest
    specifiedPrivateActivityBond: {
      boxNumber: 9,
      label: 'Specified Private Activity Bond Interest',
      required: false,
      patterns: [
        // Primary patterns
        /(?:Box\s*9|9\.?\s*Specified\s+private\s+activity)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Specified\s+private\s+activity\s+bond[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /9\s+Specified\s+private[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        // OCR error patterns
        /(?:Box\s*[9g]|[9g]\.?\s*Specified\s+private)[:\s]*\$?\s*([0-9,Og]+\.?\d{0,2})/i,
        // Enhanced patterns
        /Private\s+activity\s+bond[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Specified.*?private.*?activity.*?([0-9,]+\.?\d{0,2})/i,
        // Position-based patterns
        /(?:^|\n)\s*9[:\s]*([0-9,]+\.?\d{0,2})/m
      ]
    },
    
    // Box 10: Market Discount (MISSING FIELD - $800)
    marketDiscount: {
      boxNumber: 10,
      label: 'Market Discount',
      required: false,
      patterns: [
        // Primary patterns
        /(?:Box\s*10|10\.?\s*Market\s+discount)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Market\s+discount[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /10\s+Market\s+discount[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        // OCR error patterns
        /(?:Box\s*[1l][0O]|[1l][0O]\.?\s*Market)[:\s]*\$?\s*([0-9,OlI]+\.?\d{0,2})/i,
        // Enhanced patterns
        /Market.*?discount.*?([0-9,]+\.?\d{0,2})/i,
        /Accrued.*?market.*?discount.*?([0-9,]+\.?\d{0,2})/i,
        // Position-based patterns
        /(?:^|\n)\s*10[:\s]*([0-9,]+\.?\d{0,2})/m,
        // Bond-related patterns
        /bond.*?market.*?discount.*?([0-9,]+\.?\d{0,2})/i
      ]
    },
    
    // Box 11: Bond Premium (MISSING FIELD - $700)
    bondPremium: {
      boxNumber: 11,
      label: 'Bond Premium',
      required: false,
      patterns: [
        // Primary patterns
        /(?:Box\s*11|11\.?\s*Bond\s+premium)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Bond\s+premium[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /11\s+Bond\s+premium[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        // OCR error patterns
        /(?:Box\s*[1l][1l]|[1l][1l]\.?\s*Bond)[:\s]*\$?\s*([0-9,OlI]+\.?\d{0,2})/i,
        // Enhanced patterns
        /Bond.*?premium.*?([0-9,]+\.?\d{0,2})/i,
        /Premium.*?amortization.*?([0-9,]+\.?\d{0,2})/i,
        // Position-based patterns
        /(?:^|\n)\s*11[:\s]*([0-9,]+\.?\d{0,2})/m,
        // Amortization patterns
        /amortizable.*?bond.*?premium.*?([0-9,]+\.?\d{0,2})/i
      ]
    },
    
    // Box 12: Bond Premium on Treasury Obligations
    bondPremiumTreasury: {
      boxNumber: 12,
      label: 'Bond Premium on Treasury Obligations',
      required: false,
      patterns: [
        // Primary patterns
        /(?:Box\s*12|12\.?\s*Bond\s+premium\s+on\s+Treasury)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Bond\s+premium\s+on\s+Treasury\s+obligations[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /12\s+Bond\s+premium\s+on\s+Treasury[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        // OCR error patterns
        /(?:Box\s*[1l][2Z]|[1l][2Z]\.?\s*Bond\s+premium\s+on)[:\s]*\$?\s*([0-9,OlIZ]+\.?\d{0,2})/i,
        // Enhanced patterns
        /Bond.*?premium.*?Treasury.*?([0-9,]+\.?\d{0,2})/i,
        /Treasury.*?bond.*?premium.*?([0-9,]+\.?\d{0,2})/i,
        // Position-based patterns
        /(?:^|\n)\s*12[:\s]*([0-9,]+\.?\d{0,2})/m
      ]
    },
    
    // Box 13: Bond Premium on Tax-Exempt Bond
    bondPremiumTaxExempt: {
      boxNumber: 13,
      label: 'Bond Premium on Tax-Exempt Bond',
      required: false,
      patterns: [
        // Primary patterns
        /(?:Box\s*13|13\.?\s*Bond\s+premium\s+on\s+tax-exempt)[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /Bond\s+premium\s+on\s+tax-exempt\s+bond[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        /13\s+Bond\s+premium\s+on\s+tax-exempt[:\s]*\$?\s*([0-9,]+\.?\d{0,2})/i,
        // OCR error patterns
        /(?:Box\s*[1l][3E]|[1l][3E]\.?\s*Bond\s+premium\s+on\s+tax)[:\s]*\$?\s*([0-9,OlIE]+\.?\d{0,2})/i,
        // Enhanced patterns
        /Bond.*?premium.*?tax.*?exempt.*?([0-9,]+\.?\d{0,2})/i,
        /Tax.*?exempt.*?bond.*?premium.*?([0-9,]+\.?\d{0,2})/i,
        // Position-based patterns
        /(?:^|\n)\s*13[:\s]*([0-9,]+\.?\d{0,2})/m
      ]
    },
    
    // Box 14: Tax-Exempt and Tax Credit Bond CUSIP No.
    cusipNumber: {
      boxNumber: 14,
      label: 'Tax-Exempt and Tax Credit Bond CUSIP No.',
      required: false,
      patterns: [
        /(?:Box\s*14|14\.?\s*CUSIP)[:\s]*([A-Z0-9]+)/i,
        /CUSIP[:\s]*([A-Z0-9]+)/i,
        /(?:^|\n)\s*14[:\s]*([A-Z0-9]+)/m
      ]
    }
  };

  constructor(config: AzureDocumentIntelligenceConfig) {
    this.config = config;
    this.client = new DocumentAnalysisClient(
      this.config.endpoint,
      new AzureKeyCredential(this.config.apiKey)
    );
  }

  /**
   * Main extraction method that prioritizes OCR-based extraction
   */
  async extractDataFromDocument(
    documentPathOrBuffer: string | Buffer,
    documentType: string
  ): Promise<ExtractedFieldData> {
    try {
      console.log('🔍 [Azure DI] Starting aggressive OCR-based extraction...');
      
      // Get document buffer
      const documentBuffer = typeof documentPathOrBuffer === 'string' 
        ? await readFile(documentPathOrBuffer)
        : documentPathOrBuffer;
      
      // Always use OCR-first approach for 1099-INT
      const poller = await this.client.beginAnalyzeDocument('prebuilt-read', documentBuffer);
      const result = await poller.pollUntilDone();
      
      console.log('✅ [Azure DI] OCR analysis completed');
      
      // Extract full text
      const fullText = result.content || '';
      
      // For 1099-INT, use aggressive OCR-based extraction
      if (documentType === 'FORM_1099_INT') {
        return this.extract1099IntFieldsAggressively(fullText);
      }
      
      // Fallback for other document types
      return this.extractGenericFields(fullText);
      
    } catch (error: any) {
      console.error('❌ [Azure DI] Processing error:', error);
      return {
        extractionWarnings: [`Azure Document Intelligence processing failed: ${error?.message || 'Unknown error'}`],
        confidence: 0
      };
    }
  }

  /**
   * Aggressive 1099-INT field extraction using comprehensive regex patterns
   */
  private extract1099IntFieldsAggressively(ocrText: string): ExtractedFieldData {
    console.log('🔍 [Azure DI] Starting aggressive 1099-INT field extraction...');
    
    const extractedData: ExtractedFieldData = {
      fullText: ocrText,
      extractionWarnings: [],
      confidence: 0.9, // High confidence for OCR-based extraction
      fieldConfidences: {}
    };
    
    // Preprocess OCR text to handle common OCR errors
    const preprocessedText = this.preprocessOCRTextForExtraction(ocrText);
    
    let fieldsExtracted = 0;
    const totalFields = Object.keys(this.FORM_1099_INT_FIELDS).length;
    
    // Extract each field using comprehensive regex patterns
    for (const [fieldName, fieldConfig] of Object.entries(this.FORM_1099_INT_FIELDS)) {
      const extractedValue = this.extractFieldByRegex(preprocessedText, fieldConfig.patterns, fieldName);
      
      if (extractedValue !== null && extractedValue !== undefined) {
        // Handle different field types
        if (fieldName === 'foreignCountry' || fieldName === 'cusipNumber') {
          extractedData[fieldName] = String(extractedValue).trim();
          extractedData.fieldConfidences![fieldName] = 0.85;
        } else {
          const numericValue = this.parseAmountRobustly(extractedValue);
          if (numericValue > 0) {
            extractedData[fieldName] = numericValue;
            extractedData.fieldConfidences![fieldName] = 0.9;
            fieldsExtracted++;
            console.log(`✅ [Azure DI] Extracted ${fieldName}: $${numericValue} (Box ${fieldConfig.boxNumber})`);
          }
        }
      }
    }
    
    // Extract personal information
    const personalInfo = this.extractPersonalInformation(preprocessedText);
    Object.assign(extractedData, personalInfo);
    
    // Calculate confidence based on extraction success
    const extractionRate = fieldsExtracted / totalFields;
    extractedData.confidence = Math.max(0.5, 0.7 + (extractionRate * 0.3));
    
    // Add warnings for missing critical fields
    const warnings: string[] = [];
    const criticalFields = ['interestIncome', 'earlyWithdrawalPenalty', 'interestOnUSavingsBonds', 'foreignTaxPaid', 'marketDiscount', 'bondPremium'];
    
    for (const field of criticalFields) {
      if (!extractedData[field] || extractedData[field] === 0) {
        warnings.push(`Missing or zero value for critical field: ${field}`);
      }
    }
    
    extractedData.extractionWarnings = warnings;
    
    console.log(`✅ [Azure DI] Aggressive extraction completed. Fields extracted: ${fieldsExtracted}/${totalFields}, Confidence: ${extractedData.confidence}`);
    
    return extractedData;
  }

  /**
   * Extract field value using multiple regex patterns with fallback
   */
  private extractFieldByRegex(text: string, patterns: RegExp[], fieldName: string): any {
    for (let i = 0; i < patterns.length; i++) {
      const pattern = patterns[i];
      const match = text.match(pattern);
      
      if (match && match[1]) {
        const value = match[1].trim();
        console.log(`🎯 [Azure DI] Pattern ${i + 1} matched for ${fieldName}: "${value}"`);
        return value;
      }
    }
    
    // Additional fallback: search for field name + amount pattern
    const fallbackPattern = new RegExp(`${fieldName.replace(/([A-Z])/g, '\\s*$1').toLowerCase()}.*?([0-9,]+\\.?\\d{0,2})`, 'i');
    const fallbackMatch = text.match(fallbackPattern);
    if (fallbackMatch && fallbackMatch[1]) {
      console.log(`🔄 [Azure DI] Fallback pattern matched for ${fieldName}: "${fallbackMatch[1]}"`);
      return fallbackMatch[1];
    }
    
    return null;
  }

  /**
   * Robust amount parsing with OCR error correction
   */
  private parseAmountRobustly(value: any): number {
    if (typeof value === 'number') {
      return value;
    }
    
    if (typeof value === 'string') {
      // Remove currency symbols, commas, and whitespace
      let cleanValue = value.replace(/[$,\s]/g, '').trim();
      
      // Handle empty or non-numeric strings
      if (!cleanValue || cleanValue === '' || cleanValue === '-' || cleanValue === 'N/A') {
        return 0;
      }
      
      // OCR error corrections for numbers - comprehensive mapping
      const ocrCorrections = {
        'O': '0', 'o': '0',  // O -> 0
        'l': '1', 'I': '1', '|': '1',  // l, I, | -> 1
        'S': '5', 's': '5',  // S -> 5
        'Z': '2', 'z': '2',  // Z -> 2
        'G': '6', 'g': '6',  // G -> 6 (sometimes)
        'B': '8', 'b': '8',  // B -> 8 (sometimes)
        'T': '7', 't': '7',  // T -> 7 (sometimes)
        'A': '4', 'a': '4'   // A -> 4 (sometimes)
      };
      
      // Apply OCR corrections
      for (const [wrong, correct] of Object.entries(ocrCorrections)) {
        cleanValue = cleanValue.replace(new RegExp(wrong, 'g'), correct);
      }
      
      // Parse the cleaned value
      const parsed = parseFloat(cleanValue);
      
      // Return 0 for invalid numbers
      return isNaN(parsed) ? 0 : parsed;
    }
    
    return 0;
  }

  /**
   * Preprocess OCR text to handle common OCR errors and improve matching
   */
  private preprocessOCRTextForExtraction(ocrText: string): string {
    let processed = ocrText;
    
    // Common OCR character corrections
    const corrections = [
      // Numbers in context
      { from: /O(?=\d)/g, to: '0' }, // O followed by digit -> 0
      { from: /(?<=\d)O/g, to: '0' }, // O preceded by digit -> 0
      { from: /l(?=\d)/g, to: '1' }, // l followed by digit -> 1
      { from: /(?<=\d)l/g, to: '1' }, // l preceded by digit -> 1
      { from: /I(?=\d)/g, to: '1' }, // I followed by digit -> 1
      { from: /(?<=\d)I/g, to: '1' }, // I preceded by digit -> 1
      { from: /S(?=\d)/g, to: '5' }, // S followed by digit -> 5
      { from: /(?<=\d)S/g, to: '5' }, // S preceded by digit -> 5
      { from: /Z(?=\d)/g, to: '2' }, // Z followed by digit -> 2
      { from: /(?<=\d)Z/g, to: '2' }, // Z preceded by digit -> 2
      { from: /G(?=\d)/g, to: '6' }, // G followed by digit -> 6
      { from: /(?<=\d)G/g, to: '6' }, // G preceded by digit -> 6
      { from: /B(?=\d)/g, to: '8' }, // B followed by digit -> 8
      { from: /(?<=\d)B/g, to: '8' }, // B preceded by digit -> 8
      { from: /g(?=\d)/g, to: '9' }, // g followed by digit -> 9
      { from: /(?<=\d)g/g, to: '9' }, // g preceded by digit -> 9
      
      // Common word corrections
      { from: /lnterest/g, to: 'Interest' },
      { from: /Eariy/g, to: 'Early' },
      { from: /Forelgn/g, to: 'Foreign' },
      { from: /Treasuiy/g, to: 'Treasury' },
      { from: /Wlthheld/g, to: 'Withheld' },
      { from: /Pald/g, to: 'Paid' },
      { from: /Dlscount/g, to: 'Discount' },
      { from: /Premlum/g, to: 'Premium' },
      { from: /Obllgations/g, to: 'Obligations' },
      { from: /Savlngs/g, to: 'Savings' },
      { from: /Wlthdrawal/g, to: 'Withdrawal' },
      { from: /Penaliy/g, to: 'Penalty' },
      { from: /Exempi/g, to: 'Exempt' },
      { from: /Prlvate/g, to: 'Private' },
      { from: /Actlvity/g, to: 'Activity' },
      { from: /Markel/g, to: 'Market' },
      { from: /Amortlzation/g, to: 'Amortization' }
    ];
    
    for (const correction of corrections) {
      processed = processed.replace(correction.from, correction.to);
    }
    
    return processed;
  }

  /**
   * Extract personal information (payer, recipient) from OCR text
   */
  private extractPersonalInformation(ocrText: string): Partial<ExtractedFieldData> {
    const personalInfo: Partial<ExtractedFieldData> = {};
    
    // Recipient name patterns
    const recipientNamePatterns = [
      /(?:RECIPIENT'S?\s+name|Recipient'?s?\s+name)\s*\n([A-Za-z\s]+?)(?:\n|$)/i,
      /(?:RECIPIENT'S?\s+NAME|Recipient'?s?\s+name)[:\s]+([A-Z\s]+)/i,
      /Recipient[:\s]+([A-Za-z\s]+?)(?:\n|TIN)/i
    ];
    
    for (const pattern of recipientNamePatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1] && match[1].trim().length > 2) {
        personalInfo.recipientName = match[1].trim();
        break;
      }
    }
    
    // Recipient TIN patterns
    const recipientTINPatterns = [
      /(?:RECIPIENT'S?\s+TIN|Recipient'?s?\s+TIN)[:\s]*(\d{3}-?\d{2}-?\d{4}|\d{2}-?\d{7})/i,
      /TIN[:\s]*(\d{3}-?\d{2}-?\d{4}|\d{2}-?\d{7})/i,
      /(?:SSN|Social Security)[:\s]*(\d{3}-?\d{2}-?\d{4})/i
    ];
    
    for (const pattern of recipientTINPatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        personalInfo.recipientTIN = match[1].trim();
        break;
      }
    }
    
    // Payer name patterns
    const payerNamePatterns = [
      /(?:PAYER'S?\s+name|Payer'?s?\s+name)\s*\n([A-Za-z\s&.,]+?)(?:\n|$)/i,
      /(?:PAYER'S?\s+NAME|Payer'?s?\s+name)[:\s]+([A-Z\s&.,]+)/i,
      /Payer[:\s]+([A-Za-z\s&.,]+?)(?:\n|TIN)/i
    ];
    
    for (const pattern of payerNamePatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1] && match[1].trim().length > 2) {
        personalInfo.payerName = match[1].trim();
        break;
      }
    }
    
    // Payer TIN patterns
    const payerTINPatterns = [
      /(?:PAYER'S?\s+TIN|Payer'?s?\s+TIN)[:\s]*(\d{2}-?\d{7})/i,
      /(?:EIN|Employer ID)[:\s]*(\d{2}-?\d{7})/i
    ];
    
    for (const pattern of payerTINPatterns) {
      const match = ocrText.match(pattern);
      if (match && match[1]) {
        personalInfo.payerTIN = match[1].trim();
        break;
      }
    }
    
    return personalInfo;
  }

  /**
   * Generic field extraction for non-1099-INT documents
   */
  private extractGenericFields(ocrText: string): ExtractedFieldData {
    return {
      fullText: ocrText,
      extractionWarnings: ['Generic extraction used - document type not specifically supported'],
      confidence: 0.3
    };
  }

  /**
   * Parse 1099-INT document specifically with enhanced validation
   */
  async parse1099Int(documentPathOrBuffer: string | Buffer): Promise<ExtractedFieldData> {
    console.log('🔍 [Azure DI] Starting specialized 1099-INT parsing...');
    
    const result = await this.extractDataFromDocument(documentPathOrBuffer, 'FORM_1099_INT');
    
    // Apply 1099-INT specific validation and corrections
    return this.validate1099IntFields(result);
  }

  /**
   * Validate and correct 1099-INT fields based on business rules
   */
  private validate1099IntFields(data: ExtractedFieldData): ExtractedFieldData {
    const validatedData = { ...data };
    const warnings = [...(data.extractionWarnings || [])];
    
    // Get numeric values for validation
    const interestIncome = this.parseAmountRobustly(data.interestIncome) || 0;
    const earlyWithdrawalPenalty = this.parseAmountRobustly(data.earlyWithdrawalPenalty) || 0;
    const federalTaxWithheld = this.parseAmountRobustly(data.federalTaxWithheld) || 0;
    const foreignTaxPaid = this.parseAmountRobustly(data.foreignTaxPaid) || 0;
    const interestOnUSavingsBonds = this.parseAmountRobustly(data.interestOnUSavingsBonds) || 0;
    
    // Validation Rule 1: Early withdrawal penalty should not exceed interest income
    if (earlyWithdrawalPenalty > 0 && interestIncome > 0 && earlyWithdrawalPenalty > interestIncome) {
      warnings.push(`Early withdrawal penalty ($${earlyWithdrawalPenalty}) exceeds interest income ($${interestIncome}) - possible OCR error`);
    }
    
    // Validation Rule 2: Federal tax withheld should be reasonable (typically < 50% of interest)
    if (federalTaxWithheld > 0 && interestIncome > 0 && federalTaxWithheld > interestIncome * 0.5) {
      warnings.push(`Federal tax withheld ($${federalTaxWithheld}) seems high compared to interest income ($${interestIncome})`);
    }
    
    // Validation Rule 3: Foreign tax paid should be reasonable
    if (foreignTaxPaid > 0 && interestIncome > 0 && foreignTaxPaid > interestIncome * 0.3) {
      warnings.push(`Foreign tax paid ($${foreignTaxPaid}) seems high compared to interest income ($${interestIncome})`);
    }
    
    // Validation Rule 4: Treasury interest should not exceed total interest
    if (interestOnUSavingsBonds > 0 && interestIncome > 0 && interestOnUSavingsBonds > interestIncome) {
      warnings.push(`Treasury interest ($${interestOnUSavingsBonds}) exceeds total interest income ($${interestIncome}) - possible OCR error`);
    }
    
    // OCR Error Correction: Check for suspiciously small amounts that might be OCR errors
    if (earlyWithdrawalPenalty > 0 && earlyWithdrawalPenalty < 10 && interestIncome > 1000) {
      warnings.push(`Early withdrawal penalty ($${earlyWithdrawalPenalty}) seems unusually small - possible OCR error (expected range: $100-$2000)`);
    }
    
    if (foreignTaxPaid > 0 && foreignTaxPaid < 20 && interestIncome > 1000) {
      warnings.push(`Foreign tax paid ($${foreignTaxPaid}) seems unusually small - possible OCR error (expected range: $50-$500)`);
    }
    
    validatedData.extractionWarnings = warnings;
    
    // Adjust confidence based on validation results
    const validationPenalty = warnings.length * 0.05;
    validatedData.confidence = Math.max(0.1, (data.confidence || 0.5) - validationPenalty);
    
    console.log(`✅ [Azure DI] 1099-INT validation completed. Warnings: ${warnings.length}, Final confidence: ${validatedData.confidence}`);
    
    return validatedData;
  }
}

/**
 * Factory function to create and return an AzureDocumentIntelligenceService instance
 * This function is expected by the route files for dependency injection
 */
export function getAzureDocumentIntelligenceService(): AzureDocumentIntelligenceService {
  const config: AzureDocumentIntelligenceConfig = {
    endpoint: process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT || '',
    apiKey: process.env.AZURE_DOCUMENT_INTELLIGENCE_API_KEY || ''
  };
  
  if (!config.endpoint || !config.apiKey) {
    throw new Error('Azure Document Intelligence configuration is missing. Please set AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT and AZURE_DOCUMENT_INTELLIGENCE_API_KEY environment variables.');
  }
  
  return new AzureDocumentIntelligenceService(config);
}

// Export for testing and demonstration
export const test1099IntExtraction = async (ocrText: string) => {
  const service = new AzureDocumentIntelligenceService({
    endpoint: 'https://test.cognitiveservices.azure.com/',
    apiKey: 'test-key'
  });
  
  // Test the extraction logic without actual Azure calls
  return (service as any).extract1099IntFieldsAggressively(ocrText);
};
