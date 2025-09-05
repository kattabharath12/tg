
import { Form1040Data } from './form-1040-types';

export class Form1099ToForm1040Mapper {
  /**
   * Maps 1099 extracted data to Form 1040 fields
   */
  static map1099ToForm1040(form1099Data: any, existingForm1040?: Partial<Form1040Data>): Partial<Form1040Data> {
    console.log('🔍 [1099 MAPPER] Starting 1099 to 1040 mapping...');
    if (process.env.NODE_ENV === 'development') {
      console.log('🔍 [1099 MAPPER] Input 1099Data structure:', JSON.stringify(form1099Data, null, 2));
      console.log('🔍 [1099 MAPPER] Existing form1040 data:', JSON.stringify(existingForm1040, null, 2));
    }

    const form1040Data: Partial<Form1040Data> = {
      ...existingForm1040,
    };

    // Handle nested data structures - the data might be nested under extractedData
    let actual1099Data = form1099Data;
    if (form1099Data.extractedData && typeof form1099Data.extractedData === 'object') {
      console.log('🔍 [1099 MAPPER] Found nested extractedData, using that instead');
      actual1099Data = form1099Data.extractedData;
    }

    console.log('🔍 [1099 MAPPER] Using actual1099Data:', JSON.stringify(actual1099Data, null, 2));

    // Personal Information Mapping - Enhanced to prioritize 1099 data
    console.log('🔍 [1099 MAPPER] Starting personal information mapping...');
    
    const recipientName = actual1099Data.recipientName || actual1099Data.Recipient?.Name || actual1099Data['Recipient.Name'];
    if (recipientName) {
      console.log('🔍 [1099 MAPPER] Mapping recipient name from 1099:', recipientName);
      const nameParts = recipientName.trim().split(/\s+/);
      // Only use 1099 data for personal info if not already set by W2
      if (!form1040Data.firstName || !form1040Data.personalInfo?.sourceDocument?.includes('W2')) {
        form1040Data.firstName = nameParts[0] || '';
        form1040Data.lastName = nameParts.slice(1).join(' ') || '';
        console.log('✅ [1099 MAPPER] Mapped name - First:', form1040Data.firstName, 'Last:', form1040Data.lastName);
      }
    }

    const recipientTIN = actual1099Data.recipientTIN || actual1099Data.Recipient?.TIN || actual1099Data['Recipient.TIN'];
    if (recipientTIN) {
      console.log('🔍 [1099 MAPPER] Mapping recipient TIN from 1099:', recipientTIN);
      // Only use 1099 data for SSN if not already set by W2
      if (!form1040Data.ssn || !form1040Data.personalInfo?.sourceDocument?.includes('W2')) {
        form1040Data.ssn = this.formatSSN(recipientTIN);
        console.log('✅ [1099 MAPPER] Mapped SSN:', form1040Data.ssn);
      }
    }

    // Enhanced address mapping
    const recipientAddress = actual1099Data.recipientAddress || actual1099Data.Recipient?.Address || actual1099Data['Recipient.Address'];
    
    if (recipientAddress && (!form1040Data.address || !form1040Data.personalInfo?.sourceDocument?.includes('W2'))) {
      console.log('🔍 [1099 MAPPER] Parsing recipient address from 1099:', recipientAddress);
      const addressParts = this.parseAddress(recipientAddress);
      form1040Data.address = addressParts.street;
      form1040Data.city = addressParts.city;
      form1040Data.state = addressParts.state;
      form1040Data.zipCode = addressParts.zipCode;
      console.log('✅ [1099 MAPPER] Mapped parsed address:', {
        street: form1040Data.address,
        city: form1040Data.city,
        state: form1040Data.state,
        zipCode: form1040Data.zipCode
      });
    }

    // Create or update personal info object
    if (!form1040Data.personalInfo || !form1040Data.personalInfo.sourceDocument?.includes('W2')) {
      const personalInfo = {
        firstName: form1040Data.firstName ?? '',
        lastName: form1040Data.lastName ?? '',
        ssn: form1040Data.ssn ?? '',
        address: form1040Data.address ?? '',
        city: form1040Data.city ?? '',
        state: form1040Data.state ?? '',
        zipCode: form1040Data.zipCode ?? '',
        sourceDocument: form1040Data.personalInfo?.sourceDocument ? 
          `${form1040Data.personalInfo.sourceDocument}, 1099` : '1099',
        sourceDocumentId: String(actual1099Data.documentId || 'unknown')
      };

      form1040Data.personalInfo = personalInfo;
      console.log('✅ [1099 MAPPER] Created/updated personalInfo object:', personalInfo);
    }

    // Income Mapping based on 1099 type
    console.log('🔍 [1099 MAPPER] Starting income mapping...');

    // 1099-INT: Interest Income → Line 2b (Taxable interest)
    const interestIncome = this.parseAmount(actual1099Data.interestIncome);
    if (interestIncome > 0) {
      form1040Data.line2b = (form1040Data.line2b || 0) + interestIncome;
      console.log('✅ [1099 MAPPER] Mapped interest income to Line 2b:', form1040Data.line2b);
    }

    // 1099-DIV: Dividend Income
    const ordinaryDividends = this.parseAmount(actual1099Data.ordinaryDividends);
    if (ordinaryDividends > 0) {
      form1040Data.line3b = (form1040Data.line3b || 0) + ordinaryDividends;
      console.log('✅ [1099 MAPPER] Mapped ordinary dividends to Line 3b:', form1040Data.line3b);
    }

    const qualifiedDividends = this.parseAmount(actual1099Data.qualifiedDividends);
    if (qualifiedDividends > 0) {
      form1040Data.line3a = (form1040Data.line3a || 0) + qualifiedDividends;
      console.log('✅ [1099 MAPPER] Mapped qualified dividends to Line 3a:', form1040Data.line3a);
    }

    // 1099-DIV: Capital Gain Distributions → Line 7 (Capital gains or losses)
    const totalCapitalGain = this.parseAmount(actual1099Data.totalCapitalGain);
    if (totalCapitalGain > 0) {
      form1040Data.line7 = (form1040Data.line7 || 0) + totalCapitalGain;
      console.log('✅ [1099 MAPPER] Mapped capital gain distributions to Line 7:', form1040Data.line7);
    }

    // 1099-MISC: Various income types
    const rents = this.parseAmount(actual1099Data.rents);
    if (rents > 0) {
      // Rents go to Schedule E, but for now we'll add to Line 8 (Additional income)
      form1040Data.line8 = (form1040Data.line8 || 0) + rents;
      console.log('✅ [1099 MAPPER] Mapped rents to Line 8 (Schedule E income):', form1040Data.line8);
    }

    const royalties = this.parseAmount(actual1099Data.royalties);
    if (royalties > 0) {
      // Royalties go to Schedule E, but for now we'll add to Line 8 (Additional income)
      form1040Data.line8 = (form1040Data.line8 || 0) + royalties;
      console.log('✅ [1099 MAPPER] Mapped royalties to Line 8 (Schedule E income):', form1040Data.line8);
    }

    const otherIncome = this.parseAmount(actual1099Data.otherIncome);
    if (otherIncome > 0) {
      // Other income goes to Line 8i (Other income)
      form1040Data.line8 = (form1040Data.line8 || 0) + otherIncome;
      console.log('✅ [1099 MAPPER] Mapped other income to Line 8:', form1040Data.line8);
    }

    // 1099-NEC: Nonemployee compensation → Schedule C (business income)
    const nonemployeeCompensation = this.parseAmount(actual1099Data.nonemployeeCompensation);
    if (nonemployeeCompensation > 0) {
      // This should go to Schedule C, but for now we'll add to Line 8 (Additional income)
      form1040Data.line8 = (form1040Data.line8 || 0) + nonemployeeCompensation;
      console.log('✅ [1099 MAPPER] Mapped nonemployee compensation to Line 8 (Schedule C income):', form1040Data.line8);
    }

    // Federal Tax Withheld → Line 25a
    const federalTaxWithheldValue = actual1099Data.federalTaxWithheld || actual1099Data.FederalIncomeTaxWithheld || 
                                   actual1099Data['FederalIncomeTaxWithheld'];
    
    console.log('🔍 [1099 MAPPER] Found federal tax withheld value:', federalTaxWithheldValue, 'type:', typeof federalTaxWithheldValue);
    const federalTaxWithheld = this.parseAmount(federalTaxWithheldValue);
    console.log('🔍 [1099 MAPPER] Parsed federal tax withheld amount:', federalTaxWithheld);
    
    if (federalTaxWithheld > 0) {
      form1040Data.line25a = (form1040Data.line25a || 0) + federalTaxWithheld;
      console.log('✅ [1099 MAPPER] Successfully mapped federal tax withheld to Line 25a:', form1040Data.line25a);
    } else {
      console.log('⚠️ [1099 MAPPER] No valid federal tax withheld found to map to Line 25a');
    }

    // Recalculate totals
    this.recalculateForm1040Totals(form1040Data);

    console.log('✅ [1099 MAPPER] Mapping completed successfully');
    if (process.env.NODE_ENV === 'development') {
      console.log('🔍 [1099 MAPPER] Final form1040Data:', JSON.stringify(form1040Data, null, 2));
    }
    return form1040Data;
  }

  /**
   * Creates a mapping summary showing what 1099 fields mapped to which 1040 lines
   */
  static createMappingSummary(form1099Data: any): Array<{
    form1099Field: string;
    form1099Value: any;
    form1040Line: string;
    form1040Value: any;
    description: string;
  }> {
    const mappings = [];

    if (form1099Data.recipientName) {
      mappings.push({
        form1099Field: 'Recipient Name',
        form1099Value: form1099Data.recipientName,
        form1040Line: 'Header',
        form1040Value: form1099Data.recipientName,
        description: 'Taxpayer name from 1099'
      });
    }

    if (form1099Data.recipientTIN) {
      mappings.push({
        form1099Field: 'Recipient TIN',
        form1099Value: form1099Data.recipientTIN,
        form1040Line: 'Header',
        form1040Value: this.formatSSN(form1099Data.recipientTIN),
        description: 'Taxpayer SSN from 1099'
      });
    }

    if (form1099Data.recipientAddress) {
      mappings.push({
        form1099Field: 'Recipient Address',
        form1099Value: form1099Data.recipientAddress,
        form1040Line: 'Header',
        form1040Value: form1099Data.recipientAddress,
        description: 'Taxpayer address from 1099'
      });
    }

    if (form1099Data.interestIncome) {
      mappings.push({
        form1099Field: 'Interest Income',
        form1099Value: form1099Data.interestIncome,
        form1040Line: 'Line 2b',
        form1040Value: this.parseAmount(form1099Data.interestIncome),
        description: 'Taxable interest income'
      });
    }

    if (form1099Data.ordinaryDividends) {
      mappings.push({
        form1099Field: 'Ordinary Dividends',
        form1099Value: form1099Data.ordinaryDividends,
        form1040Line: 'Line 3b',
        form1040Value: this.parseAmount(form1099Data.ordinaryDividends),
        description: 'Ordinary dividends'
      });
    }

    if (form1099Data.qualifiedDividends) {
      mappings.push({
        form1099Field: 'Qualified Dividends',
        form1099Value: form1099Data.qualifiedDividends,
        form1040Line: 'Line 3a',
        form1040Value: this.parseAmount(form1099Data.qualifiedDividends),
        description: 'Qualified dividends'
      });
    }

    if (form1099Data.totalCapitalGain) {
      mappings.push({
        form1099Field: 'Total Capital Gain Distributions',
        form1099Value: form1099Data.totalCapitalGain,
        form1040Line: 'Line 7',
        form1040Value: this.parseAmount(form1099Data.totalCapitalGain),
        description: 'Capital gain distributions'
      });
    }

    if (form1099Data.rents) {
      mappings.push({
        form1099Field: 'Rents',
        form1099Value: form1099Data.rents,
        form1040Line: 'Line 8 (Schedule E)',
        form1040Value: this.parseAmount(form1099Data.rents),
        description: 'Rental income (Schedule E)'
      });
    }

    if (form1099Data.royalties) {
      mappings.push({
        form1099Field: 'Royalties',
        form1099Value: form1099Data.royalties,
        form1040Line: 'Line 8 (Schedule E)',
        form1040Value: this.parseAmount(form1099Data.royalties),
        description: 'Royalty income (Schedule E)'
      });
    }

    if (form1099Data.otherIncome) {
      mappings.push({
        form1099Field: 'Other Income',
        form1099Value: form1099Data.otherIncome,
        form1040Line: 'Line 8',
        form1040Value: this.parseAmount(form1099Data.otherIncome),
        description: 'Other income'
      });
    }

    if (form1099Data.nonemployeeCompensation) {
      mappings.push({
        form1099Field: 'Nonemployee Compensation',
        form1099Value: form1099Data.nonemployeeCompensation,
        form1040Line: 'Line 8 (Schedule C)',
        form1040Value: this.parseAmount(form1099Data.nonemployeeCompensation),
        description: 'Business income (Schedule C)'
      });
    }

    if (form1099Data.federalTaxWithheld) {
      mappings.push({
        form1099Field: 'Federal Tax Withheld',
        form1099Value: form1099Data.federalTaxWithheld,
        form1040Line: 'Line 25a',
        form1040Value: this.parseAmount(form1099Data.federalTaxWithheld),
        description: 'Federal income tax withheld'
      });
    }

    return mappings;
  }

  private static parseAmount(value: any): number {
    console.log('🔍 [1099 PARSE AMOUNT] Input value:', value, 'type:', typeof value);
    
    if (value === null || value === undefined) {
      console.log('🔍 [1099 PARSE AMOUNT] Value is null/undefined, returning 0');
      return 0;
    }
    
    if (typeof value === 'number') {
      console.log('🔍 [1099 PARSE AMOUNT] Value is already a number:', value);
      return isNaN(value) ? 0 : value;
    }
    
    if (typeof value === 'string') {
      // Remove currency symbols, commas, and whitespace
      const cleaned = value.replace(/[$,\s]/g, '');
      console.log('🔍 [1099 PARSE AMOUNT] Cleaned string:', cleaned);
      const parsed = parseFloat(cleaned);
      const result = isNaN(parsed) ? 0 : parsed;
      console.log('🔍 [1099 PARSE AMOUNT] Parsed result:', result);
      return result;
    }
    
    // Handle objects that might have a value property
    if (typeof value === 'object' && value.value !== undefined) {
      console.log('🔍 [1099 PARSE AMOUNT] Object with value property:', value.value);
      return this.parseAmount(value.value);
    }
    
    // Handle objects that might have a content property (Azure DI format)
    if (typeof value === 'object' && value.content !== undefined) {
      console.log('🔍 [1099 PARSE AMOUNT] Object with content property:', value.content);
      return this.parseAmount(value.content);
    }
    
    console.log('🔍 [1099 PARSE AMOUNT] Unable to parse value, returning 0');
    return 0;
  }

  private static formatSSN(ssn: string): string {
    if (!ssn) return '';
    // Remove all non-digits
    const cleaned = ssn.replace(/\D/g, '');
    // Format as XXX-XX-XXXX
    if (cleaned.length === 9) {
      return `${cleaned.slice(0, 3)}-${cleaned.slice(3, 5)}-${cleaned.slice(5)}`;
    }
    return cleaned;
  }

  private static parseAddress(address: string): {
    street: string;
    city: string;
    state: string;
    zipCode: string;
  } {
    console.log('🔍 [1099 ADDRESS PARSER] Parsing address:', address);
    
    // Enhanced address parsing to handle various formats
    // Try comma-separated format first
    const commaParts = address.split(',').map(part => part.trim());
    
    if (commaParts.length >= 3) {
      const street = commaParts.slice(0, -2).join(', ');
      const city = commaParts[commaParts.length - 2];
      const stateZip = commaParts[commaParts.length - 1];
      const stateZipMatch = stateZip.match(/^([A-Z]{2})\s*(\d{5}(-\d{4})?)$/);
      
      if (stateZipMatch) {
        const result = {
          street,
          city,
          state: stateZipMatch[1],
          zipCode: stateZipMatch[2]
        };
        console.log('✅ [1099 ADDRESS PARSER] Parsed comma-separated address:', result);
        return result;
      }
    }
    
    // Handle format: "Street Address City, STATE ZIP" (2 parts)
    if (commaParts.length === 2) {
      const beforeComma = commaParts[0];
      const afterComma = commaParts[1];
      const stateZipMatch = afterComma.match(/^([A-Z]{2})\s+(\d{5}(-\d{4})?)$/);
      
      if (stateZipMatch) {
        const words = beforeComma.trim().split(/\s+/);
        if (words.length >= 2) {
          const streetEndIndex = words.length - 2;
          const street = words.slice(0, streetEndIndex + 1).join(' ');
          const city = words.slice(streetEndIndex + 1).join(' ');
          
          const result = {
            street,
            city,
            state: stateZipMatch[1],
            zipCode: stateZipMatch[2]
          };
          console.log('✅ [1099 ADDRESS PARSER] Parsed 2-part comma address:', result);
          return result;
        }
      }
    }
    
    // Try space-separated format: "Street Address City STATE ZIP"
    const spaceMatch = address.match(/^(.+?)\s+([A-Za-z\s]+?)\s+([A-Z]{2})\s+(\d{5}(-\d{4})?)$/);
    if (spaceMatch) {
      const result = {
        street: spaceMatch[1].trim(),
        city: spaceMatch[2].trim(),
        state: spaceMatch[3],
        zipCode: spaceMatch[4]
      };
      console.log('✅ [1099 ADDRESS PARSER] Parsed space-separated address:', result);
      return result;
    }
    
    // Try alternative pattern where city might have multiple words
    const stateZipPattern = /\s+([A-Z]{2})\s+(\d{5}(-\d{4})?)$/;
    const stateZipMatch = address.match(stateZipPattern);
    
    if (stateZipMatch) {
      const beforeStateZip = address.substring(0, address.length - stateZipMatch[0].length);
      const words = beforeStateZip.trim().split(/\s+/);
      
      if (words.length >= 2) {
        const streetEndIndex = words.length - 2;
        const street = words.slice(0, streetEndIndex + 1).join(' ');
        const city = words.slice(streetEndIndex + 1).join(' ');
        
        const result = {
          street,
          city,
          state: stateZipMatch[1],
          zipCode: stateZipMatch[2]
        };
        console.log('✅ [1099 ADDRESS PARSER] Parsed pattern-matched address:', result);
        return result;
      }
    }
    
    // Fallback: return the whole address as street
    const result = {
      street: address,
      city: '',
      state: '',
      zipCode: ''
    };
    console.log('⚠️ [1099 ADDRESS PARSER] Could not parse address, using as street only:', result);
    return result;
  }

  private static recalculateForm1040Totals(form1040Data: Partial<Form1040Data>): void {
    // Calculate total income (Line 9)
    form1040Data.line9 = this.calculateTotalIncome(form1040Data);

    // Calculate AGI (Line 11) - simplified (no adjustments for now)
    form1040Data.line11 = form1040Data.line9;

    // Apply standard deduction (Line 12) if not itemizing
    if (!form1040Data.line12 && form1040Data.filingStatus) {
      form1040Data.line12 = this.getStandardDeduction(form1040Data.filingStatus);
    }

    // Calculate taxable income (Line 15)
    form1040Data.line15 = Math.max(0, (form1040Data.line11 || 0) - (form1040Data.line12 || 0) - (form1040Data.line13 || 0));

    // Calculate tax liability (Line 16) - simplified
    form1040Data.line16 = this.calculateTaxLiability(form1040Data.line15 || 0, form1040Data.filingStatus);

    // Calculate total tax (Line 24) - simplified
    form1040Data.line24 = (form1040Data.line16 || 0) + (form1040Data.line17 || 0) + (form1040Data.line23 || 0);

    // Calculate total payments (Line 32)
    form1040Data.line32 = (form1040Data.line25a || 0) + (form1040Data.line25b || 0) + (form1040Data.line25c || 0) + (form1040Data.line25d || 0);

    // Calculate refund or amount owed
    const totalTax = form1040Data.line24 || 0;
    const totalPayments = form1040Data.line32 || 0;

    if (totalPayments > totalTax) {
      // Refund
      form1040Data.line33 = totalPayments - totalTax;
      form1040Data.line34 = form1040Data.line33; // Default to full refund
      form1040Data.line37 = 0;
    } else {
      // Amount owed
      form1040Data.line33 = 0;
      form1040Data.line34 = 0;
      form1040Data.line37 = totalTax - totalPayments;
    }
  }

  private static calculateTotalIncome(form1040Data: Partial<Form1040Data>): number {
    return (
      (form1040Data.line1 || 0) +
      (form1040Data.line2b || 0) +
      (form1040Data.line3b || 0) +
      (form1040Data.line4b || 0) +
      (form1040Data.line5b || 0) +
      (form1040Data.line6b || 0) +
      (form1040Data.line7 || 0) +
      (form1040Data.line8 || 0)
    );
  }

  private static getStandardDeduction(filingStatus: any): number {
    const STANDARD_DEDUCTION_2023: Record<string, number> = {
      'SINGLE': 13850,
      'MARRIED_FILING_JOINTLY': 27700,
      'MARRIED_FILING_SEPARATELY': 13850,
      'HEAD_OF_HOUSEHOLD': 20800,
      'QUALIFYING_SURVIVING_SPOUSE': 27700
    };
    
    return STANDARD_DEDUCTION_2023[filingStatus] || STANDARD_DEDUCTION_2023['SINGLE'];
  }

  private static calculateTaxLiability(taxableIncome: number, filingStatus: any): number {
    // Simplified tax calculation using 2023 tax brackets
    const brackets: Array<{ min: number; max: number; rate: number }> = 
      filingStatus === 'MARRIED_FILING_JOINTLY' ? [
        { min: 0, max: 22000, rate: 0.10 },
        { min: 22000, max: 89450, rate: 0.12 },
        { min: 89450, max: 190750, rate: 0.22 },
        { min: 190750, max: 364200, rate: 0.24 },
        { min: 364200, max: 462500, rate: 0.32 },
        { min: 462500, max: 693750, rate: 0.35 },
        { min: 693750, max: Infinity, rate: 0.37 }
      ] : [
        { min: 0, max: 11000, rate: 0.10 },
        { min: 11000, max: 44725, rate: 0.12 },
        { min: 44725, max: 95375, rate: 0.22 },
        { min: 95375, max: 182050, rate: 0.24 },
        { min: 182050, max: 231250, rate: 0.32 },
        { min: 231250, max: 578125, rate: 0.35 },
        { min: 578125, max: Infinity, rate: 0.37 }
      ];

    let tax = 0;
    let remainingIncome = taxableIncome;

    for (const bracket of brackets) {
      if (remainingIncome <= 0) break;
      
      const taxableAtThisBracket = Math.min(remainingIncome, bracket.max - bracket.min);
      tax += taxableAtThisBracket * bracket.rate;
      remainingIncome -= taxableAtThisBracket;
    }

    return Math.round(tax * 100) / 100; // Round to 2 decimal places
  }

  /**
   * Validates that 1099 data can be properly mapped to 1040
   */
  static validate1099DataForMapping(form1099Data: any): {
    isValid: boolean;
    errors: string[];
    warnings: string[];
  } {
    const errors: string[] = [];
    const warnings: string[] = [];

    // Check if there's any income data to map
    const hasIncomeData = form1099Data.interestIncome || 
                         form1099Data.ordinaryDividends || 
                         form1099Data.qualifiedDividends ||
                         form1099Data.totalCapitalGain ||
                         form1099Data.rents || 
                         form1099Data.royalties || 
                         form1099Data.otherIncome ||
                         form1099Data.nonemployeeCompensation;

    if (!hasIncomeData) {
      errors.push('No income data found in 1099 form');
    }

    // Warnings for missing optional but important fields
    if (!form1099Data.recipientName) {
      warnings.push('Recipient name not found - personal info may not be auto-populated');
    }

    if (!form1099Data.recipientTIN) {
      warnings.push('Recipient TIN not found - SSN field may not be auto-populated');
    }

    if (!form1099Data.recipientAddress) {
      warnings.push('Recipient address not found - address fields may not be auto-populated');
    }

    if (!form1099Data.payerName) {
      warnings.push('Payer name not found - may be needed for verification');
    }

    if (!form1099Data.payerTIN) {
      warnings.push('Payer TIN not found - may be needed for verification');
    }

    if (!form1099Data.federalTaxWithheld) {
      warnings.push('Federal tax withheld not found - no withholdings will be applied');
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }
}
