
import { prisma } from "@/lib/db";

export interface DuplicateDetectionResult {
  isDuplicate: boolean;
  confidence: number;
  matchingDocuments: Array<{
    id: string;
    fileName: string;
    documentType: string;
    matchedFields: string[];
    similarity: number;
    createdAt: Date;
  }>;
  matchCriteria: {
    documentType: boolean;
    employerInfo: boolean;
    recipientInfo: boolean;
    amountSimilarity: boolean;
    nameSimilarity: boolean;
  };
}

export interface ExtractedDocumentData {
  documentType: string;
  extractedData: any;
  taxReturnId: string;
}

export class DuplicateDetectionService {
  private static readonly SIMILARITY_THRESHOLDS = {
    NAME_SIMILARITY: 0.8,
    AMOUNT_SIMILARITY: 0.95, // 95% similarity for amounts
    EMPLOYER_SIMILARITY: 0.85,
    OVERALL_DUPLICATE: 0.85
  };

  static async checkForDuplicates(
    newDocument: ExtractedDocumentData
  ): Promise<DuplicateDetectionResult> {
    console.log("🔍 [DUPLICATE] Starting duplicate detection for document type:", newDocument.documentType);

    try {
      // Find existing documents of the same type in the same tax return
      const allDocuments = await prisma.document.findMany({
        where: {
          taxReturnId: newDocument.taxReturnId,
          documentType: newDocument.documentType as any, // Cast to handle enum type
          processingStatus: 'COMPLETED'
        },
        select: {
          id: true,
          fileName: true,
          documentType: true,
          extractedData: true,
          createdAt: true
        }
      });

      // Filter out documents with null extractedData in TypeScript
      const existingDocuments = allDocuments.filter((doc: any) => doc.extractedData !== null);

      console.log("🔍 [DUPLICATE] Found", existingDocuments.length, "existing documents of same type");

      if (existingDocuments.length === 0) {
        return {
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

      const matches: DuplicateDetectionResult['matchingDocuments'] = [];
      
      for (const existingDoc of existingDocuments) {
        const similarity = this.calculateDocumentSimilarity(
          newDocument,
          existingDoc.extractedData as any
        );

        console.log("🔍 [DUPLICATE] Similarity with document", existingDoc.fileName, ":", similarity.overallScore);

        if (similarity.overallScore >= this.SIMILARITY_THRESHOLDS.OVERALL_DUPLICATE) {
          matches.push({
            id: existingDoc.id,
            fileName: existingDoc.fileName,
            documentType: existingDoc.documentType,
            matchedFields: similarity.matchedFields,
            similarity: similarity.overallScore,
            createdAt: existingDoc.createdAt
          });
        }
      }

      const bestMatch = matches.reduce((best, current) => 
        current.similarity > best.similarity ? current : best, 
        { similarity: 0 } as any
      );

      return {
        isDuplicate: matches.length > 0,
        confidence: bestMatch.similarity || 0,
        matchingDocuments: matches.sort((a, b) => b.similarity - a.similarity),
        matchCriteria: bestMatch.similarity > 0 ? this.getMatchCriteria(newDocument, bestMatch) : {
          documentType: false,
          employerInfo: false,
          recipientInfo: false,
          amountSimilarity: false,
          nameSimilarity: false
        }
      };

    } catch (error) {
      console.error("💥 [DUPLICATE] Error in duplicate detection:", error);
      // On error, assume no duplicates to allow processing to continue
      return {
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
  }

  private static calculateDocumentSimilarity(
    newDoc: ExtractedDocumentData,
    existingExtractedData: any
  ): { overallScore: number; matchedFields: string[] } {
    const newData = newDoc.extractedData;
    const existingData = existingExtractedData.extractedData || existingExtractedData;
    
    let totalScore = 0;
    let maxPossibleScore = 0;
    const matchedFields: string[] = [];

    // Document type match (mandatory for this comparison)
    if (newDoc.documentType === existingExtractedData.documentType) {
      totalScore += 10;
      matchedFields.push('documentType');
    }
    maxPossibleScore += 10;

    // Different comparison logic based on document type
    switch (newDoc.documentType) {
      case 'W2':
        return this.compareW2Documents(newData, existingData, totalScore, maxPossibleScore, matchedFields);
      
      case 'FORM_1099_INT':
      case 'FORM_1099_DIV':
      case 'FORM_1099_MISC':
      case 'FORM_1099_NEC':
      case 'FORM_1099_R':
      case 'FORM_1099_G':
        return this.compare1099Documents(newData, existingData, totalScore, maxPossibleScore, matchedFields);
      
      default:
        return this.compareGenericDocuments(newData, existingData, totalScore, maxPossibleScore, matchedFields);
    }
  }

  private static compareW2Documents(
    newData: any,
    existingData: any,
    totalScore: number,
    maxPossibleScore: number,
    matchedFields: string[]
  ): { overallScore: number; matchedFields: string[] } {
    // Employer information (high weight)
    const employerSimilarity = this.compareStrings(newData.employerName, existingData.employerName);
    const einMatch = this.normalizeEIN(newData.employerEIN) === this.normalizeEIN(existingData.employerEIN);
    
    if (employerSimilarity >= this.SIMILARITY_THRESHOLDS.EMPLOYER_SIMILARITY) {
      totalScore += 25;
      matchedFields.push('employerName');
    }
    if (einMatch && newData.employerEIN && existingData.employerEIN) {
      totalScore += 25;
      matchedFields.push('employerEIN');
    }
    maxPossibleScore += 50;

    // Employee information (high weight)
    const employeeSimilarity = this.compareStrings(newData.employeeName, existingData.employeeName);
    const ssnMatch = this.normalizeSSN(newData.employeeSSN) === this.normalizeSSN(existingData.employeeSSN);
    
    if (employeeSimilarity >= this.SIMILARITY_THRESHOLDS.NAME_SIMILARITY) {
      totalScore += 20;
      matchedFields.push('employeeName');
    }
    if (ssnMatch && newData.employeeSSN && existingData.employeeSSN) {
      totalScore += 30;
      matchedFields.push('employeeSSN');
    }
    maxPossibleScore += 50;

    // Amount comparisons (moderate weight)
    const wagesMatch = this.compareAmounts(newData.wages, existingData.wages);
    const federalTaxMatch = this.compareAmounts(newData.federalTaxWithheld, existingData.federalTaxWithheld);
    
    if (wagesMatch) {
      totalScore += 15;
      matchedFields.push('wages');
    }
    if (federalTaxMatch) {
      totalScore += 10;
      matchedFields.push('federalTaxWithheld');
    }
    maxPossibleScore += 25;

    const overallScore = maxPossibleScore > 0 ? totalScore / maxPossibleScore : 0;
    return { overallScore, matchedFields };
  }

  private static compare1099Documents(
    newData: any,
    existingData: any,
    totalScore: number,
    maxPossibleScore: number,
    matchedFields: string[]
  ): { overallScore: number; matchedFields: string[] } {
    // Payer information (high weight)
    const payerSimilarity = this.compareStrings(newData.payerName, existingData.payerName);
    const tinMatch = this.normalizeTIN(newData.payerTIN) === this.normalizeTIN(existingData.payerTIN);
    
    if (payerSimilarity >= this.SIMILARITY_THRESHOLDS.EMPLOYER_SIMILARITY) {
      totalScore += 25;
      matchedFields.push('payerName');
    }
    if (tinMatch && newData.payerTIN && existingData.payerTIN) {
      totalScore += 25;
      matchedFields.push('payerTIN');
    }
    maxPossibleScore += 50;

    // Recipient information (high weight)
    const recipientSimilarity = this.compareStrings(newData.recipientName, existingData.recipientName);
    const recipientTINMatch = this.normalizeTIN(newData.recipientTIN) === this.normalizeTIN(existingData.recipientTIN);
    
    if (recipientSimilarity >= this.SIMILARITY_THRESHOLDS.NAME_SIMILARITY) {
      totalScore += 20;
      matchedFields.push('recipientName');
    }
    if (recipientTINMatch && newData.recipientTIN && existingData.recipientTIN) {
      totalScore += 30;
      matchedFields.push('recipientTIN');
    }
    maxPossibleScore += 50;

    // Income amount comparisons (moderate weight)
    const incomeFields = this.getIncomeFieldsFor1099(newData);
    let incomeMatches = 0;
    let totalIncomeFields = 0;

    for (const field of incomeFields) {
      totalIncomeFields++;
      if (this.compareAmounts(newData[field], existingData[field])) {
        incomeMatches++;
        matchedFields.push(field);
      }
    }

    if (totalIncomeFields > 0) {
      const incomeScore = (incomeMatches / totalIncomeFields) * 20;
      totalScore += incomeScore;
    }
    maxPossibleScore += 20;

    // Federal tax withheld comparison
    if (this.compareAmounts(newData.federalTaxWithheld, existingData.federalTaxWithheld)) {
      totalScore += 10;
      matchedFields.push('federalTaxWithheld');
    }
    maxPossibleScore += 10;

    const overallScore = maxPossibleScore > 0 ? totalScore / maxPossibleScore : 0;
    return { overallScore, matchedFields };
  }

  private static compareGenericDocuments(
    newData: any,
    existingData: any,
    totalScore: number,
    maxPossibleScore: number,
    matchedFields: string[]
  ): { overallScore: number; matchedFields: string[] } {
    // Generic comparison for unknown document types
    const commonFields = ['payerName', 'recipientName', 'incomeAmount', 'taxWithheld'];
    
    for (const field of commonFields) {
      if (newData[field] && existingData[field]) {
        if (field.includes('Name')) {
          if (this.compareStrings(newData[field], existingData[field]) >= this.SIMILARITY_THRESHOLDS.NAME_SIMILARITY) {
            totalScore += 20;
            matchedFields.push(field);
          }
        } else if (field.includes('Amount') || field.includes('Withheld')) {
          if (this.compareAmounts(newData[field], existingData[field])) {
            totalScore += 15;
            matchedFields.push(field);
          }
        }
        maxPossibleScore += field.includes('Name') ? 20 : 15;
      }
    }

    const overallScore = maxPossibleScore > 0 ? totalScore / maxPossibleScore : 0;
    return { overallScore, matchedFields };
  }

  private static getIncomeFieldsFor1099(data: any): string[] {
    // Return relevant income fields based on what's available in the data
    const allIncomeFields = [
      'interestIncome', 'ordinaryDividends', 'qualifiedDividends', 
      'nonemployeeCompensation', 'rents', 'royalties', 'otherIncome',
      'totalCapitalGain', 'distributionAmount'
    ];
    
    return allIncomeFields.filter(field => data[field] !== undefined && data[field] !== '');
  }

  // Utility methods for comparison
  private static compareStrings(str1: string, str2: string): number {
    if (!str1 || !str2) return 0;
    
    const normalizedStr1 = this.normalizeString(str1);
    const normalizedStr2 = this.normalizeString(str2);
    
    if (normalizedStr1 === normalizedStr2) return 1;
    
    return this.calculateLevenshteinSimilarity(normalizedStr1, normalizedStr2);
  }

  private static compareAmounts(amount1: string | number, amount2: string | number): boolean {
    if (!amount1 || !amount2) return false;
    
    const num1 = this.normalizeAmount(amount1);
    const num2 = this.normalizeAmount(amount2);
    
    if (num1 === 0 && num2 === 0) return true;
    if (num1 === 0 || num2 === 0) return false;
    
    const similarity = 1 - Math.abs(num1 - num2) / Math.max(num1, num2);
    return similarity >= this.SIMILARITY_THRESHOLDS.AMOUNT_SIMILARITY;
  }

  private static normalizeString(str: string): string {
    return str.toLowerCase()
      .replace(/[^a-z0-9\s]/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  private static normalizeAmount(amount: string | number): number {
    if (typeof amount === 'number') return amount;
    if (!amount) return 0;
    
    // Remove currency symbols, commas, and other non-numeric characters except decimal point
    const cleanAmount = amount.toString().replace(/[^\d.-]/g, '');
    return parseFloat(cleanAmount) || 0;
  }

  private static normalizeSSN(ssn: string): string {
    if (!ssn) return '';
    return ssn.replace(/[^\d]/g, '');
  }

  private static normalizeEIN(ein: string): string {
    if (!ein) return '';
    return ein.replace(/[^\d]/g, '');
  }

  private static normalizeTIN(tin: string): string {
    if (!tin) return '';
    return tin.replace(/[^\d]/g, '');
  }

  private static calculateLevenshteinSimilarity(str1: string, str2: string): number {
    const matrix = Array(str2.length + 1).fill(null).map(() => Array(str1.length + 1).fill(null));
    
    for (let i = 0; i <= str1.length; i++) matrix[0][i] = i;
    for (let j = 0; j <= str2.length; j++) matrix[j][0] = j;
    
    for (let j = 1; j <= str2.length; j++) {
      for (let i = 1; i <= str1.length; i++) {
        const indicator = str1[i - 1] === str2[j - 1] ? 0 : 1;
        matrix[j][i] = Math.min(
          matrix[j][i - 1] + 1,
          matrix[j - 1][i] + 1,
          matrix[j - 1][i - 1] + indicator
        );
      }
    }
    
    const distance = matrix[str2.length][str1.length];
    const maxLength = Math.max(str1.length, str2.length);
    return maxLength === 0 ? 1 : 1 - distance / maxLength;
  }

  private static getMatchCriteria(newDoc: ExtractedDocumentData, match: any): DuplicateDetectionResult['matchCriteria'] {
    const matchedFields = match.matchedFields || [];
    
    return {
      documentType: matchedFields.includes('documentType'),
      employerInfo: matchedFields.some((field: string) => 
        field.includes('employer') || field.includes('payer') || field.includes('EIN') || field.includes('TIN')
      ),
      recipientInfo: matchedFields.some((field: string) => 
        field.includes('employee') || field.includes('recipient') || field.includes('SSN')
      ),
      amountSimilarity: matchedFields.some((field: string) => 
        field.includes('wages') || field.includes('Income') || field.includes('Amount') || field.includes('Withheld')
      ),
      nameSimilarity: matchedFields.some((field: string) => 
        field.includes('Name')
      )
    };
  }
}
