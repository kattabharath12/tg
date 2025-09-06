import { DocumentAnalysisClient, AzureKeyCredential } from "@azure/ai-form-recognizer";
import { DocumentType } from "@prisma/client";

export interface ExtractedFormData {
  documentType: DocumentType;
  personalInfo: {
    employerName?: string;
    employerEIN?: string;
    employeeSSN?: string;
    employeeName?: string;
    payerName?: string;
    payerTIN?: string;
    recipientSSN?: string;
    recipientName?: string;
    recipientAddress?: string;
  };
  formData: Record<string, any>;
  confidence: number;
  rawData?: any;
}

export class AzureDocumentIntelligenceService {
  private client: DocumentAnalysisClient;

  constructor() {
    const endpoint = process.env.AZURE_FORM_RECOGNIZER_ENDPOINT;
    const apiKey = process.env.AZURE_FORM_RECOGNIZER_API_KEY;

    if (!endpoint || !apiKey) {
      throw new Error("Azure Form Recognizer credentials not configured");
    }

    this.client = new DocumentAnalysisClient(endpoint, new AzureKeyCredential(apiKey));
  }

  async extractFormData(documentBuffer: Buffer, documentType: DocumentType): Promise<ExtractedFormData> {
    try {
      const poller = await this.client.beginAnalyzeDocument("prebuilt-document", documentBuffer);
      const result = await poller.pollUntilDone();

      if (!result.documents || result.documents.length === 0) {
        throw new Error("No documents found in the analysis result");
      }

      const document = result.documents[0];
      
      switch (documentType) {
        case DocumentType.FORM_W2:
          return this.extractW2Data(document, result);
        case DocumentType.FORM_1099_MISC:
          return this.extract1099MiscData(document, result);
        case DocumentType.FORM_1099_INT:
          return this.extract1099IntData(document, result);
        default:
          throw new Error(`Unsupported document type: ${documentType}`);
      }
    } catch (error) {
      console.error("Error extracting form data:", error);
      throw new Error(`Failed to extract form data: ${error.message}`);
    }
  }

  private extractW2Data(document: any, result: any): ExtractedFormData {
    const fields = document.fields || {};
    
    return {
      documentType: DocumentType.FORM_W2,
      personalInfo: {
        employerName: this.getFieldValue(fields["EmployerName"]) || this.extractByKeywords(result, ["employer", "company name"]),
        employerEIN: this.getFieldValue(fields["EmployerIdNumber"]) || this.extractByPattern(result, /\b\d{2}-\d{7}\b/),
        employeeSSN: this.getFieldValue(fields["EmployeeSSN"]) || this.extractByPattern(result, /\b\d{3}-\d{2}-\d{4}\b/),
        employeeName: this.getFieldValue(fields["EmployeeName"]) || this.extractByKeywords(result, ["employee", "name"])
      },
      formData: {
        wagesBox1: this.parseAmount(this.getFieldValue(fields["WagesTipsOtherComp"])),
        federalTaxWithheldBox2: this.parseAmount(this.getFieldValue(fields["FederalIncomeTaxWithheld"])),
        socialSecurityWagesBox3: this.parseAmount(this.getFieldValue(fields["SocialSecurityWages"])),
        socialSecurityTaxBox4: this.parseAmount(this.getFieldValue(fields["SocialSecurityTaxWithheld"])),
        medicareWagesBox5: this.parseAmount(this.getFieldValue(fields["MedicareWagesAndTips"])),
        medicareTaxBox6: this.parseAmount(this.getFieldValue(fields["MedicareTaxWithheld"])),
        socialSecurityTipsBox7: this.parseAmount(this.getFieldValue(fields["SocialSecurityTips"])),
        allocatedTipsBox8: this.parseAmount(this.getFieldValue(fields["AllocatedTips"])),
        dependentCareBenefitsBox10: this.parseAmount(this.getFieldValue(fields["DependentCareBenefits"])),
        nonqualifiedPlansBox11: this.parseAmount(this.getFieldValue(fields["NonqualifiedPlans"])),
        box12: this.extractBox12Codes(fields),
        statutoryEmployee: this.getFieldValue(fields["StatutoryEmployee"]) === "true",
        retirementPlan: this.getFieldValue(fields["RetirementPlan"]) === "true",
        thirdPartySickPay: this.getFieldValue(fields["ThirdPartySickPay"]) === "true",
        stateWagesBox16: this.parseAmount(this.getFieldValue(fields["StateWagesTipsEtc"])),
        stateTaxBox17: this.parseAmount(this.getFieldValue(fields["StateIncomeTax"])),
        localWagesBox18: this.parseAmount(this.getFieldValue(fields["LocalWagesTipsEtc"])),
        localTaxBox19: this.parseAmount(this.getFieldValue(fields["LocalIncomeTax"]))
      },
      confidence: this.calculateConfidence(fields),
      rawData: document
    };
  }

  private extract1099MiscData(document: any, result: any): ExtractedFormData {
    const fields = document.fields || {};
    
    return {
      documentType: DocumentType.FORM_1099_MISC,
      personalInfo: {
        payerName: this.getFieldValue(fields["PayerName"]) || this.extractByKeywords(result, ["payer", "company"]),
        payerTIN: this.getFieldValue(fields["PayerTIN"]) || this.extractByPattern(result, /\b\d{2}-\d{7}\b/),
        recipientSSN: this.getFieldValue(fields["RecipientTIN"]) || this.extractByPattern(result, /\b\d{3}-\d{2}-\d{4}\b/),
        recipientName: this.getFieldValue(fields["RecipientName"]) || this.extractByKeywords(result, ["recipient", "name"]),
        recipientAddress: this.getFieldValue(fields["RecipientAddress"])
      },
      formData: {
        rentsBox1: this.parseAmount(this.getFieldValue(fields["Rents"])),
        royaltiesBox2: this.parseAmount(this.getFieldValue(fields["Royalties"])),
        otherIncomeBox3: this.parseAmount(this.getFieldValue(fields["OtherIncome"])),
        federalTaxWithheldBox4: this.parseAmount(this.getFieldValue(fields["FederalIncomeTaxWithheld"])),
        fishingBoatProceedsBox5: this.parseAmount(this.getFieldValue(fields["FishingBoatProceeds"])),
        medicalHealthPaymentsBox6: this.parseAmount(this.getFieldValue(fields["MedicalHealthPayments"])),
        nonemployeeCompBox7: this.parseAmount(this.getFieldValue(fields["NonemployeeCompensation"])),
        substitutePaymentsBox8: this.parseAmount(this.getFieldValue(fields["SubstitutePayments"])),
        cropInsuranceProceedsBox9: this.parseAmount(this.getFieldValue(fields["CropInsuranceProceeds"])),
        grossProceedsBox10: this.parseAmount(this.getFieldValue(fields["GrossProceeds"])),
        excessGoldenParachuteBox11: this.parseAmount(this.getFieldValue(fields["ExcessGoldenParachute"])),
        section409ADeferralsBox12: this.parseAmount(this.getFieldValue(fields["Section409ADeferrals"])),
        eppaBox13: this.parseAmount(this.getFieldValue(fields["EPPA"])),
        nonqualifiedDeferredCompBox14: this.parseAmount(this.getFieldValue(fields["NonqualifiedDeferredComp"])),
        stateNumber: this.getFieldValue(fields["StateNumber"]),
        stateTaxWithheldBox16: this.parseAmount(this.getFieldValue(fields["StateTaxWithheld"])),
        stateIncomeBox17: this.parseAmount(this.getFieldValue(fields["StateIncome"]))
      },
      confidence: this.calculateConfidence(fields),
      rawData: document
    };
  }

  private extract1099IntData(document: any, result: any): ExtractedFormData {
    const fields = document.fields || {};
    
    // Enhanced extraction for 1099-INT specific fields
    const interestIncome = this.parseAmount(
      this.getFieldValue(fields["InterestIncome"]) || 
      this.extractByBoxNumber(result, "1") ||
      this.extractByKeywords(result, ["interest income", "box 1"])
    );

    const earlyWithdrawalPenalty = this.parseAmount(
      this.getFieldValue(fields["EarlyWithdrawalPenalty"]) || 
      this.extractByBoxNumber(result, "2") ||
      this.extractByKeywords(result, ["early withdrawal penalty", "box 2"])
    );

    const interestOnUSSavingsBonds = this.parseAmount(
      this.getFieldValue(fields["InterestOnUSSavingsBonds"]) || 
      this.extractByBoxNumber(result, "3") ||
      this.extractByKeywords(result, ["us savings bonds", "treasury obligations", "box 3"])
    );

    const federalTaxWithheld = this.parseAmount(
      this.getFieldValue(fields["FederalIncomeTaxWithheld"]) || 
      this.extractByBoxNumber(result, "4") ||
      this.extractByKeywords(result, ["federal income tax withheld", "box 4"])
    );

    const investmentExpenses = this.parseAmount(
      this.getFieldValue(fields["InvestmentExpenses"]) || 
      this.extractByBoxNumber(result, "5") ||
      this.extractByKeywords(result, ["investment expenses", "box 5"])
    );

    const foreignTaxPaid = this.parseAmount(
      this.getFieldValue(fields["ForeignTaxPaid"]) || 
      this.extractByBoxNumber(result, "6") ||
      this.extractByKeywords(result, ["foreign tax paid", "box 6"])
    );

    const foreignCountry = this.getFieldValue(fields["ForeignCountry"]) || 
      this.extractByKeywords(result, ["foreign country", "box 7"]);

    const taxExemptInterest = this.parseAmount(
      this.getFieldValue(fields["TaxExemptInterest"]) || 
      this.extractByBoxNumber(result, "8") ||
      this.extractByKeywords(result, ["tax-exempt interest", "box 8"])
    );

    const specifiedPrivateActivityBondInterest = this.parseAmount(
      this.getFieldValue(fields["SpecifiedPrivateActivityBondInterest"]) || 
      this.extractByBoxNumber(result, "9") ||
      this.extractByKeywords(result, ["private activity bond", "box 9"])
    );

    const marketDiscount = this.parseAmount(
      this.getFieldValue(fields["MarketDiscount"]) || 
      this.extractByBoxNumber(result, "10") ||
      this.extractByKeywords(result, ["market discount", "box 10"])
    );

    const bondPremium = this.parseAmount(
      this.getFieldValue(fields["BondPremium"]) || 
      this.extractByBoxNumber(result, "11") ||
      this.extractByKeywords(result, ["bond premium", "box 11"])
    );

    const bondPremiumOnTaxExemptBond = this.parseAmount(
      this.getFieldValue(fields["BondPremiumOnTaxExemptBond"]) || 
      this.extractByBoxNumber(result, "12") ||
      this.extractByKeywords(result, ["bond premium on tax-exempt bond", "box 12"])
    );

    const bondPremiumOnUSTreasury = this.parseAmount(
      this.getFieldValue(fields["BondPremiumOnUSTreasury"]) || 
      this.extractByBoxNumber(result, "13") ||
      this.extractByKeywords(result, ["bond premium on treasury", "box 13"])
    );

    const stateTaxWithheld = this.parseAmount(
      this.getFieldValue(fields["StateTaxWithheld"]) || 
      this.extractByBoxNumber(result, "14") ||
      this.extractByKeywords(result, ["state tax withheld", "box 14"])
    );

    const state = this.getFieldValue(fields["State"]) || 
      this.extractByKeywords(result, ["state", "box 15"]);

    const stateIdentificationNumber = this.getFieldValue(fields["StateIdentificationNumber"]) || 
      this.extractByKeywords(result, ["state identification", "box 16"]);

    return {
      documentType: DocumentType.FORM_1099_INT,
      personalInfo: {
        payerName: this.getFieldValue(fields["PayerName"]) || this.extractByKeywords(result, ["payer", "financial institution"]),
        payerTIN: this.getFieldValue(fields["PayerTIN"]) || this.extractByPattern(result, /\b\d{2}-\d{7}\b/),
        recipientSSN: this.getFieldValue(fields["RecipientTIN"]) || this.extractByPattern(result, /\b\d{3}-\d{2}-\d{4}\b/),
        recipientName: this.getFieldValue(fields["RecipientName"]) || this.extractByKeywords(result, ["recipient", "account holder"]),
        recipientAddress: this.getFieldValue(fields["RecipientAddress"])
      },
      formData: {
        interestIncomeBox1: interestIncome,
        earlyWithdrawalPenaltyBox2: earlyWithdrawalPenalty,
        interestOnUSSavingsBondsBox3: interestOnUSSavingsBonds,
        federalTaxWithheldBox4: federalTaxWithheld,
        investmentExpensesBox5: investmentExpenses,
        foreignTaxPaidBox6: foreignTaxPaid,
        foreignCountryBox7: foreignCountry,
        taxExemptInterestBox8: taxExemptInterest,
        specifiedPrivateActivityBondInterestBox9: specifiedPrivateActivityBondInterest,
        marketDiscountBox10: marketDiscount,
        bondPremiumBox11: bondPremium,
        bondPremiumOnTaxExemptBondBox12: bondPremiumOnTaxExemptBond,
        bondPremiumOnUSTreasuryBox13: bondPremiumOnUSTreasury,
        stateTaxWithheldBox14: stateTaxWithheld,
        stateBox15: state,
        stateIdentificationNumberBox16: stateIdentificationNumber
      },
      confidence: this.calculate1099IntConfidence({
        interestIncome,
        federalTaxWithheld,
        payerName: this.getFieldValue(fields["PayerName"]),
        recipientName: this.getFieldValue(fields["RecipientName"])
      }),
      rawData: document
    };
  }

  private getFieldValue(field: any): string | undefined {
    if (!field) return undefined;
    return field.content || field.value || field.valueString;
  }

  private parseAmount(value: string | undefined): number {
    if (!value) return 0;
    const cleanValue = value.replace(/[$,\s]/g, "");
    const parsed = parseFloat(cleanValue);
    return isNaN(parsed) ? 0 : parsed;
  }

  private extractByKeywords(result: any, keywords: string[]): string | undefined {
    if (!result.content) return undefined;
    
    const content = result.content.toLowerCase();
    for (const keyword of keywords) {
      const index = content.indexOf(keyword.toLowerCase());
      if (index !== -1) {
        // Extract text near the keyword
        const start = Math.max(0, index - 50);
        const end = Math.min(content.length, index + keyword.length + 50);
        const context = content.substring(start, end);
        
        // Look for numbers or text patterns near the keyword
        const numberMatch = context.match(/[\d,]+\.?\d*/);
        if (numberMatch) {
          return numberMatch[0];
        }
      }
    }
    return undefined;
  }

  private extractByPattern(result: any, pattern: RegExp): string | undefined {
    if (!result.content) return undefined;
    const match = result.content.match(pattern);
    return match ? match[0] : undefined;
  }

  private extractByBoxNumber(result: any, boxNumber: string): string | undefined {
    if (!result.content) return undefined;
    
    const content = result.content.toLowerCase();
    const boxPattern = new RegExp(`box\\s*${boxNumber}[:\\s]*([\\d,]+\\.?\\d*)`, 'i');
    const match = content.match(boxPattern);
    
    if (match) {
      return match[1];
    }
    
    // Alternative pattern for just the box number followed by amount
    const altPattern = new RegExp(`${boxNumber}[\\s\\$]*([\\d,]+\\.?\\d*)`, 'i');
    const altMatch = content.match(altPattern);
    
    return altMatch ? altMatch[1] : undefined;
  }

  private extractBox12Codes(fields: any): Array<{ code: string; amount: number }> {
    const box12Codes: Array<{ code: string; amount: number }> = [];
    
    // Look for Box 12 codes (a-dd)
    const codeLetters = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', 'aa', 'bb', 'cc', 'dd'];
    
    for (const code of codeLetters) {
      const fieldName = `Box12${code.toUpperCase()}`;
      if (fields[fieldName]) {
        const amount = this.parseAmount(this.getFieldValue(fields[fieldName]));
        if (amount > 0) {
          box12Codes.push({ code: code.toUpperCase(), amount });
        }
      }
    }
    
    return box12Codes;
  }

  private calculateConfidence(fields: any): number {
    const totalFields = Object.keys(fields).length;
    const filledFields = Object.values(fields).filter(field => 
      field && (field.content || field.value || field.valueString)
    ).length;
    
    return totalFields > 0 ? (filledFields / totalFields) * 100 : 0;
  }

  private calculate1099IntConfidence(data: any): number {
    let score = 0;
    let totalChecks = 0;

    // Check for required fields
    if (data.interestIncome > 0) {
      score += 30; // Interest income is the most important field
    }
    totalChecks += 30;

    if (data.payerName) {
      score += 25;
    }
    totalChecks += 25;

    if (data.recipientName) {
      score += 25;
    }
    totalChecks += 25;

    if (data.federalTaxWithheld >= 0) { // Can be 0
      score += 20;
    }
    totalChecks += 20;

    return totalChecks > 0 ? (score / totalChecks) * 100 : 0;
  }

  async validateDocument(documentBuffer: Buffer, expectedType: DocumentType): Promise<boolean> {
    try {
      const extractedData = await this.extractFormData(documentBuffer, expectedType);
      
      switch (expectedType) {
        case DocumentType.FORM_W2:
          return this.validateW2Data(extractedData);
        case DocumentType.FORM_1099_MISC:
          return this.validate1099MiscData(extractedData);
        case DocumentType.FORM_1099_INT:
          return this.validate1099IntData(extractedData);
        default:
          return false;
      }
    } catch (error) {
      console.error("Document validation failed:", error);
      return false;
    }
  }

  private validateW2Data(data: ExtractedFormData): boolean {
    const formData = data.formData;
    return !!(
      data.personalInfo.employerName &&
      data.personalInfo.employeeSSN &&
      formData.wagesBox1 >= 0 &&
      data.confidence > 50
    );
  }

  private validate1099MiscData(data: ExtractedFormData): boolean {
    const formData = data.formData;
    return !!(
      data.personalInfo.payerName &&
      data.personalInfo.recipientSSN &&
      (formData.rentsBox1 > 0 || formData.royaltiesBox2 > 0 || formData.otherIncomeBox3 > 0 || formData.nonemployeeCompBox7 > 0) &&
      data.confidence > 50
    );
  }

  private validate1099IntData(data: ExtractedFormData): boolean {
    const formData = data.formData;
    return !!(
      data.personalInfo.payerName &&
      data.personalInfo.recipientSSN &&
      formData.interestIncomeBox1 > 0 &&
      data.confidence > 50
    );
  }
}

export default AzureDocumentIntelligenceService;
