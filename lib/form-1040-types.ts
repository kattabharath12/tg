

// Form 1040 field definitions and types

export interface Form1040Data {
  // Header Information
  taxYear: number;
  firstName: string;
  lastName: string;
  ssn: string;
  spouseFirstName?: string;
  spouseLastName?: string;
  spouseSSN?: string;
  address: string;
  city: string;
  state: string;
  zipCode: string;
  filingStatus: FilingStatus;
  
  // Personal Info from W2 (new field for tracking source)
  personalInfo?: {
    firstName: string;
    lastName: string;
    ssn: string;
    address: string;
    city: string;
    state: string;
    zipCode: string;
    sourceDocument: string;
    sourceDocumentId: string;
  };

  // ADD THIS NEW PROPERTY HERE:
  employerInfo?: {
    name: string;
    address: string;
    ein: string;
  };
  
  // Dependents
  dependents: Dependent[];
  
  // Income Section (Lines 1-9)
  line1: number;  // Total amount from Form(s) W-2, box 1
  line2a: number; // Tax-exempt interest
  line2b: number; // Taxable interest
  line3a: number; // Qualified dividends
  line3b: number; // Ordinary dividends
  line4a: number; // IRA distributions
  line4b: number; // Taxable amount
  line5a: number; // Pensions and annuities
  line5b: number; // Taxable amount
  line6a: number; // Social security benefits
  line6b: number; // Taxable amount
  line7: number;  // Capital gain or (loss)
  line8: number;  // Additional income from Schedule 1, line 10
  line9: number;  // Add lines 1, 2b, 3b, 4b, 5b, 6b, 7, and 8. This is your total income
  
  // Adjusted Gross Income Section (Lines 10-11)
  line10: number; // Adjustments to income from Schedule 1, line 26
  line11: number; // Subtract line 10 from line 9. This is your adjusted gross income
  
  // Standard Deduction or Itemized Deductions (Line 12)
  line12: number; // Standard deduction or itemized deductions from Schedule A
  
  // Taxable Income (Line 15)
  line13: number; // Qualified business income deduction from Form 8995 or 8995-A
  line14: number; // Add lines 12 and 13
  line15: number; // Subtract line 14 from line 11. If zero or less, enter -0-. This is your taxable income
  
  // Tax and Credits Section (Lines 16-20)
  line16: number; // Tax (see instructions. Check if any from Form(s): 8814, 4972, other)
  line17: number; // Amount from Schedule 2, line 3
  line18: number; // Add lines 16 and 17
  line19: number; // Child tax credit and credit for other dependents from Schedule 8812
  line20: number; // Amount from Schedule 3, line 8
  
  // Payments Section (Lines 21-24)
  line21: number; // Add lines 19 and 20
  line22: number; // Subtract line 21 from line 18. If zero or less, enter -0-
  line23: number; // Other taxes from Schedule 2, line 21
  line24: number; // Add lines 22 and 23. This is your total tax
  
  // Federal Income Tax Withheld (Lines 25-31)
  line25a: number; // Federal income tax withheld from Form(s) W-2 and 1099
  line25b: number; // 2023 estimated tax payments and amount applied from 2022 return
  line25c: number; // Earned income credit (EIC)
  line25d: number; // Additional child tax credit from Schedule 8812
  
  // Refund or Amount Owed Section (Lines 32-37)
  line32: number; // Add lines 25a through 31 (total payments)
  line33: number; // If line 32 is more than line 24, subtract line 24 from line 32. This is the amount you overpaid
  line34: number; // Amount of line 33 you want refunded to you
  line35a: number; // Routing number
  line35b: string; // Account type: Checking or Savings
  line35c: number; // Account number
  line36: number; // Amount of line 33 you want applied to your 2024 estimated tax
  line37: number; // Subtract line 33 from line 24. This is the amount you owe
  
  // Third Party Designee and Sign Here sections
  thirdPartyDesignee: boolean;
  designeeName?: string;
  designeePhone?: string;
  designeePin?: string;
  
  // Paid Preparer Use Only (if applicable)
  preparerName?: string;
  preparerSSN?: string;
  preparerFirm?: string;
  preparerAddress?: string;
  preparerPhone?: string;
}

export interface Dependent {
  firstName: string;
  lastName: string;
  ssn: string;
  relationship: string;
  birthDate: Date;
  qualifiesForCTC: boolean; // Child Tax Credit
  qualifiesForEITC: boolean; // Earned Income Tax Credit
}

export enum FilingStatus {
  SINGLE = 'SINGLE',
  MARRIED_FILING_JOINTLY = 'MARRIED_FILING_JOINTLY',
  MARRIED_FILING_SEPARATELY = 'MARRIED_FILING_SEPARATELY',
  HEAD_OF_HOUSEHOLD = 'HEAD_OF_HOUSEHOLD',
  QUALIFYING_SURVIVING_SPOUSE = 'QUALIFYING_SURVIVING_SPOUSE'
}

// W2 to 1040 field mapping
export interface W2ToForm1040Mapping {
  // Direct mappings from W2 to 1040 lines
  wages_to_line1: (w2Data: any) => number;          // W2 Box 1 → 1040 Line 1
  federalTaxWithheld_to_line25a: (w2Data: any) => number; // W2 Box 2 → 1040 Line 25a
  socialSecurityWages_to_socialSecurity: (w2Data: any) => number; // W2 Box 3 (informational)
  medicareWages_to_medicare: (w2Data: any) => number;     // W2 Box 5 (informational)
  stateWages_to_stateReturn: (w2Data: any) => number;     // W2 Box 16 (for state return)
  stateTaxWithheld_to_stateReturn: (w2Data: any) => number; // W2 Box 17 (for state return)
  
  // Personal info mappings (new)
  employeeName_to_personalInfo: (w2Data: any) => { firstName: string; lastName: string };
  employeeSSN_to_personalInfo: (w2Data: any) => string;
  employeeAddress_to_personalInfo: (w2Data: any) => { address: string; city: string; state: string; zipCode: string };
}

// Standard deduction amounts for 2023 tax year
export const STANDARD_DEDUCTION_2023 = {
  [FilingStatus.SINGLE]: 13850,
  [FilingStatus.MARRIED_FILING_JOINTLY]: 27700,
  [FilingStatus.MARRIED_FILING_SEPARATELY]: 13850,
  [FilingStatus.HEAD_OF_HOUSEHOLD]: 20800,
  [FilingStatus.QUALIFYING_SURVIVING_SPOUSE]: 27700
};

// Tax brackets for 2023 (simplified)
export const TAX_BRACKETS_2023 = {
  [FilingStatus.SINGLE]: [
    { min: 0, max: 11000, rate: 0.10 },
    { min: 11000, max: 44725, rate: 0.12 },
    { min: 44725, max: 95375, rate: 0.22 },
    { min: 95375, max: 182050, rate: 0.24 },
    { min: 182050, max: 231250, rate: 0.32 },
    { min: 231250, max: 578125, rate: 0.35 },
    { min: 578125, max: Infinity, rate: 0.37 }
  ],
  [FilingStatus.MARRIED_FILING_JOINTLY]: [
    { min: 0, max: 22000, rate: 0.10 },
    { min: 22000, max: 89450, rate: 0.12 },
    { min: 89450, max: 190750, rate: 0.22 },
    { min: 190750, max: 364200, rate: 0.24 },
    { min: 364200, max: 462500, rate: 0.32 },
    { min: 462500, max: 693750, rate: 0.35 },
    { min: 693750, max: Infinity, rate: 0.37 }
  ],
  // ... other filing statuses can be added
};

// Child Tax Credit amounts
export const CHILD_TAX_CREDIT_2023 = {
  maxCredit: 2000,
  phaseoutThreshold: {
    [FilingStatus.SINGLE]: 200000,
    [FilingStatus.MARRIED_FILING_JOINTLY]: 400000,
    [FilingStatus.MARRIED_FILING_SEPARATELY]: 200000,
    [FilingStatus.HEAD_OF_HOUSEHOLD]: 200000,
    [FilingStatus.QUALIFYING_SURVIVING_SPOUSE]: 400000
  }
};

