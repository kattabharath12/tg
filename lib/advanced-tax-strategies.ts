
/**
 * Advanced Tax Strategies Service
 * Implementation of sophisticated US Tax Code optimization strategies
 */

export interface TaxProfile {
  adjustedGrossIncome: number
  filingStatus: string
  age: number
  spouseAge?: number
  currentItemizedDeductions: number
  dependents: any[]
  employmentType: 'W2' | 'SELF_EMPLOYED' | 'BOTH' | 'RETIRED'
  hasBusinessIncome: boolean
  hasInvestmentIncome: boolean
  hasRetirementAccounts: boolean
  currentYear: number
}

export interface AdvancedStrategy {
  category: 'RETIREMENT' | 'INVESTMENTS' | 'DEDUCTIONS' | 'TIMING' | 'ESTATE' | 'BUSINESS' | 'CHARITABLE'
  strategy: string
  description: string
  implementationSteps: string[]
  potentialSavings: number
  difficulty: 'Easy' | 'Medium' | 'Hard' | 'Expert'
  timeline: string
  requirements: string[]
  considerations: string[]
  taxCodeReference?: string
}

export interface MultiYearPlan {
  currentYear: AdvancedStrategy[]
  nextYear: AdvancedStrategy[]
  longTerm: AdvancedStrategy[]
}

export class AdvancedTaxStrategist {
  
  /**
   * Retirement Account Optimization Strategies
   */
  static getRetirementOptimization(profile: TaxProfile): AdvancedStrategy[] {
    const strategies: AdvancedStrategy[] = []
    const { adjustedGrossIncome, age, filingStatus } = profile
    
    // 2025 contribution limits
    const limits = {
      k401: 23500,
      k401Catchup: 31000, // 50+
      iraContrib: 7000,
      iraCatchup: 8000, // 50+
      sep: Math.min(70000, adjustedGrossIncome * 0.25),
      simple: 16500,
      simpleCatchup: 20000 // 50+
    }

    // Traditional vs Roth analysis
    if (adjustedGrossIncome < 150000) {
      strategies.push({
        category: 'RETIREMENT',
        strategy: 'Roth IRA Conversion Ladder Strategy',
        description: `Convert traditional IRA funds to Roth IRA in lower-income years. Based on your current income of $${adjustedGrossIncome.toLocaleString()}, you're in an optimal tax bracket for conversions.`,
        implementationSteps: [
          'Calculate optimal conversion amount to stay within current tax bracket',
          'Execute partial Roth conversions annually',
          'Pay taxes on converted amount from non-retirement funds',
          'Build tax-free retirement income stream'
        ],
        potentialSavings: Math.round(adjustedGrossIncome * 0.15),
        difficulty: 'Medium',
        timeline: 'Multi-year strategy, start immediately',
        requirements: ['Existing traditional IRA/401(k) funds', 'Cash for tax payments'],
        considerations: ['5-year waiting period for penalty-free withdrawals', 'Impact on Medicare premiums'],
        taxCodeReference: 'IRC Section 408A'
      })
    }

    // Backdoor Roth strategy for high earners
    if (adjustedGrossIncome > 153000) {
      strategies.push({
        category: 'RETIREMENT',
        strategy: 'Backdoor Roth IRA Strategy',
        description: `Since your AGI exceeds Roth IRA eligibility limits, use backdoor Roth conversions to contribute $${age >= 50 ? limits.iraCatchup : limits.iraContrib} annually.`,
        implementationSteps: [
          'Contribute to non-deductible traditional IRA',
          'Immediately convert to Roth IRA',
          'Report conversion on Form 8606',
          'Ensure no existing traditional IRA balances to avoid pro-rata rule'
        ],
        potentialSavings: Math.round((age >= 50 ? limits.iraCatchup : limits.iraContrib) * 0.22),
        difficulty: 'Hard',
        timeline: 'Annual implementation by April 15',
        requirements: ['No existing traditional IRA balances preferred'],
        considerations: ['Pro-rata rule complications', 'State tax implications'],
        taxCodeReference: 'IRC Section 408A'
      })
    }

    // Mega Backdoor Roth for high earners with 401(k)
    if (adjustedGrossIncome > 200000) {
      strategies.push({
        category: 'RETIREMENT',
        strategy: 'Mega Backdoor Roth Strategy',
        description: `Contribute up to $70,000 ($77,500 if 50+) to 401(k) through after-tax contributions and in-service withdrawals to Roth IRA.`,
        implementationSteps: [
          'Verify employer plan allows after-tax contributions',
          'Maximize traditional 401(k) contributions first',
          'Make after-tax contributions up to annual limit',
          'Execute in-service Roth conversions'
        ],
        potentialSavings: 15400, // $70k * 22% tax bracket
        difficulty: 'Expert',
        timeline: 'Ongoing throughout year',
        requirements: ['Employer 401(k) plan support', 'High income capacity'],
        considerations: ['Plan administration complexity', 'Annual limit coordination'],
        taxCodeReference: 'IRC Section 402(g)'
      })
    }

    return strategies
  }

  /**
   * Tax-Loss Harvesting Strategies
   */
  static getTaxLossHarvesting(profile: TaxProfile): AdvancedStrategy[] {
    const strategies: AdvancedStrategy[] = []
    
    if (profile.hasInvestmentIncome) {
      strategies.push({
        category: 'INVESTMENTS',
        strategy: 'Strategic Tax-Loss Harvesting',
        description: 'Systematically realize investment losses to offset gains and reduce ordinary income by up to $3,000 annually.',
        implementationSteps: [
          'Review portfolio for unrealized losses',
          'Harvest losses while avoiding wash sale rules',
          'Reinvest in substantially different securities',
          'Carry forward excess losses to future years'
        ],
        potentialSavings: Math.min(3000 * (profile.adjustedGrossIncome > 100000 ? 0.24 : 0.22), 1000),
        difficulty: 'Medium',
        timeline: 'Ongoing, with year-end review',
        requirements: ['Taxable investment accounts with unrealized losses'],
        considerations: ['30-day wash sale rule', 'Impact on investment strategy'],
        taxCodeReference: 'IRC Section 1211'
      })

      strategies.push({
        category: 'INVESTMENTS',
        strategy: 'Direct Indexing for Tax Alpha',
        description: 'Replace index funds with direct stock ownership to enable security-level tax-loss harvesting.',
        implementationSteps: [
          'Transition from index funds to direct stock ownership',
          'Implement automated tax-loss harvesting',
          'Maintain market exposure while harvesting losses',
          'Optimize for after-tax returns'
        ],
        potentialSavings: Math.round(profile.adjustedGrossIncome * 0.01),
        difficulty: 'Expert',
        timeline: 'Transition over 6-12 months',
        requirements: ['Significant investment assets ($100K+)', 'Sophisticated platform access'],
        considerations: ['Increased complexity', 'Platform fees', 'Tracking burden'],
        taxCodeReference: 'IRC Section 1211'
      })
    }

    return strategies
  }

  /**
   * Charitable Giving Optimization
   */
  static getCharitableStrategies(profile: TaxProfile): AdvancedStrategy[] {
    const strategies: AdvancedStrategy[] = []
    
    // Bunching strategy
    const standardDeduction = this.getStandardDeduction(profile.filingStatus)
    if (profile.currentItemizedDeductions < standardDeduction * 1.5) {
      strategies.push({
        category: 'CHARITABLE',
        strategy: 'Charitable Deduction Bunching',
        description: `Bundle 2-3 years of charitable giving into one year to exceed the $${standardDeduction.toLocaleString()} standard deduction threshold.`,
        implementationSteps: [
          'Calculate 2-3 years of planned charitable giving',
          'Make accelerated donations in high-income year',
          'Use standard deduction in non-bunching years',
          'Consider donor-advised fund for timing flexibility'
        ],
        potentialSavings: Math.round((standardDeduction * 0.5) * (profile.adjustedGrossIncome > 100000 ? 0.24 : 0.22)),
        difficulty: 'Easy',
        timeline: 'Annual planning decision',
        requirements: ['Regular charitable giving habits'],
        considerations: ['Cash flow impact', 'Charitable organization preferences'],
        taxCodeReference: 'IRC Section 170'
      })
    }

    // Donor-Advised Fund strategy
    if (profile.adjustedGrossIncome > 75000) {
      strategies.push({
        category: 'CHARITABLE',
        strategy: 'Donor-Advised Fund Optimization',
        description: 'Contribute appreciated securities to DAF for immediate deduction while avoiding capital gains.',
        implementationSteps: [
          'Open donor-advised fund account',
          'Contribute appreciated securities instead of cash',
          'Claim full market value deduction',
          'Recommend grants to charities over time'
        ],
        potentialSavings: Math.round(profile.adjustedGrossIncome * 0.05 * 0.15), // Assume 5% giving, 15% capital gains saved
        difficulty: 'Medium',
        timeline: 'Setup within 30 days',
        requirements: ['Appreciated securities', 'Regular charitable giving'],
        considerations: ['DAF fees', 'Grant recommendation requirements'],
        taxCodeReference: 'IRC Section 170'
      })
    }

    // Qualified Charitable Distribution for seniors
    if (profile.age >= 70.5) {
      strategies.push({
        category: 'CHARITABLE',
        strategy: 'Qualified Charitable Distribution',
        description: 'Direct IRA distributions to charity to satisfy RMD requirements without increasing AGI.',
        implementationSteps: [
          'Coordinate with IRA custodian for direct transfers',
          'Ensure transfers go directly to qualified charities',
          'Document transfers for tax reporting',
          'Apply against required minimum distributions'
        ],
        potentialSavings: Math.round(Math.min(100000, profile.adjustedGrossIncome * 0.1) * 0.22),
        difficulty: 'Medium',
        timeline: 'Annual implementation after age 70½',
        requirements: ['Traditional IRA with RMD requirements'],
        considerations: ['Must be direct transfer', 'Annual $100K limit'],
        taxCodeReference: 'IRC Section 408(d)(8)'
      })
    }

    return strategies
  }

  /**
   * Business Expense Optimization
   */
  static getBusinessStrategies(profile: TaxProfile): AdvancedStrategy[] {
    const strategies: AdvancedStrategy[] = []
    
    if (profile.hasBusinessIncome || profile.employmentType === 'SELF_EMPLOYED') {
      // Section 179 deduction
      strategies.push({
        category: 'BUSINESS',
        strategy: 'Section 179 Equipment Deduction',
        description: 'Deduct up to $1,220,000 in business equipment purchases in 2025 instead of depreciating over time.',
        implementationSteps: [
          'Identify necessary business equipment purchases',
          'Ensure equipment qualifies for Section 179',
          'Make purchases and place in service by December 31',
          'File Form 4562 with tax return'
        ],
        potentialSavings: Math.round(Math.min(50000, profile.adjustedGrossIncome * 0.1) * 0.25),
        difficulty: 'Medium',
        timeline: 'Equipment must be placed in service by year-end',
        requirements: ['Active business', 'Qualifying equipment needs'],
        considerations: ['Business income limitations', 'Recapture on sale'],
        taxCodeReference: 'IRC Section 179'
      })

      // Home office deduction
      strategies.push({
        category: 'BUSINESS',
        strategy: 'Optimized Home Office Deduction',
        description: 'Maximize home office deductions using actual expense method for larger potential deductions.',
        implementationSteps: [
          'Measure and document exclusive business use area',
          'Calculate actual expenses vs. simplified method',
          'Maintain detailed records of home expenses',
          'Choose optimal calculation method annually'
        ],
        potentialSavings: Math.round(Math.min(15000, profile.adjustedGrossIncome * 0.05) * 0.25),
        difficulty: 'Easy',
        timeline: 'Ongoing record-keeping',
        requirements: ['Exclusive business use of home space'],
        considerations: ['Depreciation recapture', 'Record-keeping burden'],
        taxCodeReference: 'IRC Section 280A'
      })

      // Solo 401(k) for self-employed
      if (profile.employmentType === 'SELF_EMPLOYED') {
        strategies.push({
          category: 'RETIREMENT',
          strategy: 'Solo 401(k) Maximization',
          description: 'Contribute up to $70,000 ($77,500 if 50+) as both employee and employer to solo 401(k).',
          implementationSteps: [
            'Establish solo 401(k) plan',
            'Make employee deferrals up to $23,500',
            'Add employer contributions up to 25% of net self-employment income',
            'Consider loan features for liquidity'
          ],
          potentialSavings: Math.round(Math.min(70000, profile.adjustedGrossIncome * 0.25) * 0.22),
          difficulty: 'Medium',
          timeline: 'Establish by December 31 for current year',
          requirements: ['Self-employment income', 'No other employees'],
          considerations: ['Annual administration', 'Contribution calculation complexity'],
          taxCodeReference: 'IRC Section 401(k)'
        })
      }
    }

    return strategies
  }

  /**
   * Income Timing Strategies
   */
  static getIncomeTimingStrategies(profile: TaxProfile): AdvancedStrategy[] {
    const strategies: AdvancedStrategy[] = []
    
    // Roth conversion timing
    if (profile.hasRetirementAccounts && profile.adjustedGrossIncome < 100000) {
      strategies.push({
        category: 'TIMING',
        strategy: 'Low-Income Year Roth Conversions',
        description: 'Convert traditional retirement accounts to Roth during lower-income years to minimize tax impact.',
        implementationSteps: [
          'Identify years with lower projected income',
          'Calculate optimal conversion amounts within tax bracket',
          'Execute conversions in low-income years',
          'Consider multi-year conversion strategy'
        ],
        potentialSavings: Math.round(profile.adjustedGrossIncome * 0.1),
        difficulty: 'Hard',
        timeline: 'Multi-year planning strategy',
        requirements: ['Traditional retirement account balances'],
        considerations: ['Future tax rate assumptions', 'Medicare premium impacts'],
        taxCodeReference: 'IRC Section 408A'
      })
    }

    // Capital gains timing
    if (profile.hasInvestmentIncome) {
      strategies.push({
        category: 'TIMING',
        strategy: 'Capital Gains Rate Optimization',
        description: 'Time capital gains realizations to take advantage of 0% long-term capital gains rate for lower-income years.',
        implementationSteps: [
          'Project annual income and capital gains',
          'Identify opportunities for 0% capital gains years',
          'Realize gains when income is below thresholds',
          'Consider gain/loss pairing strategies'
        ],
        potentialSavings: Math.round(Math.min(25000, profile.adjustedGrossIncome * 0.2) * 0.15),
        difficulty: 'Medium',
        timeline: 'Ongoing portfolio management',
        requirements: ['Appreciated investment positions'],
        considerations: ['Market timing risks', 'Investment strategy alignment'],
        taxCodeReference: 'IRC Section 1(h)'
      })
    }

    return strategies
  }

  /**
   * Tax Bracket Management
   */
  static getTaxBracketManagement(profile: TaxProfile): AdvancedStrategy[] {
    const strategies: AdvancedStrategy[] = []
    
    // Bracket threshold analysis
    const brackets2025 = this.getTaxBrackets(profile.filingStatus)
    const currentBracket = this.getCurrentTaxBracket(profile.adjustedGrossIncome, profile.filingStatus)
    
    if (currentBracket) {
      const distanceToNext = currentBracket.max - profile.adjustedGrossIncome
      
      if (distanceToNext < 20000 && distanceToNext > 0) {
        strategies.push({
          category: 'TIMING',
          strategy: 'Tax Bracket Threshold Management',
          description: `You're $${distanceToNext.toLocaleString()} away from the next tax bracket. Consider strategies to stay within current ${(currentBracket.rate * 100)}% bracket.`,
          implementationSteps: [
            'Calculate exact distance to next tax bracket',
            'Maximize deductible contributions to stay in bracket',
            'Consider deferring income to next year',
            'Accelerate deductible expenses'
          ],
          potentialSavings: Math.round(distanceToNext * 0.05), // Savings from avoiding higher bracket
          difficulty: 'Medium',
          timeline: 'Year-end planning critical',
          requirements: ['Income timing flexibility'],
          considerations: ['Next year income projections', 'Cash flow needs'],
          taxCodeReference: 'IRC Section 1'
        })
      }
    }

    return strategies
  }

  /**
   * Estate Planning Tax Strategies
   */
  static getEstatePlanningStrategies(profile: TaxProfile): AdvancedStrategy[] {
    const strategies: AdvancedStrategy[] = []
    
    if (profile.adjustedGrossIncome > 200000) {
      // Annual exclusion gifting
      strategies.push({
        category: 'ESTATE',
        strategy: 'Annual Exclusion Gift Strategy',
        description: 'Gift up to $19,000 per recipient ($38,000 if married) annually to reduce taxable estate without using lifetime exemption.',
        implementationSteps: [
          'Identify potential gift recipients',
          'Calculate maximum annual exclusion amounts',
          'Execute gifts before December 31',
          'Maintain gift tax return filing requirements'
        ],
        potentialSavings: 7600, // $19k * 40% estate tax rate
        difficulty: 'Easy',
        timeline: 'Annual strategy by December 31',
        requirements: ['Excess assets for gifting'],
        considerations: ['Future need for gifted assets', 'Family dynamics'],
        taxCodeReference: 'IRC Section 2503(b)'
      })

      // Grantor trust strategies
      if (profile.adjustedGrossIncome > 500000) {
        strategies.push({
          category: 'ESTATE',
          strategy: 'Intentionally Defective Grantor Trust',
          description: 'Transfer appreciating assets to trust while paying income taxes personally to increase gift value to beneficiaries.',
          implementationSteps: [
            'Establish IDGT with qualified attorney',
            'Transfer appreciating assets to trust',
            'Structure to trigger grantor trust status',
            'Pay income taxes on trust earnings'
          ],
          potentialSavings: Math.round(profile.adjustedGrossIncome * 0.15),
          difficulty: 'Expert',
          timeline: '6-12 months to establish',
          requirements: ['Significant assets', 'Estate planning attorney'],
          considerations: ['Complexity', 'Ongoing tax payments', 'Irrevocability'],
          taxCodeReference: 'IRC Section 671-679'
        })
      }
    }

    return strategies
  }

  /**
   * Multi-Year Tax Planning
   */
  static getMultiYearPlan(profile: TaxProfile): MultiYearPlan {
    const currentYear = this.getAllStrategies(profile)
    
    // Project next year scenarios
    const nextYearProfile = { ...profile, currentYear: profile.currentYear + 1 }
    const nextYear = this.getAllStrategies(nextYearProfile)
    
    // Long-term strategies (5+ years)
    const longTerm: AdvancedStrategy[] = [
      {
        category: 'RETIREMENT',
        strategy: 'Retirement Distribution Planning',
        description: 'Develop long-term strategy for tax-efficient retirement withdrawals across multiple account types.',
        implementationSteps: [
          'Map out retirement account types and balances',
          'Model withdrawal sequences for tax efficiency',
          'Plan Roth conversion timeline',
          'Integrate Social Security timing decisions'
        ],
        potentialSavings: Math.round(profile.adjustedGrossIncome * 0.25),
        difficulty: 'Expert',
        timeline: '5-20 year strategy',
        requirements: ['Multiple retirement account types'],
        considerations: ['Market performance', 'Tax law changes', 'Healthcare costs'],
        taxCodeReference: 'IRC Section 401(a)(9)'
      }
    ]

    return {
      currentYear,
      nextYear,
      longTerm
    }
  }

  /**
   * Comprehensive Strategy Generation
   */
  static getAllStrategies(profile: TaxProfile): AdvancedStrategy[] {
    const allStrategies: AdvancedStrategy[] = [
      ...this.getRetirementOptimization(profile),
      ...this.getTaxLossHarvesting(profile),
      ...this.getCharitableStrategies(profile),
      ...this.getBusinessStrategies(profile),
      ...this.getIncomeTimingStrategies(profile),
      ...this.getTaxBracketManagement(profile),
      ...this.getEstatePlanningStrategies(profile)
    ]

    // Sort by potential savings and difficulty
    return allStrategies
      .sort((a, b) => {
        const difficultyWeight = { 'Easy': 4, 'Medium': 3, 'Hard': 2, 'Expert': 1 }
        const aScore = a.potentialSavings + (difficultyWeight[a.difficulty] * 1000)
        const bScore = b.potentialSavings + (difficultyWeight[b.difficulty] * 1000)
        return bScore - aScore
      })
      .slice(0, 8) // Return top 8 strategies
  }

  // Utility methods
  private static getStandardDeduction(filingStatus: string): number {
    const deductions2025 = {
      'SINGLE': 15000,
      'MARRIED_FILING_JOINTLY': 30000,
      'MARRIED_FILING_SEPARATELY': 15000,
      'HEAD_OF_HOUSEHOLD': 22500
    }
    return deductions2025[filingStatus as keyof typeof deductions2025] || 15000
  }

  private static getTaxBrackets(filingStatus: string) {
    // 2025 tax brackets (projected)
    const brackets = {
      'SINGLE': [
        { min: 0, max: 11925, rate: 0.10 },
        { min: 11926, max: 48475, rate: 0.12 },
        { min: 48476, max: 103350, rate: 0.22 },
        { min: 103351, max: 197300, rate: 0.24 },
        { min: 197301, max: 250525, rate: 0.32 },
        { min: 250526, max: 626350, rate: 0.35 },
        { min: 626351, max: Infinity, rate: 0.37 }
      ]
    }
    return brackets[filingStatus as keyof typeof brackets] || brackets.SINGLE
  }

  private static getCurrentTaxBracket(income: number, filingStatus: string) {
    const brackets = this.getTaxBrackets(filingStatus)
    return brackets.find(bracket => income >= bracket.min && income <= bracket.max)
  }
}
