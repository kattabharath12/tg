
import { NextRequest, NextResponse } from 'next/server'
import { getServerSession } from 'next-auth'
import { AdvancedTaxStrategist, TaxProfile } from '@/lib/advanced-tax-strategies'

export const dynamic = "force-dynamic"

// Helper functions
function determineEmploymentType(taxReturn: any): 'W2' | 'SELF_EMPLOYED' | 'BOTH' | 'RETIRED' {
  const hasW2 = taxReturn?.incomeEntries?.some((entry: any) => entry.type === 'W2') || false
  const hasBusiness = taxReturn?.incomeEntries?.some((entry: any) => 
    entry.type === 'BUSINESS' || entry.type === '1099_NEC'
  ) || false
  
  if (hasW2 && hasBusiness) return 'BOTH'
  if (hasBusiness) return 'SELF_EMPLOYED'
  if (hasW2) return 'W2'
  return 'RETIRED'
}

function hasBusinessIncome(taxReturn: any): boolean {
  return taxReturn?.incomeEntries?.some((entry: any) => 
    ['BUSINESS', '1099_NEC', '1099_MISC'].includes(entry.type)
  ) || false
}

function hasInvestmentIncome(taxReturn: any): boolean {
  return taxReturn?.incomeEntries?.some((entry: any) => 
    ['1099_INT', '1099_DIV', 'CAPITAL_GAINS'].includes(entry.type)
  ) || false
}

function hasRetirementAccounts(taxReturn: any): boolean {
  return taxReturn?.incomeEntries?.some((entry: any) => 
    ['1099_R', 'IRA_DISTRIBUTION'].includes(entry.type)
  ) || false
}

function getCurrentBracket(income: number, filingStatus: string): string {
  if (income < 50000) return '12%'
  if (income < 100000) return '22%'
  if (income < 200000) return '24%'
  if (income < 400000) return '32%'
  if (income < 600000) return '35%'
  return '37%'
}

function calculateOptimizationPotential(profile: TaxProfile): number {
  const strategies = AdvancedTaxStrategist.getAllStrategies(profile)
  return strategies.reduce((total, strategy) => total + strategy.potentialSavings, 0)
}

function identifyRiskFactors(profile: TaxProfile): string[] {
  const risks: string[] = []
  
  if (profile.adjustedGrossIncome > 400000) {
    risks.push('Alternative Minimum Tax exposure')
    risks.push('Net Investment Income Tax (3.8%)')
  }
  
  if (profile.age >= 59.5) {
    risks.push('Required Minimum Distribution planning needed')
  }
  
  if (profile.employmentType === 'SELF_EMPLOYED') {
    risks.push('Quarterly estimated tax payments required')
    risks.push('Self-employment tax optimization needed')
  }
  
  return risks
}

async function getAIEnhancedRecommendations(profile: TaxProfile, strategies: any[]): Promise<any[]> {
  try {
    const prompt = `You are an expert tax strategist providing personalized implementation guidance. Based on this taxpayer's profile and the identified strategies, provide specific, actionable implementation advice.

TAXPAYER PROFILE:
- Income: $${profile.adjustedGrossIncome.toLocaleString()}
- Filing Status: ${profile.filingStatus.replace(/_/g, ' ')}
- Age: ${profile.age}
- Employment: ${profile.employmentType.replace(/_/g, ' ')}
- Has Business Income: ${profile.hasBusinessIncome ? 'Yes' : 'No'}
- Has Investment Income: ${profile.hasInvestmentIncome ? 'Yes' : 'No'}
- Has Retirement Accounts: ${profile.hasRetirementAccounts ? 'Yes' : 'No'}

TOP STRATEGIES IDENTIFIED:
${strategies.slice(0, 3).map((s, i) => `${i + 1}. ${s.strategy} - Potential Savings: $${s.potentialSavings.toLocaleString()}`).join('\n')}

Provide 3-4 specific implementation insights focusing on:
1. Timing and sequencing of strategies
2. Potential pitfalls and how to avoid them
3. Integration opportunities between strategies
4. Specific next steps for the taxpayer

Respond in JSON format:
[
  {
    "insight": "Implementation Insight Title",
    "description": "Detailed guidance on how to implement or sequence strategies",
    "priority": "High|Medium|Low",
    "timeframe": "Specific timeframe for action"
  }
]

Respond with raw JSON only.`

    const response = await fetch('https://apps.abacus.ai/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${process.env.ABACUSAI_API_KEY}`
      },
      body: JSON.stringify({
        model: 'gpt-4.1-mini',
        messages: [{ role: 'user', content: prompt }],
        response_format: { type: "json_object" },
        max_tokens: 1500,
        temperature: 0.3,
      }),
    })

    if (response.ok) {
      const data = await response.json()
      const aiResponse = data.choices[0]?.message?.content
      if (aiResponse) {
        const parsed = JSON.parse(aiResponse)
        return Array.isArray(parsed) ? parsed : parsed.insights || []
      }
    }
  } catch (error) {
    console.error('AI enhancement error:', error)
  }

  // Fallback insights
  return [
    {
      insight: "Strategy Sequencing",
      description: "Prioritize retirement contributions before year-end, then consider Roth conversions in lower-income periods.",
      priority: "High",
      timeframe: "Before December 31, 2025"
    }
  ]
}

export async function POST(request: NextRequest) {
  try {
    const session = await getServerSession()
    if (!session?.user?.email) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }

    const body = await request.json()
    const {
      adjustedGrossIncome,
      filingStatus,
      currentItemizedDeductions,
      dependents,
      taxReturn,
      personalInfo = {}
    } = body

    // Create comprehensive tax profile
    const profile: TaxProfile = {
      adjustedGrossIncome: Math.max(0, adjustedGrossIncome || 0),
      filingStatus: filingStatus || 'SINGLE',
      age: personalInfo.age || 35,
      spouseAge: personalInfo.spouseAge,
      currentItemizedDeductions: Math.max(0, currentItemizedDeductions || 0),
      dependents: Array.isArray(dependents) ? dependents : [],
      employmentType: determineEmploymentType(taxReturn),
      hasBusinessIncome: hasBusinessIncome(taxReturn),
      hasInvestmentIncome: hasInvestmentIncome(taxReturn),
      hasRetirementAccounts: hasRetirementAccounts(taxReturn),
      currentYear: 2025
    }

    console.log('🧠 Advanced Tax Strategy Analysis for profile:', profile)

    // Get expert-level strategies using local calculation
    const expertStrategies = AdvancedTaxStrategist.getAllStrategies(profile)
    
    // Get multi-year planning
    const multiYearPlan = AdvancedTaxStrategist.getMultiYearPlan(profile)

    // Enhance with AI-powered personalized insights
    const aiEnhancedStrategies = await getAIEnhancedRecommendations(profile, expertStrategies)

    const response = {
      expertStrategies: expertStrategies.slice(0, 6),
      multiYearPlan,
      aiInsights: aiEnhancedStrategies,
      profileSummary: {
        taxBracket: getCurrentBracket(profile.adjustedGrossIncome, profile.filingStatus),
        optimizationPotential: calculateOptimizationPotential(profile),
        riskFactors: identifyRiskFactors(profile)
      }
    }

    return NextResponse.json(response)

  } catch (error) {
    console.error('Error in advanced tax strategies:', error)
    return NextResponse.json(
      { error: 'Failed to generate advanced tax strategies' }, 
      { status: 500 }
    )
  }
}
