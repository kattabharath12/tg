
import { NextRequest } from 'next/server'
import { getServerSession } from 'next-auth'
import { AdvancedTaxStrategist, TaxProfile } from '@/lib/advanced-tax-strategies'

export const dynamic = "force-dynamic"

export async function POST(request: NextRequest) {
  try {
    const session = await getServerSession()
    if (!session?.user?.email) {
      return new Response('Unauthorized', { status: 401 })
    }

    const body = await request.json()
    const {
      adjustedGrossIncome,
      filingStatus,
      currentItemizedDeductions,
      dependents,
      taxReturn,
      personalInfo = {},
      requestType = 'comprehensive' // 'comprehensive', 'specific_strategy', 'implementation'
    } = body

    // Create tax profile
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

    // Get expert strategies for context
    const expertStrategies = AdvancedTaxStrategist.getAllStrategies(profile)
    const topStrategies = expertStrategies.slice(0, 5)

    // Create comprehensive tax planning prompt
    const prompt = createTaxPlanningPrompt(profile, topStrategies, requestType)

    console.log('🔄 Streaming tax planning insights for profile:', {
      income: profile.adjustedGrossIncome,
      filingStatus: profile.filingStatus,
      age: profile.age,
      strategiesFound: expertStrategies.length
    })

    // Stream the AI response
    const response = await fetch('https://apps.abacus.ai/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${process.env.ABACUSAI_API_KEY}`
      },
      body: JSON.stringify({
        model: 'gpt-4.1-mini',
        messages: [
          {
            role: 'system',
            content: 'You are a highly experienced tax strategist and financial planner with deep expertise in US Tax Code optimization. Provide detailed, actionable advice with specific implementation steps, timelines, and considerations.'
          },
          {
            role: 'user',
            content: prompt
          }
        ],
        stream: true,
        max_tokens: 3500,
        temperature: 0.4,
      }),
    })

    if (!response.ok) {
      throw new Error(`AI API error: ${response.status}`)
    }

    // Create readable stream for response
    const encoder = new TextEncoder()
    const readable = new ReadableStream({
      async start(controller) {
        try {
          const reader = response.body?.getReader()
          if (!reader) {
            throw new Error('No response body')
          }

          while (true) {
            const { done, value } = await reader.read()
            if (done) break

            const chunk = new TextDecoder().decode(value)
            const lines = chunk.split('\n')

            for (const line of lines) {
              if (line.startsWith('data: ')) {
                const data = line.slice(6)
                if (data === '[DONE]') {
                  controller.enqueue(encoder.encode('data: [DONE]\n\n'))
                  controller.close()
                  return
                }
                
                try {
                  const parsed = JSON.parse(data)
                  const content = parsed.choices?.[0]?.delta?.content || ''
                  if (content) {
                    controller.enqueue(encoder.encode(`data: ${JSON.stringify({content})}\n\n`))
                  }
                } catch (e) {
                  // Skip invalid JSON chunks
                }
              }
            }
          }
        } catch (error) {
          console.error('Streaming error:', error)
          controller.error(error)
        }
      }
    })

    return new Response(readable, {
      headers: {
        'Content-Type': 'text/plain; charset=utf-8',
        'Cache-Control': 'no-cache',
      },
    })

  } catch (error) {
    console.error('Error in tax planning insights streaming:', error)
    return new Response('Failed to generate tax planning insights', { status: 500 })
  }
}

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

function createTaxPlanningPrompt(profile: TaxProfile, strategies: any[], requestType: string): string {
  const baseContext = `
TAXPAYER PROFILE ANALYSIS:
• Annual Income: $${profile.adjustedGrossIncome.toLocaleString()}
• Filing Status: ${profile.filingStatus.replace(/_/g, ' ')}
• Age: ${profile.age}${profile.spouseAge ? ` (Spouse: ${profile.spouseAge})` : ''}
• Employment Type: ${profile.employmentType.replace(/_/g, ' ')}
• Current Itemized Deductions: $${profile.currentItemizedDeductions.toLocaleString()}
• Number of Dependents: ${profile.dependents.length}
• Has Business Income: ${profile.hasBusinessIncome ? 'Yes' : 'No'}
• Has Investment Income: ${profile.hasInvestmentIncome ? 'Yes' : 'No'}
• Has Retirement Accounts: ${profile.hasRetirementAccounts ? 'Yes' : 'No'}

TOP OPTIMIZATION STRATEGIES IDENTIFIED:
${strategies.map((s, i) => `${i + 1}. ${s.strategy}
   Category: ${s.category}
   Potential Savings: $${s.potentialSavings.toLocaleString()}
   Difficulty: ${s.difficulty}
   Timeline: ${s.timeline}`).join('\n\n')}
`

  switch (requestType) {
    case 'comprehensive':
      return `${baseContext}

As an expert tax strategist, provide a comprehensive tax optimization analysis for this taxpayer. Cover:

1. **IMMEDIATE PRIORITIES (Next 60 Days)**
   - Most urgent strategies with highest impact
   - Year-end deadlines and considerations
   - Quick wins with minimal complexity

2. **STRATEGIC PLANNING (Next 12 Months)**
   - Long-term strategies requiring advance planning
   - Coordination between different optimization techniques
   - Risk management and diversification considerations

3. **MULTI-YEAR TAX PLANNING**
   - Income timing strategies across multiple years
   - Retirement account optimization over time
   - Estate planning integration

4. **IMPLEMENTATION ROADMAP**
   - Step-by-step action items with specific deadlines
   - Professional resources needed (CPA, financial advisor, attorney)
   - Common pitfalls to avoid

5. **ADVANCED CONSIDERATIONS**
   - Alternative Minimum Tax implications
   - State tax optimization opportunities
   - Business structure optimization (if applicable)

Provide specific dollar amounts, deadlines, and actionable next steps. Focus on strategies that deliver the highest return on investment of time and money.`

    case 'specific_strategy':
      return `${baseContext}

Provide detailed implementation guidance for the TOP 3 strategies identified above. For each strategy, include:

1. **DETAILED IMPLEMENTATION STEPS**
   - Exact procedures and documentation needed
   - Timeline and deadlines
   - Cost-benefit analysis

2. **COORDINATION OPPORTUNITIES**
   - How this strategy integrates with others
   - Sequencing for maximum benefit
   - Potential conflicts to avoid

3. **RISK MITIGATION**
   - What could go wrong and how to prevent it
   - IRS audit considerations
   - Documentation requirements

Focus on turning these strategies into actionable plans with specific next steps.`

    case 'implementation':
      return `${baseContext}

Create a detailed 90-day implementation plan for these tax optimization strategies. Include:

1. **WEEK-BY-WEEK ACTION PLAN**
   - Specific tasks for each week
   - Dependencies and prerequisites
   - Expected time commitments

2. **PROFESSIONAL TEAM ASSEMBLY**
   - Which specialists to engage (CPA, financial advisor, estate attorney)
   - Questions to ask potential advisors
   - Cost estimates for professional services

3. **DOCUMENT PREPARATION**
   - Forms and paperwork needed
   - Record-keeping systems to establish
   - Backup documentation requirements

4. **MONITORING AND ADJUSTMENT**
   - Key metrics to track
   - Quarterly review checkpoints
   - When to adjust strategies

Provide a practical, executable plan that transforms analysis into action.`

    default:
      return `${baseContext}

Provide expert tax planning insights and recommendations based on this taxpayer's profile and the identified optimization opportunities.`
  }
}
