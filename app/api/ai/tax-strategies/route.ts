
import { NextRequest, NextResponse } from 'next/server'
import { getServerSession } from 'next-auth'

export const dynamic = "force-dynamic"

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
      taxReturn
    } = body

    // Prepare the context for AI analysis
    const taxContext = {
      income: adjustedGrossIncome,
      filingStatus: filingStatus,
      currentDeductions: currentItemizedDeductions,
      dependentCount: dependents?.length || 0,
      hasChildren: dependents?.some((d: any) => d.qualifiesForCTC) || false,
      currentYear: 2024,
      deductionEntries: taxReturn?.deductionEntries || []
    }

    const prompt = `You are a professional tax strategist. Analyze this taxpayer's situation and provide 4-6 specific, actionable tax optimization strategies.

TAXPAYER PROFILE:
- Adjusted Gross Income: $${adjustedGrossIncome.toLocaleString()}
- Filing Status: ${filingStatus.replace(/_/g, ' ')}
- Current Itemized Deductions: $${currentItemizedDeductions.toLocaleString()}
- Number of Dependents: ${dependents?.length || 0}
- Has Qualifying Children: ${dependents?.some((d: any) => d.qualifiesForCTC) ? 'Yes' : 'No'}

CURRENT DEDUCTION BREAKDOWN:
${taxReturn?.deductionEntries?.map((entry: any) => 
  `- ${entry.deductionType.replace(/_/g, ' ')}: $${parseFloat(entry.amount).toLocaleString()}`
).join('\n') || 'No itemized deductions entered yet'}

Provide specific, actionable tax strategies. For each strategy, include:
1. Strategy name (concise, under 50 characters)
2. Detailed description (2-3 sentences explaining what to do)
3. Estimated potential tax savings (realistic dollar amount)
4. Difficulty level (Easy/Medium/Hard)
5. Implementation timeline (e.g., "Before year-end", "Next tax year", etc.)

Focus on strategies like:
- Retirement contribution optimization
- Charitable giving strategies
- Medical expense timing
- Business expense optimization
- Tax-loss harvesting
- Education credits
- Energy-efficient home improvements
- HSA contributions
- Dependent care strategies

Respond in JSON format with this exact structure:
[
  {
    "strategy": "Strategy Name",
    "description": "Detailed description of what the taxpayer should do and why it helps.",
    "potentialSavings": 1200,
    "difficulty": "Easy",
    "timeline": "Before December 31, 2024"
  }
]

Respond with raw JSON only. Do not include code blocks, markdown, or any other formatting.`

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
            role: 'user',
            content: prompt
          }
        ],
        response_format: { type: "json_object" },
        max_tokens: 2000,
        temperature: 0.7,
      }),
    })

    if (!response.ok) {
      throw new Error(`AI API error: ${response.status}`)
    }

    const data = await response.json()
    const aiResponse = data.choices[0]?.message?.content

    if (!aiResponse) {
      throw new Error('No response from AI')
    }

    // Parse the JSON response
    let strategies
    try {
      const parsed = JSON.parse(aiResponse)
      // Handle both array format and object with array property
      strategies = Array.isArray(parsed) ? parsed : parsed.strategies || []
    } catch (parseError) {
      console.error('Failed to parse AI response:', parseError)
      // Fallback strategies if AI response is malformed
      strategies = [
        {
          strategy: "Maximize Retirement Contributions",
          description: "Consider increasing your 401(k) or IRA contributions to reduce taxable income. You may be able to contribute additional amounts before year-end.",
          potentialSavings: Math.round(adjustedGrossIncome * 0.02),
          difficulty: "Easy",
          timeline: "Before December 31, 2024"
        },
        {
          strategy: "Strategic Charitable Giving",
          description: "Bundle charitable donations into a single tax year or consider donating appreciated securities to maximize deduction benefits.",
          potentialSavings: Math.round(adjustedGrossIncome * 0.015),
          difficulty: "Medium",
          timeline: "Before December 31, 2024"
        }
      ]
    }

    // Validate and sanitize the response
    const validatedStrategies = strategies.slice(0, 6).map((strategy: any) => ({
      strategy: String(strategy.strategy || "Tax Strategy").substring(0, 100),
      description: String(strategy.description || "Consult a tax professional for personalized advice.").substring(0, 300),
      potentialSavings: Math.max(0, Math.min(50000, Number(strategy.potentialSavings) || 0)),
      difficulty: ['Easy', 'Medium', 'Hard'].includes(strategy.difficulty) ? strategy.difficulty : 'Medium',
      timeline: String(strategy.timeline || "Consult with tax professional").substring(0, 100)
    }))

    return NextResponse.json(validatedStrategies)

  } catch (error) {
    console.error('Error getting tax strategies:', error)
    return NextResponse.json(
      { error: 'Failed to get tax strategies' }, 
      { status: 500 }
    )
  }
}
