
import { NextRequest, NextResponse } from 'next/server'
import { getServerSession } from 'next-auth'
import { prisma } from '@/lib/db'

export const dynamic = "force-dynamic"

export async function GET(request: NextRequest) {
  try {
    const session = await getServerSession()
    if (!session?.user?.email) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }

    // Get the user's tax return data
    const user = await prisma.user.findUnique({
      where: { email: session.user.email },
      include: {
        taxReturns: {
          include: {
            incomeEntries: true,
            deductionEntries: true
          }
        }
      }
    })

    if (!user) {
      return NextResponse.json({ error: 'User not found' }, { status: 404 })
    }

    const debugData = {
      userEmail: user.email,
      taxReturnsCount: user.taxReturns.length,
      taxReturns: user.taxReturns.map((taxReturn: any) => {
        const totalItemizedDeductions = taxReturn.deductionEntries.reduce(
          (sum: number, entry: any) => sum + parseFloat(entry.amount?.toString() || '0'), 0
        )
        
        const adjustedGrossIncome = parseFloat(taxReturn.adjustedGrossIncome?.toString() || taxReturn.totalIncome?.toString() || '0')
        
        return {
          id: taxReturn.id,
          totalIncome: taxReturn.totalIncome?.toString(),
          adjustedGrossIncome: taxReturn.adjustedGrossIncome?.toString(),
          calculatedAGI: adjustedGrossIncome,
          filingStatus: taxReturn.filingStatus,
          standardDeduction: taxReturn.standardDeduction?.toString(),
          itemizedDeduction: taxReturn.itemizedDeduction?.toString(),
          totalItemizedDeductions,
          incomeEntriesCount: taxReturn.incomeEntries.length,
          deductionEntriesCount: taxReturn.deductionEntries.length,
          dependents: (taxReturn as any).dependents || [],
          incomeEntries: taxReturn.incomeEntries.map((entry: any) => ({
            incomeType: entry.incomeType,
            amount: entry.amount?.toString(),
            description: entry.description
          })),
          deductionEntries: taxReturn.deductionEntries.map((entry: any) => ({
            deductionType: entry.deductionType,
            amount: entry.amount?.toString(),
            description: entry.description
          }))
        }
      })
    }

    return NextResponse.json(debugData, { status: 200 })

  } catch (error) {
    console.error('Error getting debug tax data:', error)
    return NextResponse.json(
      { error: 'Failed to get debug data' }, 
      { status: 500 }
    )
  }
}
