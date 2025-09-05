
const { PrismaClient } = require('@prisma/client')

async function debugTaxData() {
  const prisma = new PrismaClient()
  
  try {
    console.log("=== DEBUG: Tax Return Data Analysis ===\n")
    
    // Get all users
    const users = await prisma.user.findMany({
      include: {
        taxReturns: {
          include: {
            incomeEntries: true,
            deductionEntries: true
          }
        }
      }
    })
    
    console.log(`Found ${users.length} users in database\n`)
    
    for (const user of users) {
      console.log(`User: ${user.email}`)
      console.log(`Tax Returns: ${user.taxReturns.length}`)
      
      for (const taxReturn of user.taxReturns) {
        console.log(`\n--- Tax Return ID: ${taxReturn.id} ---`)
        console.log(`Total Income: ${taxReturn.totalIncome}`)
        console.log(`Adjusted Gross Income: ${taxReturn.adjustedGrossIncome}`)
        console.log(`Filing Status: ${taxReturn.filingStatus}`)
        console.log(`Standard Deduction: ${taxReturn.standardDeduction}`)
        console.log(`Itemized Deduction: ${taxReturn.itemizedDeduction}`)
        console.log(`Income Entries: ${taxReturn.incomeEntries.length}`)
        console.log(`Deduction Entries: ${taxReturn.deductionEntries.length}`)
        
        // Calculate what the what-if scenarios should see
        const adjustedGrossIncome = parseFloat(taxReturn.adjustedGrossIncome || taxReturn.totalIncome || 0)
        const totalItemizedDeductions = taxReturn.deductionEntries.reduce(
          (sum, entry) => sum + parseFloat(entry.amount || 0), 0
        )
        
        console.log(`\n--- Calculated Values for What-If Scenarios ---`)
        console.log(`AGI for calculations: ${adjustedGrossIncome}`)
        console.log(`Total itemized deductions: ${totalItemizedDeductions}`)
        console.log(`Filing status: ${taxReturn.filingStatus}`)
        console.log(`Dependents: ${JSON.stringify(taxReturn.dependents)}`)
        
        // Test tax calculation with this data
        if (adjustedGrossIncome > 0) {
          console.log(`✅ This tax return has valid income data`)
        } else {
          console.log(`❌ This tax return has ZERO income - this is the problem!`)
        }
      }
      console.log("\n" + "=".repeat(50) + "\n")
    }
    
  } catch (error) {
    console.error('Error debugging tax data:', error)
  } finally {
    await prisma.$disconnect()
  }
}

debugTaxData()
