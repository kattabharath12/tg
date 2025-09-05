
// Test the tax calculation functions directly
const { calculateTaxImpactScenarios, calculateDeductionComparison } = require('./lib/enhanced-tax-calculations')

console.log('🧮 Testing Tax Calculation Functions\n')

// Test with sample data that should produce realistic results
const testData = {
  adjustedGrossIncome: 75000,
  filingStatus: 'SINGLE',
  currentItemizedDeductions: 8000,
  dependents: []
}

console.log('Test Input:', testData)
console.log('\n--- Testing calculateTaxImpactScenarios ---')

try {
  const scenarios = calculateTaxImpactScenarios(
    testData.adjustedGrossIncome,
    testData.filingStatus,
    testData.currentItemizedDeductions,
    testData.dependents
  )
  
  console.log(`Generated ${scenarios.length} scenarios:`)
  scenarios.forEach((scenario, index) => {
    console.log(`${index + 1}. ${scenario.scenario}`)
    console.log(`   Description: ${scenario.description}`)
    console.log(`   Tax Liability: $${scenario.taxLiability.toLocaleString()}`)
    console.log(`   Savings: $${scenario.savings.toLocaleString()}`)
    console.log('')
  })
  
  if (scenarios.every(s => s.taxLiability === 0)) {
    console.log('❌ ERROR: All scenarios show $0 tax liability - there is still an issue!')
  } else {
    console.log('✅ SUCCESS: Tax calculations are working and showing real amounts!')
  }
  
} catch (error) {
  console.error('❌ ERROR in calculateTaxImpactScenarios:', error)
}

console.log('\n--- Testing calculateDeductionComparison ---')

try {
  const comparison = calculateDeductionComparison(
    testData.adjustedGrossIncome,
    testData.filingStatus,
    testData.currentItemizedDeductions,
    testData.dependents
  )
  
  console.log('Deduction Comparison Results:')
  console.log(`Standard Deduction: $${comparison.standardDeduction.toLocaleString()}`)
  console.log(`Itemized Deduction: $${comparison.itemizedDeduction.toLocaleString()}`)
  console.log(`Standard Tax Liability: $${comparison.standardTaxLiability.toLocaleString()}`)
  console.log(`Itemized Tax Liability: $${comparison.itemizedTaxLiability.toLocaleString()}`)
  console.log(`Recommended Method: ${comparison.recommendedMethod}`)
  console.log(`Tax Savings: $${comparison.taxSavings.toLocaleString()}`)
  
  if (comparison.standardTaxLiability === 0 && comparison.itemizedTaxLiability === 0) {
    console.log('❌ ERROR: Both tax liabilities are $0 - calculation issue!')
  } else {
    console.log('✅ SUCCESS: Deduction comparison is working!')
  }
  
} catch (error) {
  console.error('❌ ERROR in calculateDeductionComparison:', error)
}

console.log('\n--- Testing Edge Cases ---')

// Test with zero income
try {
  const zeroIncomeScenarios = calculateTaxImpactScenarios(0, 'SINGLE', 0, [])
  console.log(`Zero income test: Generated ${zeroIncomeScenarios.length} scenarios`)
  console.log(`Zero income base tax liability: $${zeroIncomeScenarios[0]?.taxLiability || 'N/A'}`)
} catch (error) {
  console.error('❌ ERROR with zero income:', error)
}

// Test with high income
try {
  const highIncomeScenarios = calculateTaxImpactScenarios(200000, 'MARRIED_FILING_JOINTLY', 25000, [])
  console.log(`High income test: Base tax liability: $${highIncomeScenarios[0]?.taxLiability?.toLocaleString() || 'N/A'}`)
} catch (error) {
  console.error('❌ ERROR with high income:', error)
}

console.log('\n🔍 Test completed!')
