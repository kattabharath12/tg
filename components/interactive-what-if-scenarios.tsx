
"use client"

import { useState, useEffect } from "react"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card"
import { Badge } from "@/components/ui/badge"
import { Alert, AlertDescription } from "@/components/ui/alert"
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs"
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select"
import { 
  Calculator, 
  TrendingUp, 
  TrendingDown, 
  DollarSign, 
  Lightbulb, 
  Brain,
  Target,
  Plus,
  Minus,
  ArrowUpDown,
  Sparkles,
  Info
} from "lucide-react"
import { calculateDeductionComparison, calculateTaxImpactScenarios } from "@/lib/enhanced-tax-calculations"

interface InteractiveWhatIfScenariosProps {
  taxReturn: any
  adjustedGrossIncome: number
  currentItemizedDeductions: number
  filingStatus: string
  dependents: any[]
}

interface CustomScenario {
  id: string
  name: string
  additionalAmount: number
  deductionType: string
}

interface AIRecommendation {
  strategy: string
  description: string
  potentialSavings: number
  difficulty: 'Easy' | 'Medium' | 'Hard'
  timeline: string
}

interface AdvancedStrategy {
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

interface AdvancedTaxResponse {
  expertStrategies: AdvancedStrategy[]
  multiYearPlan: {
    currentYear: AdvancedStrategy[]
    nextYear: AdvancedStrategy[]
    longTerm: AdvancedStrategy[]
  }
  aiInsights: any[]
  profileSummary: {
    taxBracket: string
    optimizationPotential: number
    riskFactors: string[]
  }
}

export function InteractiveWhatIfScenarios({
  taxReturn,
  adjustedGrossIncome,
  currentItemizedDeductions,
  filingStatus,
  dependents
}: InteractiveWhatIfScenariosProps) {
  const [customScenarios, setCustomScenarios] = useState<CustomScenario[]>([])
  const [newScenario, setNewScenario] = useState({
    name: "",
    additionalAmount: "",
    deductionType: "CHARITABLE_CONTRIBUTIONS"
  })
  const [calculations, setCalculations] = useState<any[]>([])
  const [aiRecommendations, setAIRecommendations] = useState<AIRecommendation[]>([])
  const [loadingAI, setLoadingAI] = useState(false)
  const [quickAmount, setQuickAmount] = useState("")
  const [activeTab, setActiveTab] = useState("scenarios")
  const [debugInfo, setDebugInfo] = useState<any>(null)
  
  // Advanced strategy states
  const [advancedStrategies, setAdvancedStrategies] = useState<AdvancedStrategy[]>([])
  const [multiYearPlan, setMultiYearPlan] = useState<any>(null)
  const [profileSummary, setProfileSummary] = useState<any>(null)
  const [loadingAdvanced, setLoadingAdvanced] = useState(false)
  const [selectedStrategy, setSelectedStrategy] = useState<AdvancedStrategy | null>(null)
  
  // Streaming insights states
  const [streamingInsights, setStreamingInsights] = useState("")
  const [isStreaming, setIsStreaming] = useState(false)
  const [streamingType, setStreamingType] = useState<'comprehensive' | 'specific_strategy' | 'implementation'>('comprehensive')

  // Validate and sanitize input data
  const safeAdjustedGrossIncome = Math.max(0, parseFloat(String(adjustedGrossIncome || 0)))
  const safeCurrentItemizedDeductions = Math.max(0, parseFloat(String(currentItemizedDeductions || 0)))
  const safeFilingStatus = filingStatus || 'SINGLE'
  const safeDependents = Array.isArray(dependents) ? dependents : []

  // Debug logging
  useEffect(() => {
    const debug = {
      originalAGI: adjustedGrossIncome,
      safeAGI: safeAdjustedGrossIncome,
      originalItemized: currentItemizedDeductions,
      safeItemized: safeCurrentItemizedDeductions,
      filingStatus: safeFilingStatus,
      dependentsCount: safeDependents.length,
      taxReturn: {
        id: taxReturn?.id,
        totalIncome: taxReturn?.totalIncome,
        adjustedGrossIncome: taxReturn?.adjustedGrossIncome
      }
    }
    setDebugInfo(debug)
    console.log('🔍 What-If Scenarios Debug Info:', debug)
  }, [adjustedGrossIncome, currentItemizedDeductions, filingStatus, dependents, taxReturn])

  const deductionTypes = [
    { value: "CHARITABLE_CONTRIBUTIONS", label: "Charitable Donations" },
    { value: "MORTGAGE_INTEREST", label: "Mortgage Interest" },
    { value: "STATE_LOCAL_TAXES", label: "State & Local Taxes" },
    { value: "MEDICAL_EXPENSES", label: "Medical Expenses" },
    { value: "BUSINESS_EXPENSES", label: "Business Expenses" },
    { value: "STUDENT_LOAN_INTEREST", label: "Student Loan Interest" },
    { value: "IRA_CONTRIBUTIONS", label: "Retirement Contributions" },
    { value: "OTHER_DEDUCTIONS", label: "Other Deductions" }
  ]

  // Calculate base scenario and custom scenarios
  useEffect(() => {
    // Use safe values and add minimum income if needed for realistic calculations
    const minIncomeForCalculation = Math.max(safeAdjustedGrossIncome, 50000) // Use at least $50k for demo purposes if no income
    
    console.log('🧮 Calculating scenarios with:', {
      income: safeAdjustedGrossIncome,
      minIncome: minIncomeForCalculation,
      itemizedDeductions: safeCurrentItemizedDeductions,
      filingStatus: safeFilingStatus
    })

    const baseScenarios = calculateTaxImpactScenarios(
      minIncomeForCalculation,
      safeFilingStatus,
      safeCurrentItemizedDeductions,
      safeDependents
    )

    // Add custom scenarios
    const customCalculations = customScenarios.map(scenario => {
      const newComparison = calculateDeductionComparison(
        minIncomeForCalculation,
        safeFilingStatus,
        safeCurrentItemizedDeductions + scenario.additionalAmount,
        safeDependents
      )
      
      const baseTaxLiability = baseScenarios[0]?.taxLiability || 0
      const newTaxLiability = newComparison.recommendedMethod === 'itemized'
        ? newComparison.itemizedTaxLiability
        : newComparison.standardTaxLiability

      return {
        scenario: scenario.name,
        description: `${deductionTypes.find(t => t.value === scenario.deductionType)?.label}: +$${scenario.additionalAmount.toLocaleString()}`,
        itemizedDeductions: safeCurrentItemizedDeductions + scenario.additionalAmount,
        taxLiability: newTaxLiability,
        savings: baseTaxLiability - newTaxLiability,
        custom: true,
        id: scenario.id
      }
    })

    console.log('📊 Calculated scenarios:', [...baseScenarios, ...customCalculations])
    setCalculations([...baseScenarios, ...customCalculations])
  }, [safeAdjustedGrossIncome, safeFilingStatus, safeCurrentItemizedDeductions, safeDependents, customScenarios])

  const handleAddCustomScenario = () => {
    if (!newScenario.name || !newScenario.additionalAmount) return

    const scenario: CustomScenario = {
      id: Date.now().toString(),
      name: newScenario.name,
      additionalAmount: parseFloat(newScenario.additionalAmount),
      deductionType: newScenario.deductionType
    }

    setCustomScenarios([...customScenarios, scenario])
    setNewScenario({
      name: "",
      additionalAmount: "",
      deductionType: "CHARITABLE_CONTRIBUTIONS"
    })
  }

  const handleRemoveCustomScenario = (id: string) => {
    setCustomScenarios(customScenarios.filter(s => s.id !== id))
  }

  const handleQuickCalculation = () => {
    if (!quickAmount) return

    const amount = parseFloat(quickAmount)
    const quickScenario: CustomScenario = {
      id: `quick-${Date.now()}`,
      name: `Quick Test: $${amount.toLocaleString()}`,
      additionalAmount: amount,
      deductionType: "OTHER_DEDUCTIONS"
    }

    setCustomScenarios([...customScenarios, quickScenario])
    setQuickAmount("")
  }

  const getAIRecommendations = async () => {
    setLoadingAI(true)
    try {
      // Use minimum income for realistic AI recommendations
      const minIncomeForAI = Math.max(safeAdjustedGrossIncome, 50000)
      
      console.log('🤖 Getting AI recommendations with:', {
        income: minIncomeForAI,
        itemizedDeductions: safeCurrentItemizedDeductions,
        filingStatus: safeFilingStatus
      })

      const response = await fetch('/api/ai/tax-strategies', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          adjustedGrossIncome: minIncomeForAI,
          filingStatus: safeFilingStatus,
          currentItemizedDeductions: safeCurrentItemizedDeductions,
          dependents: safeDependents,
          taxReturn: {
            totalIncome: taxReturn?.totalIncome || minIncomeForAI,
            deductionEntries: taxReturn?.deductionEntries || []
          }
        }),
      })

      if (response.ok) {
        const recommendations = await response.json()
        console.log('🎯 Received AI recommendations:', recommendations)
        setAIRecommendations(recommendations)
      } else {
        console.error('AI API response not ok:', response.status, response.statusText)
      }
    } catch (error) {
      console.error('Error getting AI recommendations:', error)
    } finally {
      setLoadingAI(false)
    }
  }

  // Advanced tax strategies function
  const getAdvancedStrategies = async () => {
    setLoadingAdvanced(true)
    try {
      const minIncomeForAdvanced = Math.max(safeAdjustedGrossIncome, 50000)
      
      console.log('🧠 Getting advanced tax strategies with:', {
        income: minIncomeForAdvanced,
        itemizedDeductions: safeCurrentItemizedDeductions,
        filingStatus: safeFilingStatus
      })

      const response = await fetch('/api/ai/advanced-tax-strategies', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          adjustedGrossIncome: minIncomeForAdvanced,
          filingStatus: safeFilingStatus,
          currentItemizedDeductions: safeCurrentItemizedDeductions,
          dependents: safeDependents,
          taxReturn: {
            totalIncome: taxReturn?.totalIncome || minIncomeForAdvanced,
            deductionEntries: taxReturn?.deductionEntries || [],
            incomeEntries: taxReturn?.incomeEntries || []
          },
          personalInfo: {
            age: 35, // Default age, could be made configurable
            spouseAge: undefined
          }
        }),
      })

      if (response.ok) {
        const data: AdvancedTaxResponse = await response.json()
        console.log('🚀 Received advanced strategies:', data)
        setAdvancedStrategies(data.expertStrategies || [])
        setMultiYearPlan(data.multiYearPlan || null)
        setProfileSummary(data.profileSummary || null)
      } else {
        console.error('Advanced strategies API response not ok:', response.status, response.statusText)
      }
    } catch (error) {
      console.error('Error getting advanced strategies:', error)
    } finally {
      setLoadingAdvanced(false)
    }
  }

  // Streaming tax planning insights
  const getStreamingInsights = async (type: 'comprehensive' | 'specific_strategy' | 'implementation') => {
    setIsStreaming(true)
    setStreamingType(type)
    setStreamingInsights("")
    
    try {
      const minIncomeForStream = Math.max(safeAdjustedGrossIncome, 50000)
      
      console.log('📡 Starting streaming insights:', type)

      const response = await fetch('/api/ai/tax-planning-insights', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          adjustedGrossIncome: minIncomeForStream,
          filingStatus: safeFilingStatus,
          currentItemizedDeductions: safeCurrentItemizedDeductions,
          dependents: safeDependents,
          taxReturn: {
            totalIncome: taxReturn?.totalIncome || minIncomeForStream,
            deductionEntries: taxReturn?.deductionEntries || [],
            incomeEntries: taxReturn?.incomeEntries || []
          },
          personalInfo: {
            age: 35,
            spouseAge: undefined
          },
          requestType: type
        }),
      })

      if (!response.ok) {
        throw new Error(`Streaming API error: ${response.status}`)
      }

      const reader = response.body?.getReader()
      const decoder = new TextDecoder()
      let buffer = ''
      let partialRead = ''

      if (reader) {
        while (true) {
          const { done, value } = await reader.read()
          if (done) break

          partialRead += decoder.decode(value, { stream: true })
          let lines = partialRead.split('\n')
          partialRead = lines.pop() || ''

          for (const line of lines) {
            if (line.startsWith('data: ')) {
              const data = line.slice(6)
              if (data === '[DONE]') {
                setIsStreaming(false)
                return
              }
              
              try {
                const parsed = JSON.parse(data)
                if (parsed.content) {
                  buffer += parsed.content
                  setStreamingInsights(buffer)
                }
              } catch (e) {
                // Skip invalid JSON
              }
            }
          }
        }
      }
    } catch (error) {
      console.error('Error in streaming insights:', error)
      setIsStreaming(false)
    }
  }

  const baseTaxLiability = calculations[0]?.taxLiability || 0

  return (
    <Card className="border-2 border-blue-200 bg-blue-50/50">
      <CardHeader>
        <CardTitle className="flex items-center space-x-2">
          <Calculator className="h-5 w-5 text-blue-600" />
          <span>Interactive What-If Scenarios</span>
          <Sparkles className="h-4 w-4 text-yellow-500" />
        </CardTitle>
        <CardDescription>
          Explore how different deduction strategies could impact your taxes with AI-powered recommendations
        </CardDescription>
        
        {/* Debug Information Display */}
        {debugInfo && (process.env.NODE_ENV === 'development' || true) && (
          <Alert className="mt-2 bg-yellow-50 border-yellow-200">
            <Info className="h-4 w-4" />
            <AlertDescription>
              <details className="text-xs">
                <summary className="cursor-pointer font-medium">🔍 Debug Info (Click to expand)</summary>
                <div className="mt-2 space-y-1">
                  <div><strong>Income:</strong> ${debugInfo.safeAGI?.toLocaleString()} (original: ${debugInfo.originalAGI || 'N/A'})</div>
                  <div><strong>Itemized Deductions:</strong> ${debugInfo.safeItemized?.toLocaleString()}</div>
                  <div><strong>Filing Status:</strong> {debugInfo.filingStatus}</div>
                  <div><strong>Dependents:</strong> {debugInfo.dependentsCount}</div>
                  <div><strong>Tax Return ID:</strong> {debugInfo.taxReturn?.id || 'N/A'}</div>
                  <div><strong>Calculations Count:</strong> {calculations.length}</div>
                </div>
              </details>
            </AlertDescription>
          </Alert>
        )}
      </CardHeader>
      <CardContent>
        <Tabs value={activeTab} onValueChange={setActiveTab} className="w-full">
          <TabsList className="grid w-full grid-cols-6">
            <TabsTrigger value="scenarios">Scenarios</TabsTrigger>
            <TabsTrigger value="calculator">Calculator</TabsTrigger>
            <TabsTrigger value="ai-insights">AI Insights</TabsTrigger>
            <TabsTrigger value="advanced-strategies">Expert Strategies</TabsTrigger>
            <TabsTrigger value="multi-year">Multi-Year Plan</TabsTrigger>
            <TabsTrigger value="streaming-insights">Live Insights</TabsTrigger>
          </TabsList>

          <TabsContent value="scenarios" className="space-y-4">
            {/* Quick Calculation */}
            <Card>
              <CardHeader className="pb-3">
                <CardTitle className="text-lg flex items-center space-x-2">
                  <ArrowUpDown className="h-4 w-4" />
                  <span>Quick Impact Test</span>
                </CardTitle>
              </CardHeader>
              <CardContent>
                <div className="flex space-x-2">
                  <Input
                    type="number"
                    placeholder="Amount ($)"
                    value={quickAmount}
                    onChange={(e) => setQuickAmount(e.target.value)}
                    className="flex-1"
                  />
                  <Button type="button" onClick={handleQuickCalculation} disabled={!quickAmount}>
                    <Plus className="h-4 w-4 mr-1" />
                    Test Impact
                  </Button>
                </div>
              </CardContent>
            </Card>

            {/* Scenario Results */}
            <div className="space-y-3">
              <div className="flex items-center justify-between">
                <h4 className="font-medium">Tax Impact Scenarios</h4>
                <Badge variant="outline" className="bg-white">
                  Base Tax: ${baseTaxLiability.toLocaleString()}
                </Badge>
              </div>
              
              {calculations.map((scenario, index) => (
                <div 
                  key={scenario.custom ? scenario.id : index} 
                  className={`flex items-center justify-between p-4 rounded-lg border ${
                    index === 0 ? 'bg-gray-100 border-gray-300' : 'bg-white border-gray-200'
                  }`}
                >
                  <div className="flex-1">
                    <div className="flex items-center space-x-2">
                      <p className="font-medium">{scenario.scenario}</p>
                      {scenario.custom && (
                        <Button
                          type="button"
                          variant="ghost"
                          size="sm"
                          onClick={() => handleRemoveCustomScenario(scenario.id)}
                        >
                          <Minus className="h-3 w-3" />
                        </Button>
                      )}
                    </div>
                    <p className="text-sm text-gray-600">{scenario.description}</p>
                  </div>
                  <div className="text-right">
                    <p className="font-medium">
                      Tax: ${scenario.taxLiability.toLocaleString()}
                    </p>
                    {scenario.savings > 0 ? (
                      <div className="flex items-center text-green-600">
                        <TrendingDown className="h-3 w-3 mr-1" />
                        <span className="text-sm font-medium">
                          Saves: ${scenario.savings.toLocaleString()}
                        </span>
                      </div>
                    ) : scenario.savings < 0 ? (
                      <div className="flex items-center text-red-600">
                        <TrendingUp className="h-3 w-3 mr-1" />
                        <span className="text-sm font-medium">
                          Costs: ${Math.abs(scenario.savings).toLocaleString()}
                        </span>
                      </div>
                    ) : (
                      <span className="text-sm text-gray-500">No change</span>
                    )}
                  </div>
                </div>
              ))}
            </div>
          </TabsContent>

          <TabsContent value="calculator" className="space-y-4">
            {/* Custom Scenario Builder */}
            <Card>
              <CardHeader>
                <CardTitle className="text-lg">Create Custom Scenario</CardTitle>
                <CardDescription>
                  Build your own what-if scenario with specific deduction types and amounts
                </CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <Label htmlFor="scenarioName">Scenario Name</Label>
                    <Input
                      id="scenarioName"
                      placeholder="e.g., Increase charitable giving"
                      value={newScenario.name}
                      onChange={(e) => setNewScenario({...newScenario, name: e.target.value})}
                    />
                  </div>
                  <div>
                    <Label htmlFor="scenarioAmount">Additional Amount</Label>
                    <Input
                      id="scenarioAmount"
                      type="number"
                      placeholder="0.00"
                      value={newScenario.additionalAmount}
                      onChange={(e) => setNewScenario({...newScenario, additionalAmount: e.target.value})}
                    />
                  </div>
                </div>
                <div>
                  <Label htmlFor="deductionType">Deduction Type</Label>
                  <Select value={newScenario.deductionType} onValueChange={(value) => setNewScenario({...newScenario, deductionType: value})}>
                    <SelectTrigger>
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      {deductionTypes.map((type) => (
                        <SelectItem key={type.value} value={type.value}>
                          {type.label}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
                <Button 
                  type="button"
                  onClick={handleAddCustomScenario}
                  disabled={!newScenario.name || !newScenario.additionalAmount}
                  className="w-full"
                >
                  <Plus className="h-4 w-4 mr-2" />
                  Add Scenario
                </Button>
              </CardContent>
            </Card>

            {/* Custom Scenarios List */}
            {customScenarios.length > 0 && (
              <Card>
                <CardHeader>
                  <CardTitle className="text-lg">Your Custom Scenarios</CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="space-y-2">
                    {customScenarios.map((scenario) => (
                      <div key={scenario.id} className="flex items-center justify-between p-3 bg-gray-50 rounded-lg">
                        <div>
                          <p className="font-medium">{scenario.name}</p>
                          <p className="text-sm text-gray-600">
                            {deductionTypes.find(t => t.value === scenario.deductionType)?.label}: 
                            +${scenario.additionalAmount.toLocaleString()}
                          </p>
                        </div>
                        <Button
                          type="button"
                          variant="outline"
                          size="sm"
                          onClick={() => handleRemoveCustomScenario(scenario.id)}
                        >
                          <Minus className="h-4 w-4" />
                        </Button>
                      </div>
                    ))}
                  </div>
                </CardContent>
              </Card>
            )}
          </TabsContent>

          <TabsContent value="ai-insights" className="space-y-4">
            <div className="text-center">
              <Button 
                type="button"
                onClick={getAIRecommendations}
                disabled={loadingAI}
                size="lg"
                className="mb-4"
              >
                <Brain className="h-4 w-4 mr-2" />
                {loadingAI ? "Analyzing..." : "Get AI Tax Strategies"}
              </Button>
              <p className="text-sm text-gray-600">
                Get personalized tax optimization strategies based on your financial situation
              </p>
            </div>

            {aiRecommendations.length > 0 && (
              <div className="space-y-4">
                <h4 className="font-medium flex items-center space-x-2">
                  <Lightbulb className="h-4 w-4 text-yellow-500" />
                  <span>AI-Powered Tax Strategies</span>
                </h4>
                
                {aiRecommendations.map((recommendation, index) => (
                  <Card key={index} className="border-l-4 border-l-blue-500">
                    <CardContent className="pt-4">
                      <div className="flex items-start justify-between mb-2">
                        <h5 className="font-medium text-blue-900">{recommendation.strategy}</h5>
                        <div className="flex space-x-2">
                          <Badge variant="outline" className={
                            recommendation.difficulty === 'Easy' ? 'bg-green-100 text-green-800' :
                            recommendation.difficulty === 'Medium' ? 'bg-yellow-100 text-yellow-800' :
                            'bg-red-100 text-red-800'
                          }>
                            {recommendation.difficulty}
                          </Badge>
                          <Badge variant="secondary">
                            <Target className="h-3 w-3 mr-1" />
                            ${recommendation.potentialSavings.toLocaleString()}
                          </Badge>
                        </div>
                      </div>
                      <p className="text-gray-700 mb-2">{recommendation.description}</p>
                      <p className="text-sm text-gray-600">
                        <strong>Timeline:</strong> {recommendation.timeline}
                      </p>
                    </CardContent>
                  </Card>
                ))}
              </div>
            )}

            {!loadingAI && aiRecommendations.length === 0 && (
              <Alert>
                <Lightbulb className="h-4 w-4" />
                <AlertDescription>
                  Click "Get AI Tax Strategies" to receive personalized recommendations based on your tax situation.
                </AlertDescription>
              </Alert>
            )}
          </TabsContent>

          <TabsContent value="advanced-strategies" className="space-y-4">
            <div className="text-center">
              <Button 
                type="button"
                onClick={getAdvancedStrategies}
                disabled={loadingAdvanced}
                size="lg"
                className="mb-4"
              >
                <Brain className="h-4 w-4 mr-2" />
                {loadingAdvanced ? "Analyzing..." : "Get Expert Tax Strategies"}
              </Button>
              <p className="text-sm text-gray-600">
                Advanced tax optimization strategies based on expert US Tax Code knowledge
              </p>
            </div>

            {profileSummary && (
              <Card className="border-l-4 border-l-green-500">
                <CardContent className="pt-4">
                  <div className="grid grid-cols-3 gap-4 text-center">
                    <div>
                      <div className="text-2xl font-bold text-green-600">{profileSummary.taxBracket}</div>
                      <div className="text-sm text-gray-600">Current Tax Bracket</div>
                    </div>
                    <div>
                      <div className="text-2xl font-bold text-blue-600">
                        ${profileSummary.optimizationPotential?.toLocaleString() || '0'}
                      </div>
                      <div className="text-sm text-gray-600">Optimization Potential</div>
                    </div>
                    <div>
                      <div className="text-2xl font-bold text-orange-600">{profileSummary.riskFactors?.length || 0}</div>
                      <div className="text-sm text-gray-600">Risk Factors</div>
                    </div>
                  </div>
                  
                  {profileSummary.riskFactors?.length > 0 && (
                    <div className="mt-4 p-3 bg-orange-50 rounded-lg">
                      <h5 className="font-medium text-orange-800 mb-2">Tax Planning Considerations:</h5>
                      <ul className="text-sm text-orange-700 space-y-1">
                        {profileSummary.riskFactors.map((risk: string, index: number) => (
                          <li key={index}>• {risk}</li>
                        ))}
                      </ul>
                    </div>
                  )}
                </CardContent>
              </Card>
            )}

            {advancedStrategies.length > 0 && (
              <div className="space-y-4">
                <h4 className="font-medium flex items-center space-x-2">
                  <Sparkles className="h-4 w-4 text-purple-500" />
                  <span>Expert Tax Optimization Strategies</span>
                </h4>
                
                {advancedStrategies.map((strategy, index) => (
                  <Card key={index} className="border-l-4 border-l-purple-500 hover:shadow-md transition-shadow">
                    <CardContent className="pt-4">
                      <div className="flex items-start justify-between mb-3">
                        <div className="flex-1">
                          <div className="flex items-center space-x-2 mb-2">
                            <h5 className="font-medium text-purple-900">{strategy.strategy}</h5>
                            <Badge variant="outline" className={`
                              ${strategy.category === 'RETIREMENT' ? 'bg-blue-100 text-blue-800' :
                                strategy.category === 'INVESTMENTS' ? 'bg-green-100 text-green-800' :
                                strategy.category === 'BUSINESS' ? 'bg-orange-100 text-orange-800' :
                                strategy.category === 'CHARITABLE' ? 'bg-pink-100 text-pink-800' :
                                strategy.category === 'ESTATE' ? 'bg-indigo-100 text-indigo-800' :
                                'bg-gray-100 text-gray-800'}
                            `}>
                              {strategy.category}
                            </Badge>
                          </div>
                          <p className="text-gray-700 mb-3">{strategy.description}</p>
                        </div>
                        <div className="text-right ml-4">
                          <div className="text-lg font-bold text-green-600">
                            ${strategy.potentialSavings.toLocaleString()}
                          </div>
                          <div className="text-sm text-gray-600">Potential Savings</div>
                        </div>
                      </div>

                      <div className="grid grid-cols-2 gap-4 mb-3">
                        <div>
                          <Badge variant="secondary" className={
                            strategy.difficulty === 'Easy' ? 'bg-green-100 text-green-800' :
                            strategy.difficulty === 'Medium' ? 'bg-yellow-100 text-yellow-800' :
                            strategy.difficulty === 'Hard' ? 'bg-orange-100 text-orange-800' :
                            'bg-red-100 text-red-800'
                          }>
                            {strategy.difficulty}
                          </Badge>
                          <span className="text-sm text-gray-600 ml-2">Difficulty</span>
                        </div>
                        <div className="text-right">
                          <span className="text-sm font-medium">{strategy.timeline}</span>
                        </div>
                      </div>

                      <Button
                        type="button"
                        variant="outline"
                        size="sm"
                        onClick={() => setSelectedStrategy(selectedStrategy?.strategy === strategy.strategy ? null : strategy)}
                      >
                        {selectedStrategy?.strategy === strategy.strategy ? 'Hide Details' : 'Show Implementation Details'}
                      </Button>

                      {selectedStrategy?.strategy === strategy.strategy && (
                        <div className="mt-4 space-y-3 border-t pt-3">
                          <div>
                            <h6 className="font-medium text-gray-800 mb-2">Implementation Steps:</h6>
                            <ol className="text-sm text-gray-700 space-y-1">
                              {strategy.implementationSteps.map((step, stepIndex) => (
                                <li key={stepIndex} className="flex">
                                  <span className="mr-2">{stepIndex + 1}.</span>
                                  <span>{step}</span>
                                </li>
                              ))}
                            </ol>
                          </div>
                          
                          <div>
                            <h6 className="font-medium text-gray-800 mb-2">Requirements:</h6>
                            <ul className="text-sm text-gray-700 space-y-1">
                              {strategy.requirements.map((req, reqIndex) => (
                                <li key={reqIndex}>• {req}</li>
                              ))}
                            </ul>
                          </div>

                          <div>
                            <h6 className="font-medium text-gray-800 mb-2">Important Considerations:</h6>
                            <ul className="text-sm text-gray-700 space-y-1">
                              {strategy.considerations.map((consideration, consIndex) => (
                                <li key={consIndex}>• {consideration}</li>
                              ))}
                            </ul>
                          </div>

                          {strategy.taxCodeReference && (
                            <div className="text-xs text-gray-500 italic">
                              Tax Code Reference: {strategy.taxCodeReference}
                            </div>
                          )}
                        </div>
                      )}
                    </CardContent>
                  </Card>
                ))}
              </div>
            )}
          </TabsContent>

          <TabsContent value="multi-year" className="space-y-4">
            {!multiYearPlan ? (
              <div className="text-center">
                <Button 
                  type="button"
                  onClick={getAdvancedStrategies}
                  disabled={loadingAdvanced}
                  size="lg"
                  className="mb-4"
                >
                  <Target className="h-4 w-4 mr-2" />
                  {loadingAdvanced ? "Developing Plan..." : "Create Multi-Year Tax Plan"}
                </Button>
                <p className="text-sm text-gray-600">
                  Strategic tax planning across multiple years for maximum optimization
                </p>
              </div>
            ) : (
              <div className="space-y-6">
                <div className="grid grid-cols-3 gap-4">
                  <Card className="border-l-4 border-l-blue-500">
                    <CardHeader className="pb-3">
                      <CardTitle className="text-lg text-blue-900">Current Year (2025)</CardTitle>
                      <CardDescription>Immediate action items</CardDescription>
                    </CardHeader>
                    <CardContent>
                      <div className="space-y-2">
                        {multiYearPlan?.currentYear?.slice(0, 3).map((strategy: AdvancedStrategy, index: number) => (
                          <div key={index} className="p-2 bg-blue-50 rounded text-sm">
                            <div className="font-medium">{strategy.strategy}</div>
                            <div className="text-green-600 font-medium">${strategy.potentialSavings.toLocaleString()}</div>
                          </div>
                        ))}
                      </div>
                    </CardContent>
                  </Card>

                  <Card className="border-l-4 border-l-green-500">
                    <CardHeader className="pb-3">
                      <CardTitle className="text-lg text-green-900">Next Year (2026)</CardTitle>
                      <CardDescription>Planning ahead</CardDescription>
                    </CardHeader>
                    <CardContent>
                      <div className="space-y-2">
                        {multiYearPlan?.nextYear?.slice(0, 3).map((strategy: AdvancedStrategy, index: number) => (
                          <div key={index} className="p-2 bg-green-50 rounded text-sm">
                            <div className="font-medium">{strategy.strategy}</div>
                            <div className="text-green-600 font-medium">${strategy.potentialSavings.toLocaleString()}</div>
                          </div>
                        ))}
                      </div>
                    </CardContent>
                  </Card>

                  <Card className="border-l-4 border-l-purple-500">
                    <CardHeader className="pb-3">
                      <CardTitle className="text-lg text-purple-900">Long-term (5+ years)</CardTitle>
                      <CardDescription>Strategic positioning</CardDescription>
                    </CardHeader>
                    <CardContent>
                      <div className="space-y-2">
                        {multiYearPlan?.longTerm?.slice(0, 3).map((strategy: AdvancedStrategy, index: number) => (
                          <div key={index} className="p-2 bg-purple-50 rounded text-sm">
                            <div className="font-medium">{strategy.strategy}</div>
                            <div className="text-green-600 font-medium">${strategy.potentialSavings.toLocaleString()}</div>
                          </div>
                        ))}
                      </div>
                    </CardContent>
                  </Card>
                </div>

                <Card className="bg-gradient-to-r from-blue-50 to-purple-50">
                  <CardHeader>
                    <CardTitle className="text-lg">Multi-Year Optimization Summary</CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="grid grid-cols-2 gap-6">
                      <div>
                        <h6 className="font-medium mb-2">Total Potential Savings Over 5 Years:</h6>
                        <div className="text-2xl font-bold text-green-600">
                          ${(
                            (multiYearPlan?.currentYear?.reduce((sum: number, s: AdvancedStrategy) => sum + s.potentialSavings, 0) || 0) +
                            (multiYearPlan?.nextYear?.reduce((sum: number, s: AdvancedStrategy) => sum + s.potentialSavings, 0) || 0) +
                            (multiYearPlan?.longTerm?.reduce((sum: number, s: AdvancedStrategy) => sum + s.potentialSavings, 0) || 0)
                          ).toLocaleString()}
                        </div>
                      </div>
                      <div>
                        <h6 className="font-medium mb-2">Key Strategic Focus:</h6>
                        <ul className="text-sm space-y-1">
                          <li>• Retirement account optimization</li>
                          <li>• Tax bracket management</li>
                          <li>• Investment timing strategies</li>
                          <li>• Estate planning integration</li>
                        </ul>
                      </div>
                    </div>
                  </CardContent>
                </Card>
              </div>
            )}
          </TabsContent>

          <TabsContent value="streaming-insights" className="space-y-4">
            <div className="grid grid-cols-3 gap-2 mb-4">
              <Button 
                type="button"
                onClick={() => getStreamingInsights('comprehensive')}
                disabled={isStreaming}
                variant={streamingType === 'comprehensive' ? 'default' : 'outline'}
                size="sm"
              >
                Comprehensive Analysis
              </Button>
              <Button 
                type="button"
                onClick={() => getStreamingInsights('specific_strategy')}
                disabled={isStreaming}
                variant={streamingType === 'specific_strategy' ? 'default' : 'outline'}
                size="sm"
              >
                Strategy Deep-Dive
              </Button>
              <Button 
                type="button"
                onClick={() => getStreamingInsights('implementation')}
                disabled={isStreaming}
                variant={streamingType === 'implementation' ? 'default' : 'outline'}
                size="sm"
              >
                Implementation Plan
              </Button>
            </div>

            {isStreaming && (
              <Alert className="bg-blue-50 border-blue-200">
                <Brain className="h-4 w-4 animate-pulse" />
                <AlertDescription>
                  AI is analyzing your tax situation and generating personalized insights...
                </AlertDescription>
              </Alert>
            )}

            {streamingInsights && (
              <Card className="border-l-4 border-l-blue-500">
                <CardHeader>
                  <CardTitle className="flex items-center space-x-2">
                    <Brain className="h-5 w-5 text-blue-600" />
                    <span>Live Tax Planning Insights</span>
                    {isStreaming && <div className="w-2 h-2 bg-green-500 rounded-full animate-pulse"></div>}
                  </CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="prose prose-sm max-w-none">
                    <div className="whitespace-pre-wrap text-gray-700 leading-relaxed">
                      {streamingInsights}
                      {isStreaming && <span className="animate-pulse">▋</span>}
                    </div>
                  </div>
                </CardContent>
              </Card>
            )}

            {!streamingInsights && !isStreaming && (
              <Alert>
                <Lightbulb className="h-4 w-4" />
                <AlertDescription>
                  Select an analysis type above to receive real-time, personalized tax planning insights powered by AI.
                </AlertDescription>
              </Alert>
            )}
          </TabsContent>
        </Tabs>
      </CardContent>
    </Card>
  )
}
