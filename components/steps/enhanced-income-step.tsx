"use client"

import { useState, useEffect } from "react"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select"
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card"
import { Badge } from "@/components/ui/badge"
import { Alert, AlertDescription } from "@/components/ui/alert"
import { DocumentProcessor } from "@/components/document-processor"
import { NameValidationDialog } from "@/components/name-validation-dialog"
import { 
  DollarSign, 
  Plus, 
  Trash2, 
  ArrowRight, 
  ArrowLeft, 
  Info, 
  FileText, 
  CheckCircle,
  Upload,
  Sparkles
} from "lucide-react"
import { validateNames, extractNamesFromDocument, type NameValidationResult } from "@/lib/name-validation"

interface EnhancedIncomeStepProps {
  taxReturn: any
  onUpdate: (data: any) => Promise<any>
  onAutoSave: (data: any) => Promise<any>
  onCompleteStep: (data: any) => Promise<any>
  onNext: () => void
  onPrev: () => void
  onMarkUnsaved: () => void
  loading: boolean
  saving: boolean
  autoSaving: boolean
  hasUnsavedChanges: boolean
  lastSaved: Date | null
}

const incomeTypes = [
  { value: "W2_WAGES", label: "W-2 Wages" },
  { value: "INTEREST", label: "Interest Income" },
  { value: "DIVIDENDS", label: "Dividends" },
  { value: "UNEMPLOYMENT", label: "Unemployment Compensation" },
  { value: "RETIREMENT_DISTRIBUTIONS", label: "Retirement Distributions" },
  { value: "SOCIAL_SECURITY", label: "Social Security Benefits" },
  { value: "OTHER_INCOME", label: "Other Income" },
]

interface AutoPopulatedEntry {
  incomeType: string
  amount: string
  description: string
  employerName: string
  employerEIN: string
  payerName: string
  payerTIN: string
  federalTaxWithheld: string
  isAutoPopulated: boolean
  documentId?: string
  documentType?: string
  confidence?: number
}

export function EnhancedIncomeStep({ 
  taxReturn, 
  onUpdate, 
  onAutoSave, 
  onCompleteStep, 
  onNext, 
  onPrev, 
  onMarkUnsaved, 
  loading, 
  saving, 
  autoSaving, 
  hasUnsavedChanges, 
  lastSaved 
}: EnhancedIncomeStepProps) {
  const [incomeEntries, setIncomeEntries] = useState(taxReturn.incomeEntries || [])
  const [pendingAutoEntries, setPendingAutoEntries] = useState<AutoPopulatedEntry[]>([])
  const [newEntry, setNewEntry] = useState({
    incomeType: "",
    amount: "",
    description: "",
    employerName: "",
    employerEIN: "",
    payerName: "",
    payerTIN: "",
    federalTaxWithheld: "",
  })
  
  // Name validation state
  const [nameValidationDialog, setNameValidationDialog] = useState<{
    isOpen: boolean
    validationResult: NameValidationResult | null
    documentType: string
    extractedData: any
  }>({
    isOpen: false,
    validationResult: null,
    documentType: '',
    extractedData: null
  })

  const totalIncome = incomeEntries.reduce((sum: number, entry: any) => 
    sum + parseFloat(entry.amount || 0), 0
  )

  // Auto-save functionality with debouncing
  useEffect(() => {
    const timer = setTimeout(() => {
      if (hasUnsavedChanges && !autoSaving && incomeEntries.length > 0) {
        onAutoSave({ 
          totalIncome: totalIncome,
          adjustedGrossIncome: totalIncome 
        })
      }
    }, 3000) // Auto-save after 3 seconds of inactivity

    return () => clearTimeout(timer)
  }, [incomeEntries, totalIncome, hasUnsavedChanges, autoSaving, onAutoSave])

 const handleDocumentProcessed = async (extractedData: any) => {
  console.log('🔍 [CALLBACK] handleDocumentProcessed called with:', extractedData)
  
  // GUARD: Don't process if documentType is missing or if we already have pending entries
  if (!extractedData?.documentType && pendingAutoEntries.length > 0) {
    console.log('🔍 [GUARD] Skipping processing - no documentType and entries already exist')
    return
  }
  
  // GUARD: Don't process if documentType is undefined
  if (!extractedData?.documentType) {
    console.log('🔍 [GUARD] Skipping processing - documentType is undefined')
    return
  }
  
  // CRITICAL FIX: Process and display extracted data FIRST
  console.log('🔍 [CALLBACK] Converting extracted data to income entries immediately...')
  const autoEntries = convertExtractedDataToIncomeEntries(extractedData)
  console.log('🔍 [CALLBACK] Auto entries created:', autoEntries)
  
  // Set the pending entries immediately so they appear in the UI
  setPendingAutoEntries(autoEntries)
  console.log('🔍 [CALLBACK] pendingAutoEntries set immediately - UI should show extracted data now')
  
  // THEN handle name validation separately (non-blocking)
  setTimeout(() => {
    console.log('🔍 [CALLBACK] Starting name validation process...')
    
    // Extract names for validation
    const extractedNames = extractNamesFromDocument(extractedData?.extractedData || extractedData)
    const profileNames = {
      firstName: taxReturn.firstName || '',
      lastName: taxReturn.lastName || '',
      spouseFirstName: taxReturn.spouseFirstName,
      spouseLastName: taxReturn.spouseLastName
    }

    // Validate names
    const validationResult = validateNames(profileNames, extractedNames)
    console.log('🔍 [CALLBACK] Name validation result:', validationResult)
    
    // Only show name validation dialog if there are actual mismatches
    if (!validationResult.isValid && validationResult.mismatches.length > 0) {
      console.log('🔍 [CALLBACK] Name mismatches found, showing validation dialog')
      setNameValidationDialog({
        isOpen: true,
        validationResult,
        documentType: extractedData?.documentType || 'tax document',
        extractedData
      })
    } else {
      console.log('🔍 [CALLBACK] Names validated successfully, no dialog needed')
    }
  }, 100) // Small delay to ensure UI updates first
}

  const handleNameValidationConfirm = async (proceedWithMismatches: boolean) => {
    console.log('🔍 [NAME_VALIDATION] handleNameValidationConfirm called with:', proceedWithMismatches)
    
    if (!proceedWithMismatches && !nameValidationDialog.validationResult?.isValid) {
      // User chose to update profile first - close dialog but keep extracted data visible
      console.log('🔍 [NAME_VALIDATION] User chose to update profile first')
      setNameValidationDialog((prev: typeof nameValidationDialog) => ({ ...prev, isOpen: false }))
      return
    }

    // Close dialog - extracted data is already visible
    console.log('🔍 [NAME_VALIDATION] User confirmed to proceed, closing dialog')
    setNameValidationDialog((prev: typeof nameValidationDialog) => ({ ...prev, isOpen: false }))
  }

  const convertExtractedDataToIncomeEntries = (extractedData: any): AutoPopulatedEntry[] => {  
    const entries: AutoPopulatedEntry[] = []
    const data = extractedData?.extractedData || extractedData
    
    // DEBUG: Log all available fields for troubleshooting
    console.log('🔍 [CONVERT] Available fields in extracted data:', Object.keys(data || {}));
    console.log('🔍 [CONVERT] Document type:', extractedData?.documentType);
    
    // DEBUG: Log field values for critical 1099-MISC fields
    if (extractedData?.documentType === 'FORM_1099_MISC') {
      const criticalFields = ['otherIncome', 'fishingBoatProceeds', 'medicalHealthPayments', 'rents', 'royalties', 'nonemployeeCompensation'];
      criticalFields.forEach(field => {
        console.log(`🔍 [CONVERT] ${field}:`, data[field]);
      });
    }

    // Handle W-2 data
    if (extractedData?.documentType === 'W2' || data?.wages) {
      console.log('🔍 [CONVERT] Processing W-2 data...');
      entries.push({
        incomeType: 'W2_WAGES',
        amount: cleanAmount(data.wages || '0'),
        description: `W-2 Wages from ${data.employerName || 'Employer'}`,
        employerName: data.employerName || '',
        employerEIN: data.employerEIN || '',
        payerName: '',
        payerTIN: '',
        federalTaxWithheld: cleanAmount(data.federalTaxWithheld || '0'),
        isAutoPopulated: true,
        documentId: extractedData?.documentId,
        documentType: 'W2',
        confidence: extractedData?.confidence || 0.85
      })
    }

    // Handle 1099-INT data
    if (extractedData?.documentType === 'FORM_1099_INT' || data?.interestIncome) {
      console.log('🔍 [CONVERT] Processing 1099-INT data...');
      entries.push({
        incomeType: 'INTEREST',
        amount: cleanAmount(data.interestIncome || '0'),
        description: `Interest Income from ${data.payerName || 'Financial Institution'}`,
        employerName: '',
        employerEIN: '',
        payerName: data.payerName || '',
        payerTIN: data.payerTIN || '',
        federalTaxWithheld: cleanAmount(data.federalTaxWithheld || '0'),
        isAutoPopulated: true,
        documentId: extractedData?.documentId,
        documentType: 'FORM_1099_INT',
        confidence: extractedData?.confidence || 0.85
      })
    }

    // Handle 1099-DIV data
    if (extractedData?.documentType === 'FORM_1099_DIV' || data?.ordinaryDividends) {
      console.log('🔍 [CONVERT] Processing 1099-DIV data...');
      entries.push({
        incomeType: 'DIVIDENDS',
        amount: cleanAmount(data.ordinaryDividends || '0'),
        description: `Dividend Income from ${data.payerName || 'Investment Account'}`,
        employerName: '',
        employerEIN: '',
        payerName: data.payerName || '',
        payerTIN: data.payerTIN || '',
        federalTaxWithheld: cleanAmount(data.federalTaxWithheld || '0'),
        isAutoPopulated: true,
        documentId: extractedData?.documentId,
        documentType: 'FORM_1099_DIV',
        confidence: extractedData?.confidence || 0.85
      })
    }

    // Enhanced 1099-MISC data handling - Create separate entries for different income types
    if (extractedData?.documentType === 'FORM_1099_MISC') {
      console.log('🔍 [CONVERT] Processing 1099-MISC document with data:', data);
      
      const payerName = data.payerName || 'Payer'
      const payerTIN = data.payerTIN || ''
      const federalTaxWithheld = cleanAmount(data.federalTaxWithheld || '0')
      
      // Rents (Box 1)
      if (data.rents && parseFloat(cleanAmount(data.rents)) > 0) {
        entries.push({
          incomeType: 'OTHER_INCOME',
          amount: cleanAmount(data.rents),
          description: `Rental Income from ${payerName} (1099-MISC Box 1)`,
          employerName: '',
          employerEIN: '',
          payerName: payerName,
          payerTIN: payerTIN,
          federalTaxWithheld: federalTaxWithheld,
          isAutoPopulated: true,
          documentId: extractedData?.documentId,
          documentType: 'FORM_1099_MISC',
          confidence: extractedData?.confidence || 0.85
        })
      }

      // Royalties (Box 2)
      if (data.royalties && parseFloat(cleanAmount(data.royalties)) > 0) {
        entries.push({
          incomeType: 'OTHER_INCOME',
          amount: cleanAmount(data.royalties),
          description: `Royalty Income from ${payerName} (1099-MISC Box 2)`,
          employerName: '',
          employerEIN: '',
          payerName: payerName,
          payerTIN: payerTIN,
          federalTaxWithheld: federalTaxWithheld,
          isAutoPopulated: true,
          documentId: extractedData?.documentId,
          documentType: 'FORM_1099_MISC',
          confidence: extractedData?.confidence || 0.85
        })
      }

      // Other Income (Box 3)
      if (data.otherIncome && parseFloat(cleanAmount(data.otherIncome)) > 0) {
        entries.push({
          incomeType: 'OTHER_INCOME',
          amount: cleanAmount(data.otherIncome),
          description: `Other Income from ${payerName} (1099-MISC Box 3)`,
          employerName: '',
          employerEIN: '',
          payerName: payerName,
          payerTIN: payerTIN,
          federalTaxWithheld: federalTaxWithheld,
          isAutoPopulated: true,
          documentId: extractedData?.documentId,
          documentType: 'FORM_1099_MISC',
          confidence: extractedData?.confidence || 0.85
        })
      }

      // Fishing Boat Proceeds (Box 5)
      if (data.fishingBoatProceeds && parseFloat(cleanAmount(data.fishingBoatProceeds)) > 0) {
        entries.push({
          incomeType: 'OTHER_INCOME',
          amount: cleanAmount(data.fishingBoatProceeds),
          description: `Fishing Boat Proceeds from ${payerName} (1099-MISC Box 5)`,
          employerName: '',
          employerEIN: '',
          payerName: payerName,
          payerTIN: payerTIN,
          federalTaxWithheld: federalTaxWithheld,
          isAutoPopulated: true,
          documentId: extractedData?.documentId,
          documentType: 'FORM_1099_MISC',
          confidence: extractedData?.confidence || 0.85
        })
      }

      // Medical and Health Care Payments (Box 6)
      if (data.medicalHealthPayments && parseFloat(cleanAmount(data.medicalHealthPayments)) > 0) {
        entries.push({
          incomeType: 'OTHER_INCOME',
          amount: cleanAmount(data.medicalHealthPayments),
          description: `Medical/Health Care Payments from ${payerName} (1099-MISC Box 6)`,
          employerName: '',
          employerEIN: '',
          payerName: payerName,
          payerTIN: payerTIN,
          federalTaxWithheld: federalTaxWithheld,
          isAutoPopulated: true,
          documentId: extractedData?.documentId,
          documentType: 'FORM_1099_MISC',
          confidence: extractedData?.confidence || 0.85
        })
      }

      // Nonemployee Compensation (Box 7)
      if (data.nonemployeeCompensation && parseFloat(cleanAmount(data.nonemployeeCompensation)) > 0) {
        entries.push({
          incomeType: 'OTHER_INCOME',
          amount: cleanAmount(data.nonemployeeCompensation),
          description: `Nonemployee Compensation from ${payerName} (1099-MISC Box 7)`,
          employerName: '',
          employerEIN: '',
          payerName: payerName,
          payerTIN: payerTIN,
          federalTaxWithheld: federalTaxWithheld,
          isAutoPopulated: true,
          documentId: extractedData?.documentId,
          documentType: 'FORM_1099_MISC',
          confidence: extractedData?.confidence || 0.85
        })
      }

      // Substitute Payments in Lieu of Dividends or Interest (Box 8)
      if (data.substitutePayments && parseFloat(cleanAmount(data.substitutePayments)) > 0) {
        entries.push({
          incomeType: 'OTHER_INCOME',
          amount: cleanAmount(data.substitutePayments),
          description: `Substitute Payments from ${payerName} (1099-MISC Box 8)`,
          employerName: '',
          employerEIN: '',
          payerName: payerName,
          payerTIN: payerTIN,
          federalTaxWithheld: federalTaxWithheld,
          isAutoPopulated: true,
          documentId: extractedData?.documentId,
          documentType: 'FORM_1099_MISC',
          confidence: extractedData?.confidence || 0.85
        })
      }

      // Crop Insurance Proceeds (Box 9)
      if (data.cropInsuranceProceeds && parseFloat(cleanAmount(data.cropInsuranceProceeds)) > 0) {
        entries.push({
          incomeType: 'OTHER_INCOME',
          amount: cleanAmount(data.cropInsuranceProceeds),
          description: `Crop Insurance Proceeds from ${payerName} (1099-MISC Box 9)`,
          employerName: '',
          employerEIN: '',
          payerName: payerName,
          payerTIN: payerTIN,
          federalTaxWithheld: federalTaxWithheld,
          isAutoPopulated: true,
          documentId: extractedData?.documentId,
          documentType: 'FORM_1099_MISC',
          confidence: extractedData?.confidence || 0.85
        })
      }

      // Gross Proceeds Paid to an Attorney (Box 10)
      if (data.grossProceedsAttorney && parseFloat(cleanAmount(data.grossProceedsAttorney)) > 0) {
        entries.push({
          incomeType: 'OTHER_INCOME',
          amount: cleanAmount(data.grossProceedsAttorney),
          description: `Attorney Gross Proceeds from ${payerName} (1099-MISC Box 10)`,
          employerName: '',
          employerEIN: '',
          payerName: payerName,
          payerTIN: payerTIN,
          federalTaxWithheld: federalTaxWithheld,
          isAutoPopulated: true,
          documentId: extractedData?.documentId,
          documentType: 'FORM_1099_MISC',
          confidence: extractedData?.confidence || 0.85
        })
      }

      // Section 409A Deferrals (Box 12)
      if (data.section409ADeferrals && parseFloat(cleanAmount(data.section409ADeferrals)) > 0) {
        entries.push({
          incomeType: 'OTHER_INCOME',
          amount: cleanAmount(data.section409ADeferrals),
          description: `Section 409A Deferrals from ${payerName} (1099-MISC Box 12)`,
          employerName: '',
          employerEIN: '',
          payerName: payerName,
          payerTIN: payerTIN,
          federalTaxWithheld: federalTaxWithheld,
          isAutoPopulated: true,
          documentId: extractedData?.documentId,
          documentType: 'FORM_1099_MISC',
          confidence: extractedData?.confidence || 0.85
        })
      }

      // Excess Golden Parachute Payments (Box 13)
      if (data.excessGoldenParachutePayments && parseFloat(cleanAmount(data.excessGoldenParachutePayments)) > 0) {
        entries.push({
          incomeType: 'OTHER_INCOME',
          amount: cleanAmount(data.excessGoldenParachutePayments),
          description: `Excess Golden Parachute Payments from ${payerName} (1099-MISC Box 13)`,
          employerName: '',
          employerEIN: '',
          payerName: payerName,
          payerTIN: payerTIN,
          federalTaxWithheld: federalTaxWithheld,
          isAutoPopulated: true,
          documentId: extractedData?.documentId,
          documentType: 'FORM_1099_MISC',
          confidence: extractedData?.confidence || 0.85
        })
      }

      // Nonqualified Deferred Compensation (Box 14)
      if (data.nonqualifiedDeferredCompensation && parseFloat(cleanAmount(data.nonqualifiedDeferredCompensation)) > 0) {
        entries.push({
          incomeType: 'OTHER_INCOME',
          amount: cleanAmount(data.nonqualifiedDeferredCompensation),
          description: `Nonqualified Deferred Compensation from ${payerName} (1099-MISC Box 14)`,
          employerName: '',
          employerEIN: '',
          payerName: payerName,
          payerTIN: payerTIN,
          federalTaxWithheld: federalTaxWithheld,
          isAutoPopulated: true,
          documentId: extractedData?.documentId,
          documentType: 'FORM_1099_MISC',
          confidence: extractedData?.confidence || 0.85
        })
      }
    }

    // Handle 1099-NEC data
    if (extractedData?.documentType === 'FORM_1099_NEC' || data?.nonemployeeCompensation) {
      console.log('🔍 [CONVERT] Processing 1099-NEC data...');
      entries.push({
        incomeType: 'OTHER_INCOME',
        amount: cleanAmount(data.nonemployeeCompensation || '0'),
        description: `1099-NEC Nonemployee Compensation from ${data.payerName || 'Payer'}`,
        employerName: '',
        employerEIN: '',
        payerName: data.payerName || '',
        payerTIN: data.payerTIN || '',
        federalTaxWithheld: cleanAmount(data.federalTaxWithheld || '0'),
        isAutoPopulated: true,
        documentId: extractedData?.documentId,
        documentType: 'FORM_1099_NEC',
        confidence: extractedData?.confidence || 0.85
      })
    }

    console.log(`🔍 [CONVERT] Total entries created: ${entries.length}`);
    return entries.filter(entry => parseFloat(entry.amount) > 0)
  }

  const cleanAmount = (amount: string | number): string => {
    if (!amount && amount !== 0) return '0'
    // Handle both string and number inputs
    const amountStr = amount.toString()
    // Remove currency symbols, commas, and extra spaces
    return amountStr.replace(/[$,\s]/g, '').replace(/[^\d.-]/g, '') || '0'
  }

  const handleAcceptAutoEntry = async (entry: AutoPopulatedEntry, index: number) => {
    console.log('🔍 [ACCEPT] Accepting auto entry:', entry);
    try {
      const entryData = {
        incomeType: entry.incomeType,
        amount: parseFloat(entry.amount),
        description: entry.description,
        employerName: entry.employerName,
        employerEIN: entry.employerEIN,
        payerName: entry.payerName,
        payerTIN: entry.payerTIN,
        federalTaxWithheld: parseFloat(entry.federalTaxWithheld || '0'),
        documentId: entry.documentId, // Include documentId for linking
      }

      const response = await fetch(`/api/tax-returns/${taxReturn.id}/income`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(entryData),
      })

      if (response.ok) {
        const savedEntry = await response.json()
        // Mark as auto-populated for visual indication
        savedEntry.isAutoPopulated = true
        savedEntry.confidence = entry.confidence
        
        setIncomeEntries((prev: any[]) => [...prev, savedEntry])
        setPendingAutoEntries((prev: AutoPopulatedEntry[]) => prev.filter((_, i) => i !== index))
        console.log('✅ [ACCEPT] Entry accepted and saved successfully');
      }
    } catch (error) {
      console.error("❌ [ACCEPT] Error adding auto-populated income entry:", error)
    }
  }

  const handleRejectAutoEntry = (index: number) => {
    console.log('🔍 [REJECT] Rejecting auto entry at index:', index);
    setPendingAutoEntries((prev: AutoPopulatedEntry[]) => prev.filter((_, i) => i !== index))
  }

  const handleAcceptAllAutoEntries = async () => {
    console.log('🔍 [ACCEPT_ALL] Accepting all auto entries:', pendingAutoEntries.length);
    for (let i = 0; i < pendingAutoEntries.length; i++) {
      await handleAcceptAutoEntry(pendingAutoEntries[i], 0) // Always use index 0 since array shrinks
    }
  }

  const handleAddEntry = async () => {
    if (!newEntry.incomeType || !newEntry.amount) return

    const entry = {
      incomeType: newEntry.incomeType,
      amount: parseFloat(newEntry.amount),
      description: newEntry.description,
      employerName: newEntry.employerName,
      employerEIN: newEntry.employerEIN,
      payerName: newEntry.payerName,
      payerTIN: newEntry.payerTIN,
    }

    try {
      const response = await fetch(`/api/tax-returns/${taxReturn.id}/income`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(entry),
      })

      if (response.ok) {
        const savedEntry = await response.json()
        setIncomeEntries([...incomeEntries, savedEntry])
        onMarkUnsaved()
        setNewEntry({
          incomeType: "",
          amount: "",
          description: "",
          employerName: "",
          employerEIN: "",
          payerName: "",
          payerTIN: "",
          federalTaxWithheld: "",
        })
      }
    } catch (error) {
      console.error("Error adding income entry:", error)
    }
  }

  const handleDeleteEntry = async (entryId: string) => {
    try {
      const response = await fetch(`/api/tax-returns/${taxReturn.id}/income/${entryId}`, {
        method: "DELETE",
      })

      if (response.ok) {
        setIncomeEntries(incomeEntries.filter((entry: any) => entry.id !== entryId))
        onMarkUnsaved()
      }
    } catch (error) {
      console.error("Error deleting income entry:", error)
    }
  }

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    await onUpdate({ 
      totalIncome: totalIncome,
      adjustedGrossIncome: totalIncome // For Stage 1, AGI = Total Income
    })
    onNext()
  }

  const handleSaveAndContinue = async (e: React.FormEvent) => {
    e.preventDefault()
    await onCompleteStep({ 
      totalIncome: totalIncome,
      adjustedGrossIncome: totalIncome
    })
    onNext()
  }

  const handleSaveOnly = async () => {
    await onAutoSave({ 
      totalIncome: totalIncome,
      adjustedGrossIncome: totalIncome
    })
  }

  return (
    <form onSubmit={handleSubmit} className="space-y-6">
      <div className="space-y-6">
        <Alert>
          <Info className="h-4 w-4" />
          <AlertDescription>
            Upload your tax documents to automatically extract income information, or add income sources manually. Documents are validated for accuracy and name matching.
          </AlertDescription>
        </Alert>

        {/* Document Upload Section */}
        <Card>
          <CardHeader>
            <CardTitle className="flex items-center space-x-2">
              <Upload className="h-5 w-5" />
              <span>Upload Tax Documents</span>
            </CardTitle>
            <CardDescription>
              Upload W-2s, 1099s, and other tax documents to automatically populate your income information
            </CardDescription>
          </CardHeader>
          <CardContent>
            <DocumentProcessor
              taxReturnId={taxReturn.id}
              onDocumentProcessed={handleDocumentProcessed}
              onUploadMoreRequested={() => {
                // Clear pending auto entries when user wants to upload more documents
                setPendingAutoEntries([])
                // Close any open name validation dialog
                setNameValidationDialog(prev => ({ ...prev, isOpen: false }))
              }}
            />
          </CardContent>
        </Card>

        {/* Auto-populated entries awaiting approval */}
        {pendingAutoEntries.length > 0 && (
          <Card className="border-blue-200 bg-blue-50/50">
            <CardHeader>
              <CardTitle className="flex items-center space-x-2">
                <Sparkles className="h-5 w-5 text-blue-600" />
                <span>Extracted Income Data</span>
              </CardTitle>
              <CardDescription>
                Review the income information extracted from your documents. Accept individual entries or all at once.
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="flex justify-end mb-4">
                <Button
                  type="button"
                  onClick={handleAcceptAllAutoEntries}
                  className="bg-blue-600 hover:bg-blue-700"
                >
                  <CheckCircle className="mr-2 h-4 w-4" />
                  Accept All ({pendingAutoEntries.length})
                </Button>
              </div>
              
              {pendingAutoEntries.map((entry, index) => (
                <div key={index} className="bg-white p-4 rounded-lg border border-blue-200">
                  <div className="flex items-start justify-between">
                    <div className="flex-1">
                      <div className="flex items-center space-x-2 mb-2">
                        <Badge variant="secondary" className="bg-blue-100 text-blue-800">
                          {incomeTypes.find(t => t.value === entry.incomeType)?.label}
                        </Badge>
                        <Badge variant="outline" className="text-xs">
                          <Sparkles className="mr-1 h-3 w-3" />
                          Auto-extracted
                        </Badge>
                        {entry.confidence && (
                          <Badge variant="outline" className="text-xs">
                            {Math.round(entry.confidence * 100)}% confidence
                          </Badge>
                        )}
                      </div>
                      <div className="font-medium text-lg mb-1">
                        ${parseFloat(entry.amount).toLocaleString()}
                      </div>
                      {entry.description && (
                        <p className="text-sm text-gray-600 mb-2">{entry.description}</p>
                      )}
                      <div className="text-xs text-gray-500 space-y-1">
                        {entry.employerName && <p>Employer: {entry.employerName}</p>}
                        {entry.employerEIN && <p>Employer EIN: {entry.employerEIN}</p>}
                        {entry.payerName && <p>Payer: {entry.payerName}</p>}
                        {entry.payerTIN && <p>Payer TIN: {entry.payerTIN}</p>}
                        {parseFloat(entry.federalTaxWithheld) > 0 && (
                          <p>Federal Tax Withheld: ${parseFloat(entry.federalTaxWithheld).toLocaleString()}</p>
                        )}
                      </div>
                    </div>
                    <div className="flex space-x-2 ml-4">
                      <Button
                        type="button"
                        size="sm"
                        onClick={() => handleAcceptAutoEntry(entry, index)}
                        className="bg-green-600 hover:bg-green-700"
                      >
                        <CheckCircle className="h-4 w-4" />
                      </Button>
                      <Button
                        type="button"
                        variant="outline"
                        size="sm"
                        onClick={() => handleRejectAutoEntry(index)}
                      >
                        <Trash2 className="h-4 w-4" />
                      </Button>
                    </div>
                  </div>
                </div>
              ))}
            </CardContent>
          </Card>
        )}

        {/* Current Income Entries */}
        {incomeEntries.length > 0 && (
          <Card>
            <CardHeader>
              <CardTitle>Your Income Sources</CardTitle>
              <CardDescription>
                Current total: ${totalIncome.toLocaleString()}
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              {incomeEntries.map((entry: any) => (
                <div key={entry.id} className="flex items-center justify-between p-4 bg-gray-50 rounded-lg">
                  <div className="flex-1">
                    <div className="flex items-center space-x-2">
                      <Badge variant="secondary">
                        {incomeTypes.find(t => t.value === entry.incomeType)?.label}
                      </Badge>
                      {entry.isAutoPopulated && (
                        <Badge variant="outline" className="text-xs text-blue-600">
                          <Sparkles className="mr-1 h-3 w-3" />
                          Auto-populated
                        </Badge>
                      )}
                      <span className="font-medium">${parseFloat(entry.amount).toLocaleString()}</span>
                    </div>
                    {entry.description && (
                      <p className="text-sm text-gray-600 mt-1">{entry.description}</p>
                    )}
                    {entry.employerName && (
                      <p className="text-sm text-gray-600 mt-1">Employer: {entry.employerName}</p>
                    )}
                    {entry.payerName && (
                      <p className="text-sm text-gray-600 mt-1">Payer: {entry.payerName}</p>
                    )}
                  </div>
                  <Button
                    type="button"
                    variant="ghost"
                    size="sm"
                    onClick={() => handleDeleteEntry(entry.id)}
                  >
                    <Trash2 className="h-4 w-4" />
                  </Button>
                </div>
              ))}
            </CardContent>
          </Card>
        )}

        {/* Add New Income Entry - Manual */}
        <Card>
          <CardHeader>
            <CardTitle className="flex items-center space-x-2">
              <Plus className="h-5 w-5" />
              <span>Add Income Source Manually</span>
            </CardTitle>
            <CardDescription>
              Manually add wages, interest, dividends, and other income
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
            <div className="grid grid-cols-2 gap-4">
              <div>
                <Label htmlFor="incomeType">Income Type *</Label>
                <Select value={newEntry.incomeType} onValueChange={(value) => setNewEntry({...newEntry, incomeType: value})}>
                  <SelectTrigger>
                    <SelectValue placeholder="Select income type" />
                  </SelectTrigger>
                  <SelectContent>
                    {incomeTypes.map((type) => (
                      <SelectItem key={type.value} value={type.value}>
                        {type.label}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              <div>
                <Label htmlFor="amount">Amount *</Label>
                <div className="relative">
                  <DollarSign className="absolute left-3 top-3 h-4 w-4 text-gray-400" />
                  <Input
                    id="amount"
                    type="number"
                    step="0.01"
                    value={newEntry.amount}
                    onChange={(e) => setNewEntry({...newEntry, amount: e.target.value})}
                    className="pl-10"
                    placeholder="0.00"
                  />
                </div>
              </div>
            </div>

            <div>
              <Label htmlFor="description">Description (Optional)</Label>
              <Input
                id="description"
                value={newEntry.description}
                onChange={(e) => setNewEntry({...newEntry, description: e.target.value})}
                placeholder="Additional details about this income"
              />
            </div>

            {newEntry.incomeType === "W2_WAGES" && (
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <Label htmlFor="employerName">Employer Name</Label>
                  <Input
                    id="employerName"
                    value={newEntry.employerName}
                    onChange={(e) => setNewEntry({...newEntry, employerName: e.target.value})}
                    placeholder="Your employer's name"
                  />
                </div>
                <div>
                  <Label htmlFor="employerEIN">Employer EIN</Label>
                  <Input
                    id="employerEIN"
                    value={newEntry.employerEIN}
                    onChange={(e) => setNewEntry({...newEntry, employerEIN: e.target.value})}
                    placeholder="00-0000000"
                  />
                </div>
              </div>
            )}

            {(newEntry.incomeType === "INTEREST" || newEntry.incomeType === "DIVIDENDS") && (
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <Label htmlFor="payerName">Payer Name</Label>
                  <Input
                    id="payerName"
                    value={newEntry.payerName}
                    onChange={(e) => setNewEntry({...newEntry, payerName: e.target.value})}
                    placeholder="Bank or institution name"
                  />
                </div>
                <div>
                  <Label htmlFor="payerTIN">Payer TIN</Label>
                  <Input
                    id="payerTIN"
                    value={newEntry.payerTIN}
                    onChange={(e) => setNewEntry({...newEntry, payerTIN: e.target.value})}
                    placeholder="00-0000000"
                  />
                </div>
              </div>
            )}

            <Button
              type="button"
              onClick={handleAddEntry}
              disabled={!newEntry.incomeType || !newEntry.amount}
              className="w-full"
            >
              <Plus className="mr-2 h-4 w-4" />
              Add Income Source
            </Button>
          </CardContent>
        </Card>

        {/* Summary Card */}
        <Card className="bg-green-50 border-green-200">
          <CardHeader>
            <CardTitle className="flex items-center space-x-2">
              <DollarSign className="h-5 w-5 text-green-600" />
              <span>Income Summary</span>
            </CardTitle>
          </CardHeader>
          <CardContent>
            <div className="text-2xl font-bold text-green-700">
              Total Income: ${totalIncome.toLocaleString()}
            </div>
            <p className="text-sm text-green-600 mt-1">
              This will be used as your Adjusted Gross Income for Stage 1 calculations
            </p>
          </CardContent>
        </Card>
      </div>

      {/* Navigation */}
      <div className="flex justify-between pt-6">
        <Button type="button" variant="outline" onClick={onPrev}>
          <ArrowLeft className="mr-2 h-4 w-4" />
          Previous
        </Button>
        <div className="flex space-x-2">
          <Button
            type="button"
            variant="outline"
            onClick={handleSaveOnly}
            disabled={saving || autoSaving}
          >
            {autoSaving ? "Auto-saving..." : "Save"}
          </Button>
          <Button
            type="submit"
            disabled={loading || saving}
            onClick={handleSaveAndContinue}
          >
            {loading || saving ? "Saving..." : "Save & Continue"}
            <ArrowRight className="ml-2 h-4 w-4" />
          </Button>
        </div>
      </div>

      {/* Name Validation Dialog */}
      <NameValidationDialog
        isOpen={nameValidationDialog.isOpen}
        onClose={() => setNameValidationDialog(prev => ({ ...prev, isOpen: false }))}
        onConfirm={handleNameValidationConfirm}
        validationResult={nameValidationDialog.validationResult}
        documentType={nameValidationDialog.documentType}
      />
    </form>
  )
}
