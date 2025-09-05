
"use client"

import { useState } from "react"
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card"
import { Button } from "@/components/ui/button"
import { Badge } from "@/components/ui/badge"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select"
import { Separator } from "@/components/ui/separator"
import { Check, X, Edit, DollarSign, FileText, AlertCircle } from "lucide-react"

interface ExtractedEntry {
  id: string
  type: 'income' | 'deduction'
  documentType: string
  extractedData: any
  confidence: number
  isVerified: boolean
  isEdited: boolean
}

interface ExtractedDataVerificationProps {
  entries: ExtractedEntry[]
  onAccept: (entryId: string, data: any) => void
  onReject: (entryId: string) => void
  onEdit: (entryId: string, data: any) => void
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

export function ExtractedDataVerification({ 
  entries, 
  onAccept, 
  onReject, 
  onEdit 
}: ExtractedDataVerificationProps) {
  const [editingEntry, setEditingEntry] = useState<string | null>(null)
  const [editedData, setEditedData] = useState<any>({})

  const handleStartEdit = (entry: ExtractedEntry) => {
    setEditingEntry(entry.id)
    setEditedData(entry.extractedData)
  }

  const handleSaveEdit = (entryId: string) => {
    onEdit(entryId, editedData)
    setEditingEntry(null)
    setEditedData({})
  }

  const handleCancelEdit = () => {
    setEditingEntry(null)
    setEditedData({})
  }

  const getConfidenceColor = (confidence: number) => {
    if (confidence >= 0.8) return "text-green-600"
    if (confidence >= 0.6) return "text-yellow-600"
    return "text-red-600"
  }

  const getConfidenceLabel = (confidence: number) => {
    if (confidence >= 0.8) return "High"
    if (confidence >= 0.6) return "Medium"
    return "Low"
  }

  const formatDocumentType = (documentType: string) => {
    const types: Record<string, string> = {
      'W2': 'W-2 Form',
      'FORM_1099_INT': '1099-INT',
      'FORM_1099_DIV': '1099-DIV',
      'FORM_1099_MISC': '1099-MISC',
      'FORM_1099_NEC': '1099-NEC',
      'OTHER_TAX_DOCUMENT': 'Other Tax Document'
    }
    return types[documentType] || documentType
  }

  if (entries.length === 0) {
    return (
      <Card>
        <CardContent className="py-8 text-center">
          <FileText className="h-12 w-12 text-gray-400 mx-auto mb-4" />
          <p className="text-gray-600">No extracted data to verify</p>
        </CardContent>
      </Card>
    )
  }

  return (
    <div className="space-y-4">
      <Card>
        <CardHeader>
          <CardTitle>Extracted Data Verification</CardTitle>
          <CardDescription>
            Review and verify the information extracted from your documents
          </CardDescription>
        </CardHeader>
        <CardContent>
          <div className="space-y-6">
            {entries.map((entry) => (
              <div key={entry.id} className="border rounded-lg p-4">
                <div className="flex items-center justify-between mb-4">
                  <div className="flex items-center space-x-2">
                    <Badge variant="secondary">
                      {formatDocumentType(entry.documentType)}
                    </Badge>
                    <Badge 
                      variant="outline" 
                      className={getConfidenceColor(entry.confidence)}
                    >
                      {getConfidenceLabel(entry.confidence)} Confidence
                    </Badge>
                    {entry.isVerified && (
                      <Badge variant="default" className="text-green-600">
                        <Check className="h-3 w-3 mr-1" />
                        Verified
                      </Badge>
                    )}
                    {entry.isEdited && (
                      <Badge variant="outline" className="text-blue-600">
                        <Edit className="h-3 w-3 mr-1" />
                        Edited
                      </Badge>
                    )}
                  </div>
                  
                  {!entry.isVerified && (
                    <div className="flex space-x-2">
                      <Button
                        variant="outline"
                        size="sm"
                        onClick={() => handleStartEdit(entry)}
                        disabled={editingEntry === entry.id}
                      >
                        <Edit className="h-4 w-4 mr-1" />
                        Edit
                      </Button>
                      <Button
                        variant="outline"
                        size="sm"
                        onClick={() => onAccept(entry.id, entry.extractedData)}
                      >
                        <Check className="h-4 w-4 mr-1" />
                        Accept
                      </Button>
                      <Button
                        variant="outline"
                        size="sm"
                        onClick={() => onReject(entry.id)}
                      >
                        <X className="h-4 w-4 mr-1" />
                        Reject
                      </Button>
                    </div>
                  )}
                </div>

                {editingEntry === entry.id ? (
                  <EditingForm
                    data={editedData}
                    onChange={setEditedData}
                    onSave={() => handleSaveEdit(entry.id)}
                    onCancel={handleCancelEdit}
                    type={entry.type}
                  />
                ) : (
                  <DataPreview data={entry.extractedData} type={entry.type} documentType={entry.documentType} />
                )}
              </div>
            ))}
          </div>
        </CardContent>
      </Card>
    </div>
  )
}

interface EditingFormProps {
  data: any
  onChange: (data: any) => void
  onSave: () => void
  onCancel: () => void
  type: 'income' | 'deduction'
}

function EditingForm({ data, onChange, onSave, onCancel, type }: EditingFormProps) {
  const updateData = (field: string, value: any) => {
    onChange({ ...data, [field]: value })
  }

  // Helper function to render currency input
  const renderCurrencyInput = (label: string, field: string) => (
    <div>
      <Label>{label}</Label>
      <div className="relative">
        <DollarSign className="absolute left-3 top-3 h-4 w-4 text-gray-400" />
        <Input
          type="number"
          step="0.01"
          value={data[field] || ''}
          onChange={(e) => updateData(field, parseFloat(e.target.value) || 0)}
          className="pl-10"
        />
      </div>
    </div>
  );

  // Helper function to render text input
  const renderTextInput = (label: string, field: string) => (
    <div>
      <Label>{label}</Label>
      <Input
        value={data[field] || ''}
        onChange={(e) => updateData(field, e.target.value)}
      />
    </div>
  );

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 gap-4">
        {type === 'income' && (
          <div>
            <Label>Income Type</Label>
            <Select 
              value={data.incomeType || ''} 
              onValueChange={(value) => updateData('incomeType', value)}
            >
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
        )}
        
        <div>
          <Label>Amount</Label>
          <div className="relative">
            <DollarSign className="absolute left-3 top-3 h-4 w-4 text-gray-400" />
            <Input
              type="number"
              step="0.01"
              value={data.amount || ''}
              onChange={(e) => updateData('amount', parseFloat(e.target.value) || 0)}
              className="pl-10"
            />
          </div>
        </div>
      </div>

      <div>
        <Label>Description</Label>
        <Input
          value={data.description || ''}
          onChange={(e) => updateData('description', e.target.value)}
        />
      </div>

      {/* Employer Information */}
      {data.employerName !== undefined && (
        <div className="grid grid-cols-2 gap-4">
          {renderTextInput('Employer Name', 'employerName')}
          {renderTextInput('Employer EIN', 'employerEIN')}
        </div>
      )}

      {/* Payer Information */}
      {data.payerName !== undefined && (
        <div className="space-y-4">
          <div className="grid grid-cols-2 gap-4">
            {renderTextInput('Payer Name', 'payerName')}
            {renderTextInput('Payer TIN', 'payerTIN')}
          </div>
          {data.payerAddress !== undefined && (
            <div>
              {renderTextInput('Payer Address', 'payerAddress')}
            </div>
          )}
        </div>
      )}

      {/* Recipient Information */}
      {data.recipientName !== undefined && (
        <div className="space-y-4">
          <div className="grid grid-cols-2 gap-4">
            {renderTextInput('Recipient Name', 'recipientName')}
            {renderTextInput('Recipient TIN', 'recipientTIN')}
          </div>
          {data.recipientAddress !== undefined && (
            <div>
              {renderTextInput('Recipient Address', 'recipientAddress')}
            </div>
          )}
          {data.accountNumber !== undefined && (
            <div className="grid grid-cols-2 gap-4">
              {renderTextInput('Account Number', 'accountNumber')}
            </div>
          )}
        </div>
      )}

      {/* 1099-INT Specific Fields */}
      {(data.interestIncome !== undefined || data.earlyWithdrawalPenalty !== undefined) && (
        <div className="space-y-4">
          <Separator />
          <h4 className="text-sm font-semibold text-gray-700">1099-INT Details</h4>
          
          <div className="grid grid-cols-2 gap-4">
            {data.interestIncome !== undefined && renderCurrencyInput('Box 1 - Interest Income', 'interestIncome')}
            {data.earlyWithdrawalPenalty !== undefined && renderCurrencyInput('Box 2 - Early Withdrawal Penalty', 'earlyWithdrawalPenalty')}
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            {data.interestOnUSavingsBonds !== undefined && renderCurrencyInput('Box 3 - Interest on US Savings Bonds', 'interestOnUSavingsBonds')}
            {data.federalTaxWithheld !== undefined && renderCurrencyInput('Box 4 - Federal Tax Withheld', 'federalTaxWithheld')}
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            {data.investmentExpenses !== undefined && renderCurrencyInput('Box 5 - Investment Expenses', 'investmentExpenses')}
            {data.foreignTaxPaid !== undefined && renderCurrencyInput('Box 6 - Foreign Tax Paid', 'foreignTaxPaid')}
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            {data.taxExemptInterest !== undefined && renderCurrencyInput('Box 8 - Tax-Exempt Interest', 'taxExemptInterest')}
            {data.specifiedPrivateActivityBondInterest !== undefined && renderCurrencyInput('Box 9 - Specified Private Activity Bond Interest', 'specifiedPrivateActivityBondInterest')}
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            {data.marketDiscount !== undefined && renderCurrencyInput('Box 10 - Market Discount', 'marketDiscount')}
            {data.bondPremium !== undefined && renderCurrencyInput('Box 11 - Bond Premium', 'bondPremium')}
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            {data.bondPremiumOnTreasuryObligations !== undefined && renderCurrencyInput('Box 12 - Bond Premium on Treasury Obligations', 'bondPremiumOnTreasuryObligations')}
            {data.bondPremiumOnTaxExemptBond !== undefined && renderCurrencyInput('Box 13 - Bond Premium on Tax-Exempt Bond', 'bondPremiumOnTaxExemptBond')}
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            {data.taxExemptAndTaxCreditBondCUSIPNo !== undefined && renderTextInput('Box 14 - Tax-Exempt and Tax Credit Bond CUSIP No.', 'taxExemptAndTaxCreditBondCUSIPNo')}
            {data.state !== undefined && renderTextInput('Box 15 - State', 'state')}
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            {data.stateIdentificationNo !== undefined && renderTextInput('Box 16 - State Identification No.', 'stateIdentificationNo')}
            {data.stateTaxWithheld !== undefined && renderCurrencyInput('Box 17 - State Tax Withheld', 'stateTaxWithheld')}
          </div>
        </div>
      )}

      <div className="flex justify-end space-x-2">
        <Button variant="outline" onClick={onCancel}>
          Cancel
        </Button>
        <Button onClick={onSave}>
          Save Changes
        </Button>
      </div>
    </div>
  )
}

interface DataPreviewProps {
  data: any
  type: 'income' | 'deduction'
  documentType?: string
}

function DataPreview({ data, type, documentType }: DataPreviewProps) {
  // Helper function to format currency
  const formatCurrency = (amount: any) => {
    if (!amount || amount === 0) return '$0.00';
    return `$${parseFloat(amount).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  };

  // Helper function to render field if it exists
  const renderField = (label: string, value: any, isCurrency = false) => {
    if (!value && value !== 0) return null;
    return (
      <div>
        <Label className="text-sm font-medium text-gray-500">{label}</Label>
        <p className="text-sm font-medium">
          {isCurrency ? formatCurrency(value) : value}
        </p>
      </div>
    );
  };

  return (
    <div className="space-y-4">
      {/* Basic Information */}
      <div className="grid grid-cols-2 gap-4">
        <div>
          <Label className="text-sm font-medium text-gray-500">
            {type === 'income' ? 'Income Type' : 'Deduction Type'}
          </Label>
          <p className="text-sm">
            {type === 'income' 
              ? incomeTypes.find(t => t.value === data.incomeType)?.label || data.incomeType
              : data.deductionType
            }
          </p>
        </div>
        <div>
          <Label className="text-sm font-medium text-gray-500">Amount</Label>
          <p className="text-sm font-medium">{formatCurrency(data.amount)}</p>
        </div>
      </div>

      {data.description && (
        <div>
          <Label className="text-sm font-medium text-gray-500">Description</Label>
          <p className="text-sm">{data.description}</p>
        </div>
      )}

      {/* Payer/Employer Information */}
      {(data.employerName || data.payerName) && (
        <div className="grid grid-cols-2 gap-4">
          {data.employerName && (
            <div>
              <Label className="text-sm font-medium text-gray-500">Employer</Label>
              <p className="text-sm">{data.employerName}</p>
              {data.employerEIN && (
                <p className="text-xs text-gray-400">EIN: {data.employerEIN}</p>
              )}
            </div>
          )}
          {data.payerName && (
            <div>
              <Label className="text-sm font-medium text-gray-500">Payer</Label>
              <p className="text-sm">{data.payerName}</p>
              {data.payerTIN && (
                <p className="text-xs text-gray-400">TIN: {data.payerTIN}</p>
              )}
            </div>
          )}
        </div>
      )}

      {/* Recipient Information */}
      {(data.recipientName || data.employeeName) && (
        <div className="grid grid-cols-2 gap-4">
          {(data.recipientName || data.employeeName) && (
            <div>
              <Label className="text-sm font-medium text-gray-500">
                {data.recipientName ? 'Recipient' : 'Employee'}
              </Label>
              <p className="text-sm">{data.recipientName || data.employeeName}</p>
              {(data.recipientTIN || data.employeeSSN) && (
                <p className="text-xs text-gray-400">
                  {data.recipientTIN ? `TIN: ${data.recipientTIN}` : `SSN: ${data.employeeSSN}`}
                </p>
              )}
            </div>
          )}
          {data.accountNumber && (
            <div>
              <Label className="text-sm font-medium text-gray-500">Account Number</Label>
              <p className="text-sm">{data.accountNumber}</p>
            </div>
          )}
        </div>
      )}

      {/* 1099-INT Specific Fields */}
      {documentType === 'FORM_1099_INT' && (
        <div className="space-y-4">
          <Separator />
          <h4 className="text-sm font-semibold text-gray-700">1099-INT Details</h4>
          
          {/* Primary Income Fields */}
          <div className="grid grid-cols-2 gap-4">
            {renderField('Box 1 - Interest Income', data.interestIncome, true)}
            {renderField('Box 2 - Early Withdrawal Penalty', data.earlyWithdrawalPenalty, true)}
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            {renderField('Box 3 - Interest on US Savings Bonds', data.interestOnUSavingsBonds, true)}
            {renderField('Box 4 - Federal Tax Withheld', data.federalTaxWithheld, true)}
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            {renderField('Box 5 - Investment Expenses', data.investmentExpenses, true)}
            {renderField('Box 6 - Foreign Tax Paid', data.foreignTaxPaid, true)}
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            {renderField('Box 8 - Tax-Exempt Interest', data.taxExemptInterest, true)}
            {renderField('Box 9 - Specified Private Activity Bond Interest', data.specifiedPrivateActivityBondInterest, true)}
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            {renderField('Box 10 - Market Discount', data.marketDiscount, true)}
            {renderField('Box 11 - Bond Premium', data.bondPremium, true)}
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            {renderField('Box 12 - Bond Premium on Treasury Obligations', data.bondPremiumOnTreasuryObligations, true)}
            {renderField('Box 13 - Bond Premium on Tax-Exempt Bond', data.bondPremiumOnTaxExemptBond, true)}
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            {renderField('Box 14 - Tax-Exempt and Tax Credit Bond CUSIP No.', data.taxExemptAndTaxCreditBondCUSIPNo)}
            {renderField('Box 15 - State', data.state)}
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            {renderField('Box 16 - State Identification No.', data.stateIdentificationNo)}
            {renderField('Box 17 - State Tax Withheld', data.stateTaxWithheld, true)}
          </div>
        </div>
      )}

      {/* 1099-DIV Specific Fields */}
      {documentType === 'FORM_1099_DIV' && (
        <div className="space-y-4">
          <Separator />
          <h4 className="text-sm font-semibold text-gray-700">1099-DIV Details</h4>
          
          <div className="grid grid-cols-2 gap-4">
            {renderField('Box 1a - Ordinary Dividends', data.ordinaryDividends, true)}
            {renderField('Box 1b - Qualified Dividends', data.qualifiedDividends, true)}
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            {renderField('Box 2a - Total Capital Gain Distributions', data.totalCapitalGain, true)}
            {renderField('Box 3 - Nondividend Distributions', data.nondividendDistributions, true)}
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            {renderField('Box 4 - Federal Income Tax Withheld', data.federalTaxWithheld, true)}
            {renderField('Box 5 - Section 199A Dividends', data.section199ADividends, true)}
          </div>
        </div>
      )}

      {/* 1099-MISC Specific Fields */}
      {documentType === 'FORM_1099_MISC' && (
        <div className="space-y-4">
          <Separator />
          <h4 className="text-sm font-semibold text-gray-700">1099-MISC Details</h4>
          
          <div className="grid grid-cols-2 gap-4">
            {renderField('Box 1 - Rents', data.rents, true)}
            {renderField('Box 2 - Royalties', data.royalties, true)}
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            {renderField('Box 3 - Other Income', data.otherIncome, true)}
            {renderField('Box 4 - Federal Income Tax Withheld', data.federalTaxWithheld, true)}
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            {renderField('Box 5 - Fishing Boat Proceeds', data.fishingBoatProceeds, true)}
            {renderField('Box 6 - Medical and Health Care Payments', data.medicalHealthPayments, true)}
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            {renderField('Box 7 - Nonemployee Compensation', data.nonemployeeCompensation, true)}
            {renderField('Box 8 - Substitute Payments', data.substitutePayments, true)}
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            {renderField('Box 9 - Crop Insurance Proceeds', data.cropInsuranceProceeds, true)}
            {renderField('Box 10 - Gross Proceeds Paid to an Attorney', data.grossProceedsAttorney, true)}
          </div>
        </div>
      )}
    </div>
  )
}
