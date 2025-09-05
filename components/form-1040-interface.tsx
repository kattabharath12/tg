
"use client";

import { useState, useEffect } from "react";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Alert, AlertDescription } from "@/components/ui/alert";
import { Badge } from "@/components/ui/badge";
import { Separator } from "@/components/ui/separator";
import { Textarea } from "@/components/ui/textarea";
import { Checkbox } from "@/components/ui/checkbox";
import { Form1040Data, FilingStatus } from "@/lib/form-1040-types";
import { FileText, DollarSign, User, Calculator, Download, Save, RefreshCw, AlertCircle, CheckCircle } from "lucide-react";
import { toast } from "sonner";

interface Form1040InterfaceProps {
  initialData?: Partial<Form1040Data>;
  w2MappingData?: any[];
  onSave: (data: Form1040Data) => Promise<void>;
  onGeneratePDF: (data: Form1040Data) => Promise<void>;
  taxReturnId: string;
  readonly?: boolean;
}

export function Form1040Interface({
  initialData = {},
  w2MappingData = [],
  onSave,
  onGeneratePDF,
  taxReturnId,
  readonly = false
}: Form1040InterfaceProps) {
  const [formData, setFormData] = useState<Form1040Data>({
    taxYear: 2023,
    firstName: "",
    lastName: "",
    ssn: "",
    address: "",
    city: "",
    state: "",
    zipCode: "",
    filingStatus: FilingStatus.SINGLE,
    dependents: [],
    
    // Income lines
    line1: 0,
    line2a: 0,
    line2b: 0,
    line3a: 0,
    line3b: 0,
    line4a: 0,
    line4b: 0,
    line5a: 0,
    line5b: 0,
    line6a: 0,
    line6b: 0,
    line7: 0,
    line8: 0,
    line9: 0,
    
    // AGI lines
    line10: 0,
    line11: 0,
    
    // Deduction lines
    line12: 0,
    line13: 0,
    line14: 0,
    line15: 0,
    
    // Tax lines
    line16: 0,
    line17: 0,
    line18: 0,
    line19: 0,
    line20: 0,
    line21: 0,
    line22: 0,
    line23: 0,
    line24: 0,
    
    // Payment lines
    line25a: 0,
    line25b: 0,
    line25c: 0,
    line25d: 0,
    
    // Refund/Owed lines
    line32: 0,
    line33: 0,
    line34: 0,
    line35a: 0,
    line35b: "Checking",
    line35c: 0,
    line36: 0,
    line37: 0,
    
    // Other fields
    thirdPartyDesignee: false,
    
    ...initialData
  });

  const [isCalculating, setIsCalculating] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [isGeneratingPDF, setIsGeneratingPDF] = useState(false);
  const [hasChanges, setHasChanges] = useState(false);

  // Auto-calculate dependent fields when key fields change
  useEffect(() => {
    calculateDependentFields();
  }, [
    formData.line1, formData.line2b, formData.line3b, formData.line4b,
    formData.line5b, formData.line6b, formData.line7, formData.line8,
    formData.line10, formData.line12, formData.line13, formData.line25a,
    formData.line25b, formData.line25c, formData.line25d, formData.filingStatus
  ]);

  const calculateDependentFields = () => {
    setIsCalculating(true);
    
    // Calculate total income (line 9)
    const totalIncome = (formData.line1 || 0) + (formData.line2b || 0) + 
                       (formData.line3b || 0) + (formData.line4b || 0) + 
                       (formData.line5b || 0) + (formData.line6b || 0) + 
                       (formData.line7 || 0) + (formData.line8 || 0);
    
    // Calculate AGI (line 11)
    const agi = totalIncome - (formData.line10 || 0);
    
    // Set standard deduction if not set (line 12)
    let standardDeduction = formData.line12;
    if (!standardDeduction || standardDeduction === 0) {
      standardDeduction = getStandardDeduction(formData.filingStatus);
    }
    
    // Calculate line 14
    const line14 = standardDeduction + (formData.line13 || 0);
    
    // Calculate taxable income (line 15)
    const taxableIncome = Math.max(0, agi - line14);
    
    // Calculate tax liability (simplified)
    const taxLiability = calculateTaxLiability(taxableIncome, formData.filingStatus);
    
    // Calculate total tax (line 24)
    const totalTax = taxLiability + (formData.line17 || 0) + (formData.line23 || 0);
    
    // Calculate total payments (line 32)
    const totalPayments = (formData.line25a || 0) + (formData.line25b || 0) + 
                         (formData.line25c || 0) + (formData.line25d || 0);
    
    // Calculate refund or amount owed
    const overpaid = Math.max(0, totalPayments - totalTax);
    const amountOwed = Math.max(0, totalTax - totalPayments);
    
    setFormData(prev => ({
      ...prev,
      line9: totalIncome,
      line11: agi,
      line12: standardDeduction,
      line14: line14,
      line15: taxableIncome,
      line16: taxLiability,
      line18: taxLiability + (prev.line17 || 0),
      line22: Math.max(0, (taxLiability + (prev.line17 || 0)) - ((prev.line19 || 0) + (prev.line20 || 0))),
      line24: totalTax,
      line32: totalPayments,
      line33: overpaid,
      line34: overpaid, // Default to full refund
      line37: amountOwed
    }));
    
    setIsCalculating(false);
    setHasChanges(true);
  };

  const getStandardDeduction = (filingStatus: FilingStatus): number => {
    const deductions = {
      [FilingStatus.SINGLE]: 13850,
      [FilingStatus.MARRIED_FILING_JOINTLY]: 27700,
      [FilingStatus.MARRIED_FILING_SEPARATELY]: 13850,
      [FilingStatus.HEAD_OF_HOUSEHOLD]: 20800,
      [FilingStatus.QUALIFYING_SURVIVING_SPOUSE]: 27700
    };
    return deductions[filingStatus] || deductions[FilingStatus.SINGLE];
  };

  const calculateTaxLiability = (taxableIncome: number, filingStatus: FilingStatus): number => {
    // Simplified tax calculation using 2023 brackets
    const brackets = filingStatus === FilingStatus.MARRIED_FILING_JOINTLY ? [
      { min: 0, max: 22000, rate: 0.10 },
      { min: 22000, max: 89450, rate: 0.12 },
      { min: 89450, max: 190750, rate: 0.22 },
      { min: 190750, max: 364200, rate: 0.24 },
      { min: 364200, max: 462500, rate: 0.32 },
      { min: 462500, max: 693750, rate: 0.35 },
      { min: 693750, max: Infinity, rate: 0.37 }
    ] : [
      { min: 0, max: 11000, rate: 0.10 },
      { min: 11000, max: 44725, rate: 0.12 },
      { min: 44725, max: 95375, rate: 0.22 },
      { min: 95375, max: 182050, rate: 0.24 },
      { min: 182050, max: 231250, rate: 0.32 },
      { min: 231250, max: 578125, rate: 0.35 },
      { min: 578125, max: Infinity, rate: 0.37 }
    ];

    let tax = 0;
    let remainingIncome = taxableIncome;

    for (const bracket of brackets) {
      if (remainingIncome <= 0) break;
      const taxableAtThisBracket = Math.min(remainingIncome, bracket.max - bracket.min);
      tax += taxableAtThisBracket * bracket.rate;
      remainingIncome -= taxableAtThisBracket;
    }

    return Math.round(tax * 100) / 100;
  };

  const handleFieldChange = (field: keyof Form1040Data, value: any) => {
    setFormData(prev => ({ ...prev, [field]: value }));
    setHasChanges(true);
  };

  const handleSave = async () => {
    setIsSaving(true);
    try {
      await onSave(formData);
      setHasChanges(false);
      toast.success("Form 1040 saved successfully");
    } catch (error) {
      toast.error("Failed to save form");
      console.error(error);
    } finally {
      setIsSaving(false);
    }
  };

  const handleGeneratePDF = async () => {
    setIsGeneratingPDF(true);
    try {
      await onGeneratePDF(formData);
      toast.success("PDF generated successfully");
    } catch (error) {
      toast.error("Failed to generate PDF");
      console.error(error);
    } finally {
      setIsGeneratingPDF(false);
    }
  };

  const formatCurrency = (amount: number): string => {
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD',
      minimumFractionDigits: 0,
      maximumFractionDigits: 0,
    }).format(amount);
  };

  return (
    <div className="space-y-6">
      {/* Header with action buttons */}
      <div className="flex justify-between items-center">
        <div>
          <h2 className="text-2xl font-bold">Form 1040 - Individual Income Tax Return</h2>
          <p className="text-gray-600">Tax Year {formData.taxYear}</p>
        </div>
        
        {!readonly && (
          <div className="flex space-x-3">
            <Button
              onClick={handleSave}
              disabled={!hasChanges || isSaving}
              variant="outline"
            >
              <Save className="w-4 h-4 mr-2" />
              {isSaving ? "Saving..." : "Save"}
            </Button>
            
            <Button
              onClick={handleGeneratePDF}
              disabled={isGeneratingPDF}
            >
              <Download className="w-4 h-4 mr-2" />
              {isGeneratingPDF ? "Generating..." : "Download PDF"}
            </Button>
          </div>
        )}
      </div>

      {/* W2 Mapping Summary */}
      {w2MappingData && w2MappingData.length > 0 && (
        <Alert>
          <CheckCircle className="h-4 w-4" />
          <AlertDescription>
            Data from {w2MappingData.length} W-2 form(s) has been automatically populated in this form. 
            Please review and verify all amounts before submitting.
            <div className="mt-2 text-sm">
              <strong>Populated fields:</strong>
              {formData.line1 > 0 && <div>• Line 1 (Wages): ${formData.line1.toLocaleString()}</div>}
              {formData.line25a > 0 && <div>• Line 25a (Federal Tax Withheld): ${formData.line25a.toLocaleString()}</div>}
            </div>
          </AlertDescription>
        </Alert>
      )}

      {/* Debug Info for W2 Mapping (only in development) */}
      {process.env.NODE_ENV === 'development' && w2MappingData && w2MappingData.length > 0 && (
        <Alert className="border-blue-200 bg-blue-50">
          <AlertCircle className="h-4 w-4 text-blue-600" />
          <AlertDescription className="text-blue-800">
            <strong>Debug: W2 Mapping Data</strong>
            <details className="mt-2">
              <summary className="cursor-pointer">View W2 mapping details</summary>
              <pre className="text-xs mt-2 overflow-x-auto">
                {JSON.stringify(w2MappingData, null, 2)}
              </pre>
            </details>
          </AlertDescription>
        </Alert>
      )}

      {/* Calculation indicator */}
      {isCalculating && (
        <Alert>
          <RefreshCw className="h-4 w-4 animate-spin" />
          <AlertDescription>Recalculating tax amounts...</AlertDescription>
        </Alert>
      )}

      <Tabs defaultValue="personal" className="space-y-6">
        <TabsList className="grid w-full grid-cols-6">
          <TabsTrigger value="personal">Personal Info</TabsTrigger>
          <TabsTrigger value="income">Income</TabsTrigger>
          <TabsTrigger value="deductions">Deductions</TabsTrigger>
          <TabsTrigger value="tax">Tax & Credits</TabsTrigger>
          <TabsTrigger value="payments">Payments</TabsTrigger>
          <TabsTrigger value="refund">Refund/Owed</TabsTrigger>
        </TabsList>

        {/* Personal Information Tab */}
        <TabsContent value="personal">
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center">
                <User className="w-5 h-5 mr-2" />
                Personal Information
              </CardTitle>
              <CardDescription>
                Your personal details and filing status
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <Label htmlFor="firstName">First Name</Label>
                  <Input
                    id="firstName"
                    value={formData.firstName}
                    onChange={(e) => handleFieldChange('firstName', e.target.value)}
                    disabled={readonly}
                  />
                </div>
                <div>
                  <Label htmlFor="lastName">Last Name</Label>
                  <Input
                    id="lastName"
                    value={formData.lastName}
                    onChange={(e) => handleFieldChange('lastName', e.target.value)}
                    disabled={readonly}
                  />
                </div>
              </div>
              
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <Label htmlFor="ssn">Social Security Number</Label>
                  <Input
                    id="ssn"
                    value={formData.ssn}
                    onChange={(e) => handleFieldChange('ssn', e.target.value)}
                    placeholder="XXX-XX-XXXX"
                    disabled={readonly}
                  />
                </div>
                <div>
                  <Label htmlFor="filingStatus">Filing Status</Label>
                  <Select 
                    value={formData.filingStatus} 
                    onValueChange={(value) => handleFieldChange('filingStatus', value)}
                    disabled={readonly}
                  >
                    <SelectTrigger>
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value={FilingStatus.SINGLE}>Single</SelectItem>
                      <SelectItem value={FilingStatus.MARRIED_FILING_JOINTLY}>Married Filing Jointly</SelectItem>
                      <SelectItem value={FilingStatus.MARRIED_FILING_SEPARATELY}>Married Filing Separately</SelectItem>
                      <SelectItem value={FilingStatus.HEAD_OF_HOUSEHOLD}>Head of Household</SelectItem>
                      <SelectItem value={FilingStatus.QUALIFYING_SURVIVING_SPOUSE}>Qualifying Surviving Spouse</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
              </div>

              <div>
                <Label htmlFor="address">Street Address</Label>
                <Input
                  id="address"
                  value={formData.address}
                  onChange={(e) => handleFieldChange('address', e.target.value)}
                  disabled={readonly}
                />
              </div>
              
              <div className="grid grid-cols-3 gap-4">
                <div>
                  <Label htmlFor="city">City</Label>
                  <Input
                    id="city"
                    value={formData.city}
                    onChange={(e) => handleFieldChange('city', e.target.value)}
                    disabled={readonly}
                  />
                </div>
                <div>
                  <Label htmlFor="state">State</Label>
                  <Input
                    id="state"
                    value={formData.state}
                    onChange={(e) => handleFieldChange('state', e.target.value)}
                    maxLength={2}
                    disabled={readonly}
                  />
                </div>
                <div>
                  <Label htmlFor="zipCode">ZIP Code</Label>
                  <Input
                    id="zipCode"
                    value={formData.zipCode}
                    onChange={(e) => handleFieldChange('zipCode', e.target.value)}
                    disabled={readonly}
                  />
                </div>
              </div>
            </CardContent>
          </Card>
        </TabsContent>

        {/* Income Tab */}
        <TabsContent value="income">
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center">
                <DollarSign className="w-5 h-5 mr-2" />
                Income
              </CardTitle>
              <CardDescription>
                Report all sources of income for the tax year
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <Label htmlFor="line1">Line 1: Wages from Form W-2</Label>
                  <Input
                    id="line1"
                    type="number"
                    step="0.01"
                    value={formData.line1}
                    onChange={(e) => handleFieldChange('line1', parseFloat(e.target.value) || 0)}
                    disabled={readonly}
                  />
                  {w2MappingData.length > 0 && (
                    <Badge variant="secondary" className="mt-1">
                      Auto-populated from W-2
                    </Badge>
                  )}
                </div>
                <div>
                  <Label htmlFor="line2b">Line 2b: Taxable Interest</Label>
                  <Input
                    id="line2b"
                    type="number"
                    step="0.01"
                    value={formData.line2b}
                    onChange={(e) => handleFieldChange('line2b', parseFloat(e.target.value) || 0)}
                    disabled={readonly}
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <Label htmlFor="line3b">Line 3b: Ordinary Dividends</Label>
                  <Input
                    id="line3b"
                    type="number"
                    step="0.01"
                    value={formData.line3b}
                    onChange={(e) => handleFieldChange('line3b', parseFloat(e.target.value) || 0)}
                    disabled={readonly}
                  />
                </div>
                <div>
                  <Label htmlFor="line7">Line 7: Capital Gain or (Loss)</Label>
                  <Input
                    id="line7"
                    type="number"
                    step="0.01"
                    value={formData.line7}
                    onChange={(e) => handleFieldChange('line7', parseFloat(e.target.value) || 0)}
                    disabled={readonly}
                  />
                </div>
              </div>

              <Separator />
              
              <div className="bg-gray-50 p-4 rounded-lg">
                <div className="flex justify-between items-center">
                  <Label className="text-lg font-semibold">Line 9: Total Income</Label>
                  <span className="text-2xl font-bold text-green-600">
                    {formatCurrency(formData.line9)}
                  </span>
                </div>
                <p className="text-sm text-gray-600 mt-1">
                  Automatically calculated from lines 1, 2b, 3b, 4b, 5b, 6b, 7, and 8
                </p>
              </div>
            </CardContent>
          </Card>
        </TabsContent>

        {/* Deductions Tab */}
        <TabsContent value="deductions">
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center">
                <Calculator className="w-5 h-5 mr-2" />
                Deductions
              </CardTitle>
              <CardDescription>
                Standard deduction or itemized deductions
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <Label htmlFor="line10">Line 10: Adjustments to Income</Label>
                  <Input
                    id="line10"
                    type="number"
                    step="0.01"
                    value={formData.line10}
                    onChange={(e) => handleFieldChange('line10', parseFloat(e.target.value) || 0)}
                    disabled={readonly}
                  />
                </div>
                <div className="bg-gray-50 p-4 rounded-lg">
                  <Label className="font-semibold">Line 11: Adjusted Gross Income</Label>
                  <div className="text-xl font-bold text-blue-600">
                    {formatCurrency(formData.line11)}
                  </div>
                  <p className="text-sm text-gray-600">Line 9 minus Line 10</p>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <Label htmlFor="line12">Line 12: Standard Deduction</Label>
                  <Input
                    id="line12"
                    type="number"
                    step="0.01"
                    value={formData.line12}
                    onChange={(e) => handleFieldChange('line12', parseFloat(e.target.value) || 0)}
                    disabled={readonly}
                  />
                  <Badge variant="secondary" className="mt-1">
                    Standard: {formatCurrency(getStandardDeduction(formData.filingStatus))}
                  </Badge>
                </div>
                <div>
                  <Label htmlFor="line13">Line 13: QBI Deduction</Label>
                  <Input
                    id="line13"
                    type="number"
                    step="0.01"
                    value={formData.line13}
                    onChange={(e) => handleFieldChange('line13', parseFloat(e.target.value) || 0)}
                    disabled={readonly}
                  />
                </div>
              </div>

              <Separator />
              
              <div className="bg-gray-50 p-4 rounded-lg">
                <div className="flex justify-between items-center">
                  <Label className="text-lg font-semibold">Line 15: Taxable Income</Label>
                  <span className="text-2xl font-bold text-orange-600">
                    {formatCurrency(formData.line15)}
                  </span>
                </div>
                <p className="text-sm text-gray-600 mt-1">
                  AGI minus deductions (Line 11 - Line 14)
                </p>
              </div>
            </CardContent>
          </Card>
        </TabsContent>

        {/* Tax & Credits Tab */}
        <TabsContent value="tax">
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center">
                <Calculator className="w-5 h-5 mr-2" />
                Tax and Credits
              </CardTitle>
              <CardDescription>
                Tax liability and applicable credits
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="bg-gray-50 p-4 rounded-lg">
                <div className="flex justify-between items-center">
                  <Label className="text-lg font-semibold">Line 16: Tax</Label>
                  <span className="text-2xl font-bold text-red-600">
                    {formatCurrency(formData.line16)}
                  </span>
                </div>
                <p className="text-sm text-gray-600 mt-1">
                  Calculated based on taxable income and filing status
                </p>
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <Label htmlFor="line17">Line 17: Additional Tax</Label>
                  <Input
                    id="line17"
                    type="number"
                    step="0.01"
                    value={formData.line17}
                    onChange={(e) => handleFieldChange('line17', parseFloat(e.target.value) || 0)}
                    disabled={readonly}
                  />
                </div>
                <div>
                  <Label htmlFor="line19">Line 19: Child Tax Credit</Label>
                  <Input
                    id="line19"
                    type="number"
                    step="0.01"
                    value={formData.line19}
                    onChange={(e) => handleFieldChange('line19', parseFloat(e.target.value) || 0)}
                    disabled={readonly}
                  />
                </div>
              </div>

              <Separator />
              
              <div className="bg-gray-50 p-4 rounded-lg">
                <div className="flex justify-between items-center">
                  <Label className="text-lg font-semibold">Line 24: Total Tax</Label>
                  <span className="text-2xl font-bold text-red-600">
                    {formatCurrency(formData.line24)}
                  </span>
                </div>
                <p className="text-sm text-gray-600 mt-1">
                  Tax liability after credits
                </p>
              </div>
            </CardContent>
          </Card>
        </TabsContent>

        {/* Payments Tab */}
        <TabsContent value="payments">
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center">
                <DollarSign className="w-5 h-5 mr-2" />
                Payments and Withholdings
              </CardTitle>
              <CardDescription>
                Tax payments and withholdings made during the year
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <Label htmlFor="line25a">Line 25a: Federal Tax Withheld</Label>
                  <Input
                    id="line25a"
                    type="number"
                    step="0.01"
                    value={formData.line25a}
                    onChange={(e) => handleFieldChange('line25a', parseFloat(e.target.value) || 0)}
                    disabled={readonly}
                  />
                  {w2MappingData.length > 0 && (
                    <Badge variant="secondary" className="mt-1">
                      Auto-populated from W-2
                    </Badge>
                  )}
                </div>
                <div>
                  <Label htmlFor="line25b">Line 25b: Estimated Tax Payments</Label>
                  <Input
                    id="line25b"
                    type="number"
                    step="0.01"
                    value={formData.line25b}
                    onChange={(e) => handleFieldChange('line25b', parseFloat(e.target.value) || 0)}
                    disabled={readonly}
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <Label htmlFor="line25c">Line 25c: Earned Income Credit</Label>
                  <Input
                    id="line25c"
                    type="number"
                    step="0.01"
                    value={formData.line25c}
                    onChange={(e) => handleFieldChange('line25c', parseFloat(e.target.value) || 0)}
                    disabled={readonly}
                  />
                </div>
                <div>
                  <Label htmlFor="line25d">Line 25d: Additional Child Tax Credit</Label>
                  <Input
                    id="line25d"
                    type="number"
                    step="0.01"
                    value={formData.line25d}
                    onChange={(e) => handleFieldChange('line25d', parseFloat(e.target.value) || 0)}
                    disabled={readonly}
                  />
                </div>
              </div>

              <Separator />
              
              <div className="bg-gray-50 p-4 rounded-lg">
                <div className="flex justify-between items-center">
                  <Label className="text-lg font-semibold">Line 32: Total Payments</Label>
                  <span className="text-2xl font-bold text-green-600">
                    {formatCurrency(formData.line32)}
                  </span>
                </div>
                <p className="text-sm text-gray-600 mt-1">
                  Sum of all payments and refundable credits
                </p>
              </div>
            </CardContent>
          </Card>
        </TabsContent>

        {/* Refund/Owed Tab */}
        <TabsContent value="refund">
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center">
                <FileText className="w-5 h-5 mr-2" />
                Refund or Amount Owed
              </CardTitle>
              <CardDescription>
                Final calculation of refund or amount owed
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-6">
              {/* Summary Section */}
              <div className="grid grid-cols-2 gap-6">
                <div className="bg-red-50 p-4 rounded-lg border border-red-200">
                  <Label className="text-lg font-semibold text-red-800">Total Tax</Label>
                  <div className="text-3xl font-bold text-red-600">
                    {formatCurrency(formData.line24)}
                  </div>
                  <p className="text-sm text-red-600">Line 24</p>
                </div>
                
                <div className="bg-green-50 p-4 rounded-lg border border-green-200">
                  <Label className="text-lg font-semibold text-green-800">Total Payments</Label>
                  <div className="text-3xl font-bold text-green-600">
                    {formatCurrency(formData.line32)}
                  </div>
                  <p className="text-sm text-green-600">Line 32</p>
                </div>
              </div>

              {/* Result Section */}
              <div className="text-center py-8">
                {formData.line33 > 0 ? (
                  <div className="bg-green-50 p-6 rounded-lg border border-green-200">
                    <h3 className="text-2xl font-bold text-green-800 mb-2">Refund Expected</h3>
                    <div className="text-5xl font-bold text-green-600 mb-4">
                      {formatCurrency(formData.line33)}
                    </div>
                    <p className="text-green-700">
                      You overpaid your taxes and are entitled to a refund
                    </p>
                  </div>
                ) : formData.line37 > 0 ? (
                  <div className="bg-red-50 p-6 rounded-lg border border-red-200">
                    <h3 className="text-2xl font-bold text-red-800 mb-2">Amount Owed</h3>
                    <div className="text-5xl font-bold text-red-600 mb-4">
                      {formatCurrency(formData.line37)}
                    </div>
                    <p className="text-red-700">
                      Additional tax is owed. Payment due by April 15, 2024
                    </p>
                  </div>
                ) : (
                  <div className="bg-gray-50 p-6 rounded-lg border border-gray-200">
                    <h3 className="text-2xl font-bold text-gray-800 mb-2">Tax Balance</h3>
                    <div className="text-5xl font-bold text-gray-600 mb-4">
                      $0
                    </div>
                    <p className="text-gray-700">
                      Your tax payments exactly match your tax liability
                    </p>
                  </div>
                )}
              </div>

              {/* Refund Details */}
              {formData.line33 > 0 && (
                <div className="space-y-4">
                  <h4 className="text-lg font-semibold">Refund Information</h4>
                  
                  <div>
                    <Label htmlFor="line34">Line 34: Amount to be Refunded</Label>
                    <Input
                      id="line34"
                      type="number"
                      step="0.01"
                      value={formData.line34}
                      onChange={(e) => handleFieldChange('line34', parseFloat(e.target.value) || 0)}
                      max={formData.line33}
                      disabled={readonly}
                    />
                    <p className="text-sm text-gray-600 mt-1">
                      Maximum: {formatCurrency(formData.line33)}
                    </p>
                  </div>

                  <div className="grid grid-cols-3 gap-4">
                    <div>
                      <Label htmlFor="line35a">Routing Number</Label>
                      <Input
                        id="line35a"
                        type="number"
                        value={formData.line35a}
                        onChange={(e) => handleFieldChange('line35a', parseInt(e.target.value) || 0)}
                        disabled={readonly}
                      />
                    </div>
                    <div>
                      <Label htmlFor="line35b">Account Type</Label>
                      <Select 
                        value={formData.line35b} 
                        onValueChange={(value) => handleFieldChange('line35b', value)}
                        disabled={readonly}
                      >
                        <SelectTrigger>
                          <SelectValue />
                        </SelectTrigger>
                        <SelectContent>
                          <SelectItem value="Checking">Checking</SelectItem>
                          <SelectItem value="Savings">Savings</SelectItem>
                        </SelectContent>
                      </Select>
                    </div>
                    <div>
                      <Label htmlFor="line35c">Account Number</Label>
                      <Input
                        id="line35c"
                        type="number"
                        value={formData.line35c}
                        onChange={(e) => handleFieldChange('line35c', parseInt(e.target.value) || 0)}
                        disabled={readonly}
                      />
                    </div>
                  </div>
                </div>
              )}
            </CardContent>
          </Card>
        </TabsContent>
      </Tabs>
    </div>
  );
}
