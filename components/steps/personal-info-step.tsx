

"use client"

import { useState, useEffect, useCallback } from "react"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select"
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card"
import { Separator } from "@/components/ui/separator"
import { User, Users, ArrowRight, FileText, AlertCircle } from "lucide-react"
import { Alert, AlertDescription } from "@/components/ui/alert"

interface PersonalInfoStepProps {
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

const filingStatusOptions = [
  { value: "SINGLE", label: "Single" },
  { value: "MARRIED_FILING_JOINTLY", label: "Married Filing Jointly" },
  { value: "MARRIED_FILING_SEPARATELY", label: "Married Filing Separately" },
  { value: "HEAD_OF_HOUSEHOLD", label: "Head of Household" },
  { value: "QUALIFYING_SURVIVING_SPOUSE", label: "Qualifying Surviving Spouse" },
]

export function PersonalInfoStep({ 
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
}: PersonalInfoStepProps) {
  const [formData, setFormData] = useState({
    filingStatus: taxReturn.filingStatus || "SINGLE",
    firstName: taxReturn.firstName || "",
    lastName: taxReturn.lastName || "",
    ssn: taxReturn.ssn || "",
    spouseFirstName: taxReturn.spouseFirstName || "",
    spouseLastName: taxReturn.spouseLastName || "",
    spouseSsn: taxReturn.spouseSsn || "",
    address: taxReturn.address || "",
    city: taxReturn.city || "",
    state: taxReturn.state || "",
    zipCode: taxReturn.zipCode || "",
  })

  const [w2PersonalInfo, setW2PersonalInfo] = useState<any>(null)
  const [isLoadingW2Data, setIsLoadingW2Data] = useState(false)
  const [w2DataSource, setW2DataSource] = useState<string>('')

  // Load W2 personal information when component mounts
  useEffect(() => {
    const loadW2PersonalInfo = async () => {
      if (!taxReturn.id) return;
      
      setIsLoadingW2Data(true);
      try {
        console.log('🔍 [Personal Info] Loading W2 data for tax return:', taxReturn.id);
        
        const response = await fetch(`/api/tax-returns/${taxReturn.id}/form-1040`);
        if (response.ok) {
          const data = await response.json();
          console.log('✅ [Personal Info] Received 1040 data:', data);
          
          // DEBUG: Check the structure of received data
          console.log('🔍 [Personal Info DEBUG] Checking data structure:');
          console.log('  - data.form1040Data exists:', !!data.form1040Data);
          console.log('  - data.form1040Data.personalInfo exists:', !!data.form1040Data?.personalInfo);
          console.log('  - Full form1040Data:', JSON.stringify(data.form1040Data, null, 2));
          
          if (data.form1040Data?.personalInfo) {
            const personalInfo = data.form1040Data.personalInfo;
            console.log('✅ [Personal Info] Found W2 personal info:', personalInfo);
            
            setW2PersonalInfo(personalInfo);
            setW2DataSource(personalInfo.sourceDocument || 'W2');
            
            // Auto-populate form with W2 data, overriding existing values
            setFormData(prev => ({
              ...prev,
              firstName: personalInfo.firstName || prev.firstName,
              lastName: personalInfo.lastName || prev.lastName,
              ssn: personalInfo.ssn || prev.ssn,
              address: personalInfo.address || prev.address,
              city: personalInfo.city || prev.city,
              state: personalInfo.state || prev.state,
              zipCode: personalInfo.zipCode || prev.zipCode,
            }));
            
            console.log('✅ [Personal Info] Auto-populated form with W2 data');
          } else {
            console.log('⚠️ [Personal Info] No W2 personal info found in 1040 data');
            console.log('🔍 [Personal Info DEBUG] Available form1040Data keys:', Object.keys(data.form1040Data || {}));
            console.log('🔍 [Personal Info DEBUG] Checking for individual fields:');
            console.log('  - firstName:', data.form1040Data?.firstName);
            console.log('  - lastName:', data.form1040Data?.lastName);
            console.log('  - ssn:', data.form1040Data?.ssn);
            console.log('  - address:', data.form1040Data?.address);
            console.log('  - city:', data.form1040Data?.city);
            console.log('  - state:', data.form1040Data?.state);
            console.log('  - zipCode:', data.form1040Data?.zipCode);
            
            // FALLBACK: Check if individual fields from W2 are available even without personalInfo object
            const hasIndividualW2Fields = !!(
              data.form1040Data?.firstName || 
              data.form1040Data?.lastName || 
              data.form1040Data?.ssn || 
              data.form1040Data?.address
            );
            
            if (hasIndividualW2Fields) {
              console.log('✅ [Personal Info] Found individual W2 fields, auto-populating form');
              setW2DataSource('W2 (individual fields)');
              
              // Create a synthetic personalInfo object for display purposes
              const syntheticPersonalInfo = {
                firstName: data.form1040Data.firstName || '',
                lastName: data.form1040Data.lastName || '',
                ssn: data.form1040Data.ssn || '',
                address: data.form1040Data.address || '',
                city: data.form1040Data.city || '',
                state: data.form1040Data.state || '',
                zipCode: data.form1040Data.zipCode || '',
                sourceDocument: 'W2',
                sourceDocumentId: 'individual-fields'
              };
              
              setW2PersonalInfo(syntheticPersonalInfo);
              
              // Auto-populate form with individual W2 fields
              setFormData(prev => ({
                ...prev,
                firstName: data.form1040Data.firstName || prev.firstName,
                lastName: data.form1040Data.lastName || prev.lastName,
                ssn: data.form1040Data.ssn || prev.ssn,
                address: data.form1040Data.address || prev.address,
                city: data.form1040Data.city || prev.city,
                state: data.form1040Data.state || prev.state,
                zipCode: data.form1040Data.zipCode || prev.zipCode,
              }));
              
              console.log('✅ [Personal Info] Auto-populated form with individual W2 fields');
            }
          }
        } else {
          console.log('⚠️ [Personal Info] Failed to load 1040 data:', response.status);
        }
      } catch (error) {
        console.error('❌ [Personal Info] Error loading W2 data:', error);
      } finally {
        setIsLoadingW2Data(false);
      }
    };

    loadW2PersonalInfo();
  }, [taxReturn.id]);

  // Auto-save functionality with debouncing
  useEffect(() => {
    const timer = setTimeout(() => {
      if (hasUnsavedChanges && !autoSaving) {
        onAutoSave(formData)
      }
    }, 2000) // Auto-save after 2 seconds of inactivity

    return () => clearTimeout(timer)
  }, [formData, hasUnsavedChanges, autoSaving, onAutoSave])

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    await onUpdate(formData)
    onNext()
  }

  const handleSaveAndContinue = async (e: React.FormEvent) => {
    e.preventDefault()
    await onCompleteStep(formData)
    onNext()
  }

  const handleChange = (field: string, value: string) => {
    setFormData(prev => ({ ...prev, [field]: value }))
    onMarkUnsaved()
  }

  const handleSaveOnly = async () => {
    await onAutoSave(formData)
  }

  const isMarried = formData.filingStatus === "MARRIED_FILING_JOINTLY" || formData.filingStatus === "MARRIED_FILING_SEPARATELY"

  const isStepOneComplete = formData.firstName && formData.lastName && formData.filingStatus

  if (taxReturn.currentStep === 1) {
    return (
      <form onSubmit={handleSubmit} className="space-y-6">
        <div className="text-center mb-8">
          <h2 className="text-2xl font-bold text-gray-900 mb-2">Let's get started!</h2>
          <p className="text-gray-600">We'll guide you through your tax return step by step.</p>
        </div>

        <div className="space-y-6">
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center space-x-2">
                <User className="h-5 w-5" />
                <span>Basic Information</span>
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <Label htmlFor="firstName">First Name *</Label>
                  <Input
                    id="firstName"
                    value={formData.firstName}
                    onChange={(e) => handleChange("firstName", e.target.value)}
                    required
                  />
                </div>
                <div>
                  <Label htmlFor="lastName">Last Name *</Label>
                  <Input
                    id="lastName"
                    value={formData.lastName}
                    onChange={(e) => handleChange("lastName", e.target.value)}
                    required
                  />
                </div>
              </div>
              
              <div>
                <Label htmlFor="filingStatus">Filing Status *</Label>
                <Select value={formData.filingStatus} onValueChange={(value) => handleChange("filingStatus", value)}>
                  <SelectTrigger>
                    <SelectValue placeholder="Select filing status" />
                  </SelectTrigger>
                  <SelectContent>
                    {filingStatusOptions.map((option) => (
                      <SelectItem key={option.value} value={option.value}>
                        {option.label}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            </CardContent>
          </Card>

          <div className="flex justify-between items-center">
            <div className="flex items-center space-x-2">
              {hasUnsavedChanges && !autoSaving && (
                <Button
                  type="button"
                  variant="outline"
                  onClick={handleSaveOnly}
                  disabled={autoSaving}
                  size="sm"
                >
                  {autoSaving ? "Saving..." : "Save"}
                </Button>
              )}
              {autoSaving && (
                <div className="flex items-center space-x-1 text-blue-600 text-sm">
                  <div className="animate-spin h-3 w-3 border border-blue-600 border-t-transparent rounded-full"></div>
                  <span>Auto-saving...</span>
                </div>
              )}
            </div>
            
            <div className="flex space-x-2">
              <Button 
                type="button"
                variant="outline"
                onClick={handleSaveAndContinue}
                disabled={!isStepOneComplete || saving}
              >
                {saving ? "Saving..." : "Save & Continue"}
                <ArrowRight className="ml-2 h-4 w-4" />
              </Button>
              <Button type="submit" disabled={!isStepOneComplete || saving}>
                {saving ? "Saving..." : "Continue"}
                <ArrowRight className="ml-2 h-4 w-4" />
              </Button>
            </div>
          </div>
        </div>
      </form>
    )
  }

  return (
    <form onSubmit={handleSubmit} className="space-y-6">
      <div className="space-y-6">
        {/* W2 Data Source Alert */}
        {isLoadingW2Data && (
          <Alert>
            <AlertCircle className="h-4 w-4" />
            <AlertDescription>
              Loading personal information from your uploaded documents...
            </AlertDescription>
          </Alert>
        )}
        
        {w2PersonalInfo && (
          <Alert>
            <FileText className="h-4 w-4" />
            <AlertDescription>
              Personal information has been automatically populated from your uploaded {w2DataSource} document. 
              You can review and modify the information below if needed.
            </AlertDescription>
          </Alert>
        )}

        <Card>
          <CardHeader>
            <CardTitle className="flex items-center space-x-2">
              <User className="h-5 w-5" />
              <span>Personal Information</span>
              {w2PersonalInfo && (
                <span className="text-sm text-green-600 bg-green-100 px-2 py-1 rounded-full">
                  Auto-filled from {w2DataSource}
                </span>
              )}
            </CardTitle>
            <CardDescription>
              {w2PersonalInfo 
                ? "Information below was automatically extracted from your W2 document. Please review and update if necessary."
                : "Enter your personal details as they appear on your tax documents"
              }
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
            <div className="grid grid-cols-2 gap-4">
              <div>
                <Label htmlFor="firstName">First Name *</Label>
                <Input
                  id="firstName"
                  value={formData.firstName}
                  onChange={(e) => handleChange("firstName", e.target.value)}
                  required
                  className={w2PersonalInfo ? "bg-green-50 border-green-200" : ""}
                />
                {w2PersonalInfo && (
                  <p className="text-xs text-green-600 mt-1">From W2: {w2PersonalInfo.firstName}</p>
                )}
              </div>
              <div>
                <Label htmlFor="lastName">Last Name *</Label>
                <Input
                  id="lastName"
                  value={formData.lastName}
                  onChange={(e) => handleChange("lastName", e.target.value)}
                  required
                  className={w2PersonalInfo ? "bg-green-50 border-green-200" : ""}
                />
                {w2PersonalInfo && (
                  <p className="text-xs text-green-600 mt-1">From W2: {w2PersonalInfo.lastName}</p>
                )}
              </div>
            </div>
            
            <div>
              <Label htmlFor="ssn">Social Security Number *</Label>
              <Input
                id="ssn"
                value={formData.ssn}
                onChange={(e) => handleChange("ssn", e.target.value)}
                placeholder="000-00-0000"
                maxLength={11}
                required
                className={w2PersonalInfo ? "bg-green-50 border-green-200" : ""}
              />
              {w2PersonalInfo && (
                <p className="text-xs text-green-600 mt-1">From W2: {w2PersonalInfo.ssn}</p>
              )}
            </div>
            
            <div>
              <Label htmlFor="filingStatus">Filing Status *</Label>
              <Select value={formData.filingStatus} onValueChange={(value) => handleChange("filingStatus", value)}>
                <SelectTrigger>
                  <SelectValue placeholder="Select filing status" />
                </SelectTrigger>
                <SelectContent>
                  {filingStatusOptions.map((option) => (
                    <SelectItem key={option.value} value={option.value}>
                      {option.label}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
          </CardContent>
        </Card>

        {isMarried && (
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center space-x-2">
                <Users className="h-5 w-5" />
                <span>Spouse Information</span>
              </CardTitle>
              <CardDescription>
                Enter your spouse's information
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <Label htmlFor="spouseFirstName">Spouse First Name *</Label>
                  <Input
                    id="spouseFirstName"
                    value={formData.spouseFirstName}
                    onChange={(e) => handleChange("spouseFirstName", e.target.value)}
                    required
                  />
                </div>
                <div>
                  <Label htmlFor="spouseLastName">Spouse Last Name *</Label>
                  <Input
                    id="spouseLastName"
                    value={formData.spouseLastName}
                    onChange={(e) => handleChange("spouseLastName", e.target.value)}
                    required
                  />
                </div>
              </div>
              
              <div>
                <Label htmlFor="spouseSsn">Spouse Social Security Number *</Label>
                <Input
                  id="spouseSsn"
                  value={formData.spouseSsn}
                  onChange={(e) => handleChange("spouseSsn", e.target.value)}
                  placeholder="000-00-0000"
                  maxLength={11}
                  required
                />
              </div>
            </CardContent>
          </Card>
        )}

        <Card>
          <CardHeader>
            <CardTitle className="flex items-center space-x-2">
              <span>Address Information</span>
              {w2PersonalInfo && (
                <span className="text-sm text-green-600 bg-green-100 px-2 py-1 rounded-full">
                  Auto-filled from {w2DataSource}
                </span>
              )}
            </CardTitle>
            <CardDescription>
              {w2PersonalInfo 
                ? "Address information was automatically extracted from your W2 document. Please review and update if necessary."
                : "Enter your current address"
              }
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
            <div>
              <Label htmlFor="address">Street Address *</Label>
              <Input
                id="address"
                value={formData.address}
                onChange={(e) => handleChange("address", e.target.value)}
                required
                className={w2PersonalInfo ? "bg-green-50 border-green-200" : ""}
              />
              {w2PersonalInfo && (
                <p className="text-xs text-green-600 mt-1">From W2: {w2PersonalInfo.address}</p>
              )}
            </div>
            
            <div className="grid grid-cols-2 gap-4">
              <div>
                <Label htmlFor="city">City *</Label>
                <Input
                  id="city"
                  value={formData.city}
                  onChange={(e) => handleChange("city", e.target.value)}
                  required
                  className={w2PersonalInfo ? "bg-green-50 border-green-200" : ""}
                />
                {w2PersonalInfo && (
                  <p className="text-xs text-green-600 mt-1">From W2: {w2PersonalInfo.city}</p>
                )}
              </div>
              <div>
                <Label htmlFor="state">State *</Label>
                <Input
                  id="state"
                  value={formData.state}
                  onChange={(e) => handleChange("state", e.target.value)}
                  maxLength={2}
                  placeholder="CA"
                  required
                  className={w2PersonalInfo ? "bg-green-50 border-green-200" : ""}
                />
                {w2PersonalInfo && (
                  <p className="text-xs text-green-600 mt-1">From W2: {w2PersonalInfo.state}</p>
                )}
              </div>
            </div>
            
            <div className="w-1/2">
              <Label htmlFor="zipCode">ZIP Code *</Label>
              <Input
                id="zipCode"
                value={formData.zipCode}
                onChange={(e) => handleChange("zipCode", e.target.value)}
                maxLength={10}
                placeholder="12345"
                required
                className={w2PersonalInfo ? "bg-green-50 border-green-200" : ""}
              />
              {w2PersonalInfo && (
                <p className="text-xs text-green-600 mt-1">From W2: {w2PersonalInfo.zipCode}</p>
              )}
            </div>
          </CardContent>
        </Card>

        <div className="flex justify-between items-center">
          <div className="flex items-center space-x-2">
            <Button type="button" variant="outline" onClick={onPrev} disabled={taxReturn.currentStep === 1}>
              Back
            </Button>
            
            {hasUnsavedChanges && !autoSaving && (
              <Button
                type="button"
                variant="outline"
                onClick={handleSaveOnly}
                disabled={autoSaving}
                size="sm"
              >
                {autoSaving ? "Saving..." : "Save"}
              </Button>
            )}
            {autoSaving && (
              <div className="flex items-center space-x-1 text-blue-600 text-sm">
                <div className="animate-spin h-3 w-3 border border-blue-600 border-t-transparent rounded-full"></div>
                <span>Auto-saving...</span>
              </div>
            )}
          </div>
          
          <div className="flex space-x-2">
            <Button 
              type="button"
              variant="outline"
              onClick={handleSaveAndContinue}
              disabled={saving}
            >
              {saving ? "Saving..." : "Save & Continue"}
              <ArrowRight className="ml-2 h-4 w-4" />
            </Button>
            <Button type="submit" disabled={saving}>
              {saving ? "Saving..." : "Continue"}
              <ArrowRight className="ml-2 h-4 w-4" />
            </Button>
          </div>
        </div>
      </div>
    </form>
  )
}

