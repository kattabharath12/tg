
"use client";

import { useState, useEffect } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Alert, AlertDescription } from "@/components/ui/alert";
import { Badge } from "@/components/ui/badge";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Form1040Interface } from "@/components/form-1040-interface";
import { Form1040Data } from "@/lib/form-1040-types";
import { FileText, ArrowLeft, ArrowRight, Download, Eye, CheckCircle, AlertCircle, RefreshCw } from "lucide-react";
import { toast } from "sonner";

interface Form1040StepProps {
  taxReturn: any;
  onUpdate: (data: any) => Promise<any>;
  onNext: () => void;
  onPrev: () => void;
  loading: boolean;
  saving: boolean;
}

export function Form1040Step({ 
  taxReturn, 
  onUpdate, 
  onNext, 
  onPrev, 
  loading, 
  saving 
}: Form1040StepProps) {
  const [form1040Data, setForm1040Data] = useState<Form1040Data | null>(null);
  const [w2MappingData, setW2MappingData] = useState<any[]>([]);
  const [isLoadingForm, setIsLoadingForm] = useState(true);
  const [isSavingForm, setIsSavingForm] = useState(false);
  const [isGeneratingPDF, setIsGeneratingPDF] = useState(false);
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false);
  const [activeTab, setActiveTab] = useState<'form' | 'preview' | 'mappings'>('form');

  useEffect(() => {
    loadForm1040Data();
  }, [taxReturn.id]);

  const loadForm1040Data = async () => {
    try {
      setIsLoadingForm(true);
      console.log('Loading 1040 form data for tax return:', taxReturn.id);
      
      const response = await fetch(`/api/tax-returns/${taxReturn.id}/form-1040`);
      
      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'Failed to load 1040 form data');
      }
      
      const data = await response.json();
      console.log('Loaded 1040 form data:', data);
      
      setForm1040Data(data.form1040Data);
      setW2MappingData(data.w2MappingData || []);
      
    } catch (error) {
      console.error('Error loading 1040 form data:', error);
      toast.error('Failed to load 1040 form data');
    } finally {
      setIsLoadingForm(false);
    }
  };

  const handleSave1040Form = async (formData: Form1040Data) => {
    try {
      setIsSavingForm(true);
      console.log('Saving 1040 form data:', formData);
      
      const response = await fetch(`/api/tax-returns/${taxReturn.id}/form-1040`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ form1040Data: formData }),
      });
      
      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'Failed to save 1040 form');
      }
      
      const result = await response.json();
      console.log('Saved 1040 form data:', result);
      
      // Update the tax return data to reflect the saved 1040 form
      await onUpdate({
        totalIncome: formData.line9,
        adjustedGrossIncome: formData.line11,
        standardDeduction: formData.line12,
        taxableIncome: formData.line15,
        taxLiability: formData.line16,
        totalWithholdings: formData.line25a,
        refundAmount: formData.line33,
        amountOwed: formData.line37,
        lastSavedAt: new Date()
      });
      
      setForm1040Data(formData);
      setHasUnsavedChanges(false);
      toast.success('Form 1040 saved successfully');
      
    } catch (error) {
      console.error('Error saving 1040 form:', error);
      toast.error('Failed to save 1040 form');
      throw error;
    } finally {
      setIsSavingForm(false);
    }
  };

  const handleGeneratePDF = async (formData: Form1040Data) => {
    try {
      setIsGeneratingPDF(true);
      console.log('Generating PDF for 1040 form');
      
      const response = await fetch(`/api/tax-returns/${taxReturn.id}/form-1040/pdf`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ 
          form1040Data: formData,
          pdfType: 'full'
        }),
      });
      
      if (!response.ok) {
        throw new Error('Failed to generate PDF');
      }
      
      // Download the PDF
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.style.display = 'none';
      a.href = url;
      a.download = `Form_1040_${formData.taxYear}_${formData.firstName}_${formData.lastName}.pdf`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
      
      toast.success('PDF generated and downloaded successfully');
      
    } catch (error) {
      console.error('Error generating PDF:', error);
      toast.error('Failed to generate PDF');
      throw error;
    } finally {
      setIsGeneratingPDF(false);
    }
  };

  const handlePreviewForm = async () => {
    if (!form1040Data) return;
    
    try {
      const previewUrl = `/api/tax-returns/${taxReturn.id}/form-1040/pdf`;
      window.open(previewUrl, '_blank');
    } catch (error) {
      console.error('Error opening preview:', error);
      toast.error('Failed to open preview');
    }
  };

  const handleContinue = async () => {
    if (hasUnsavedChanges) {
      toast.error('Please save your changes before continuing');
      return;
    }
    
    if (!form1040Data) {
      toast.error('Please complete the 1040 form before continuing');
      return;
    }
    
    // Mark this step as completed
    await onUpdate({
      currentStep: Math.max(taxReturn.currentStep, 8), // Assuming this is step 8
      completedSteps: [...(taxReturn.completedSteps || []), 8]
    });
    
    onNext();
  };

  if (isLoadingForm) {
    return (
      <div className="flex items-center justify-center py-12">
        <RefreshCw className="w-6 h-6 animate-spin mr-3" />
        <span>Loading Form 1040...</span>
      </div>
    );
  }

  return (
    <div className="space-y-6">
      {/* Step Header */}
      <div className="text-center mb-6">
        <h2 className="text-3xl font-bold mb-2">Form 1040 Review & Completion</h2>
        <p className="text-gray-600">
          Review your automatically populated tax form and make any necessary adjustments
        </p>
      </div>

      {/* W2 Mapping Alert */}
      {w2MappingData.length > 0 && (
        <Alert>
          <CheckCircle className="h-4 w-4" />
          <AlertDescription className="flex items-center justify-between">
            <span>
              Data from {w2MappingData.length} W-2 document(s) has been automatically mapped to your Form 1040. 
              Please review all amounts carefully.
            </span>
            <Button
              variant="outline"
              size="sm"
              onClick={() => setActiveTab('mappings')}
            >
              View Mappings
            </Button>
          </AlertDescription>
        </Alert>
      )}

      {/* Unsaved Changes Warning */}
      {hasUnsavedChanges && (
        <Alert variant="destructive">
          <AlertCircle className="h-4 w-4" />
          <AlertDescription>
            You have unsaved changes. Please save your form before continuing.
          </AlertDescription>
        </Alert>
      )}

      {/* Main Content Tabs */}
      <Tabs value={activeTab} onValueChange={(value) => setActiveTab(value as any)}>
        <TabsList className="grid w-full grid-cols-3">
          <TabsTrigger value="form">Edit Form</TabsTrigger>
          <TabsTrigger value="preview">Preview</TabsTrigger>
          <TabsTrigger value="mappings">W-2 Mappings</TabsTrigger>
        </TabsList>

        {/* Form Editing Tab */}
        <TabsContent value="form">
          {form1040Data ? (
            <Form1040Interface
              initialData={form1040Data}
              w2MappingData={w2MappingData}
              onSave={handleSave1040Form}
              onGeneratePDF={handleGeneratePDF}
              taxReturnId={taxReturn.id}
              readonly={false}
            />
          ) : (
            <Card>
              <CardContent className="py-12">
                <div className="text-center">
                  <FileText className="w-12 h-12 mx-auto mb-4 text-gray-400" />
                  <h3 className="text-lg font-medium mb-2">No Form 1040 Data Available</h3>
                  <p className="text-gray-600 mb-4">
                    Unable to load Form 1040 data. This may be because no W-2 documents have been processed yet.
                  </p>
                  <Button onClick={loadForm1040Data} variant="outline">
                    <RefreshCw className="w-4 h-4 mr-2" />
                    Retry Loading
                  </Button>
                </div>
              </CardContent>
            </Card>
          )}
        </TabsContent>

        {/* Preview Tab */}
        <TabsContent value="preview">
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center justify-between">
                <div className="flex items-center">
                  <Eye className="w-5 h-5 mr-2" />
                  Form 1040 Preview
                </div>
                <div className="flex space-x-2">
                  <Button
                    onClick={handlePreviewForm}
                    variant="outline"
                    size="sm"
                  >
                    <Eye className="w-4 h-4 mr-2" />
                    Open Preview
                  </Button>
                  <Button
                    onClick={() => form1040Data && handleGeneratePDF(form1040Data)}
                    disabled={!form1040Data || isGeneratingPDF}
                    size="sm"
                  >
                    <Download className="w-4 h-4 mr-2" />
                    {isGeneratingPDF ? 'Generating...' : 'Download PDF'}
                  </Button>
                </div>
              </CardTitle>
              <CardDescription>
                Preview your completed Form 1040 before final submission
              </CardDescription>
            </CardHeader>
            <CardContent>
              {form1040Data ? (
                <div className="space-y-4">
                  {/* Quick Summary */}
                  <div className="grid grid-cols-3 gap-4">
                    <div className="bg-blue-50 p-4 rounded-lg border border-blue-200">
                      <div className="text-sm text-blue-600 font-medium">Total Income</div>
                      <div className="text-2xl font-bold text-blue-800">
                        ${form1040Data.line9?.toLocaleString() || '0'}
                      </div>
                    </div>
                    <div className="bg-orange-50 p-4 rounded-lg border border-orange-200">
                      <div className="text-sm text-orange-600 font-medium">Total Tax</div>
                      <div className="text-2xl font-bold text-orange-800">
                        ${form1040Data.line24?.toLocaleString() || '0'}
                      </div>
                    </div>
                    <div className={`p-4 rounded-lg border ${
                      (form1040Data.line33 || 0) > 0 
                        ? 'bg-green-50 border-green-200' 
                        : (form1040Data.line37 || 0) > 0
                        ? 'bg-red-50 border-red-200'
                        : 'bg-gray-50 border-gray-200'
                    }`}>
                      <div className={`text-sm font-medium ${
                        (form1040Data.line33 || 0) > 0 
                          ? 'text-green-600' 
                          : (form1040Data.line37 || 0) > 0
                          ? 'text-red-600'
                          : 'text-gray-600'
                      }`}>
                        {(form1040Data.line33 || 0) > 0 
                          ? 'Refund Expected' 
                          : (form1040Data.line37 || 0) > 0
                          ? 'Amount Owed'
                          : 'Tax Balance'
                        }
                      </div>
                      <div className={`text-2xl font-bold ${
                        (form1040Data.line33 || 0) > 0 
                          ? 'text-green-800' 
                          : (form1040Data.line37 || 0) > 0
                          ? 'text-red-800'
                          : 'text-gray-800'
                      }`}>
                        ${((form1040Data.line33 || 0) > 0 
                          ? form1040Data.line33 
                          : form1040Data.line37 || 0
                        ).toLocaleString()}
                      </div>
                    </div>
                  </div>
                  
                  <div className="text-center py-6">
                    <p className="text-gray-600">
                      Click "Open Preview" to see the complete form, or "Download PDF" to save a copy.
                    </p>
                  </div>
                </div>
              ) : (
                <div className="text-center py-12">
                  <FileText className="w-12 h-12 mx-auto mb-4 text-gray-400" />
                  <p className="text-gray-600">No form data available for preview</p>
                </div>
              )}
            </CardContent>
          </Card>
        </TabsContent>

        {/* W-2 Mappings Tab */}
        <TabsContent value="mappings">
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center">
                <FileText className="w-5 h-5 mr-2" />
                W-2 to Form 1040 Field Mappings
              </CardTitle>
              <CardDescription>
                Review how data from your W-2 documents was mapped to Form 1040 fields
              </CardDescription>
            </CardHeader>
            <CardContent>
              {w2MappingData.length > 0 ? (
                <div className="space-y-6">
                  {w2MappingData.map((w2Doc, index) => (
                    <div key={w2Doc.documentId} className="border rounded-lg p-4">
                      <div className="flex items-center justify-between mb-4">
                        <h4 className="font-medium text-lg">W-2 Document #{index + 1}</h4>
                        <Badge variant="secondary">{w2Doc.fileName}</Badge>
                      </div>
                      
                      <div className="grid gap-3">
                        {w2Doc.mappings?.map((mapping: any, mappingIndex: number) => (
                          <div key={mappingIndex} className="flex items-center justify-between py-2 border-b border-gray-100 last:border-b-0">
                            <div className="flex-1">
                              <div className="font-medium text-sm">{mapping.description}</div>
                              <div className="text-xs text-gray-600">
                                {mapping.w2Field} → {mapping.form1040Line}
                              </div>
                            </div>
                            <div className="text-right">
                              <div className="font-medium">${mapping.form1040Value?.toLocaleString() || '0'}</div>
                              <div className="text-xs text-gray-600">{mapping.w2Value}</div>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <div className="text-center py-12">
                  <FileText className="w-12 h-12 mx-auto mb-4 text-gray-400" />
                  <h3 className="text-lg font-medium mb-2">No W-2 Mappings Available</h3>
                  <p className="text-gray-600">
                    No W-2 documents have been processed yet, or mapping data is not available.
                  </p>
                </div>
              )}
            </CardContent>
          </Card>
        </TabsContent>
      </Tabs>

      {/* Navigation Buttons */}
      <div className="flex justify-between pt-6 border-t">
        <Button
          type="button"
          variant="outline"
          onClick={onPrev}
          disabled={loading || saving}
        >
          <ArrowLeft className="w-4 h-4 mr-2" />
          Previous
        </Button>

        <div className="flex space-x-3">
          {form1040Data && (
            <Button
              type="button"
              variant="outline"
              onClick={() => handleGeneratePDF(form1040Data)}
              disabled={isGeneratingPDF}
            >
              <Download className="w-4 h-4 mr-2" />
              {isGeneratingPDF ? 'Generating...' : 'Download PDF'}
            </Button>
          )}
          
          <Button
            type="button"
            onClick={handleContinue}
            disabled={loading || saving || hasUnsavedChanges || !form1040Data}
          >
            Continue
            <ArrowRight className="w-4 h-4 ml-2" />
          </Button>
        </div>
      </div>
    </div>
  );
}
