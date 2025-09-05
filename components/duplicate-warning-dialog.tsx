
'use client';

import { useState } from 'react';
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { 
  AlertTriangle, 
  FileText, 
  Calendar, 
  Building, 
  User, 
  DollarSign,
  CheckCircle2,
  XCircle,
  RefreshCw
} from 'lucide-react';

interface DuplicateDetectionResult {
  isDuplicate: boolean;
  confidence: number;
  matchingDocuments: Array<{
    id: string;
    fileName: string;
    documentType: string;
    matchedFields: string[];
    similarity: number;
    createdAt: Date;
  }>;
  matchCriteria: {
    documentType: boolean;
    employerInfo: boolean;
    recipientInfo: boolean;
    amountSimilarity: boolean;
    nameSimilarity: boolean;
  };
}

interface DuplicateWarningDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  documentFileName: string;
  duplicateDetection: DuplicateDetectionResult;
  onAction: (action: 'proceed' | 'cancel' | 'replace', replacementDocumentId?: string) => void;
  isProcessing?: boolean;
}

export function DuplicateWarningDialog({
  open,
  onOpenChange,
  documentFileName,
  duplicateDetection,
  onAction,
  isProcessing = false
}: DuplicateWarningDialogProps) {
  const [selectedAction, setSelectedAction] = useState<'proceed' | 'cancel' | 'replace'>('cancel');
  const [selectedReplacementId, setSelectedReplacementId] = useState<string>('');

  const handleAction = () => {
    if (selectedAction === 'replace' && selectedReplacementId) {
      onAction('replace', selectedReplacementId);
    } else {
      onAction(selectedAction);
    }
  };

  const getMatchedFieldsDisplay = (fields: string[]) => {
    const fieldLabels: { [key: string]: string } = {
      'employerName': 'Employer Name',
      'employerEIN': 'Employer EIN',
      'employeeName': 'Employee Name',
      'employeeSSN': 'Employee SSN',
      'payerName': 'Payer Name',
      'payerTIN': 'Payer TIN',
      'recipientName': 'Recipient Name',
      'recipientTIN': 'Recipient TIN',
      'wages': 'Wages',
      'federalTaxWithheld': 'Federal Tax Withheld',
      'interestIncome': 'Interest Income',
      'ordinaryDividends': 'Ordinary Dividends',
      'nonemployeeCompensation': 'Nonemployee Compensation',
      'documentType': 'Document Type'
    };

    return fields.map(field => fieldLabels[field] || field).join(', ');
  };

  const getConfidenceColor = (confidence: number) => {
    if (confidence >= 0.9) return 'destructive';
    if (confidence >= 0.8) return 'secondary';
    return 'outline';
  };

  const getConfidenceText = (confidence: number) => {
    if (confidence >= 0.9) return 'Very High';
    if (confidence >= 0.8) return 'High';
    if (confidence >= 0.7) return 'Medium';
    return 'Low';
  };

  const bestMatch = duplicateDetection.matchingDocuments?.[0];

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-2xl">
        <DialogHeader>
          <DialogTitle className="flex items-center gap-2">
            <AlertTriangle className="h-5 w-5 text-amber-500" />
            Potential Duplicate Document Detected
          </DialogTitle>
          <DialogDescription>
            The document "{documentFileName}" appears to be similar to existing documents in your tax return.
          </DialogDescription>
        </DialogHeader>

        <div className="space-y-4">
          {/* Confidence Score */}
          <Alert>
            <AlertTriangle className="h-4 w-4" />
            <AlertDescription>
              <div className="flex items-center justify-between">
                <span>Similarity confidence:</span>
                <Badge variant={getConfidenceColor(duplicateDetection.confidence)}>
                  {getConfidenceText(duplicateDetection.confidence)} ({Math.round(duplicateDetection.confidence * 100)}%)
                </Badge>
              </div>
            </AlertDescription>
          </Alert>

          {/* Match Criteria */}
          <div className="grid grid-cols-2 gap-4 p-4 bg-muted rounded-lg">
            <div className="space-y-2">
              <h4 className="font-semibold text-sm">Matching Criteria:</h4>
              <div className="space-y-1">
                {duplicateDetection.matchCriteria.documentType && (
                  <div className="flex items-center gap-2 text-sm">
                    <CheckCircle2 className="h-3 w-3 text-green-500" />
                    Document Type
                  </div>
                )}
                {duplicateDetection.matchCriteria.employerInfo && (
                  <div className="flex items-center gap-2 text-sm">
                    <CheckCircle2 className="h-3 w-3 text-green-500" />
                    Employer/Payer Info
                  </div>
                )}
                {duplicateDetection.matchCriteria.recipientInfo && (
                  <div className="flex items-center gap-2 text-sm">
                    <CheckCircle2 className="h-3 w-3 text-green-500" />
                    Recipient/Employee Info
                  </div>
                )}
                {duplicateDetection.matchCriteria.amountSimilarity && (
                  <div className="flex items-center gap-2 text-sm">
                    <CheckCircle2 className="h-3 w-3 text-green-500" />
                    Income/Tax Amounts
                  </div>
                )}
                {duplicateDetection.matchCriteria.nameSimilarity && (
                  <div className="flex items-center gap-2 text-sm">
                    <CheckCircle2 className="h-3 w-3 text-green-500" />
                    Name Similarity
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* Matching Documents */}
          {duplicateDetection.matchingDocuments.length > 0 && (
            <div className="space-y-3">
              <h4 className="font-semibold">Similar Documents Found:</h4>
              {duplicateDetection.matchingDocuments.map((match) => (
                <div key={match.id} className="border rounded-lg p-3 space-y-2">
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-2">
                      <FileText className="h-4 w-4" />
                      <span className="font-medium">{match.fileName}</span>
                    </div>
                    <Badge variant="outline">
                      {Math.round(match.similarity * 100)}% match
                    </Badge>
                  </div>
                  
                  <div className="text-sm text-muted-foreground">
                    <div className="flex items-center gap-4">
                      <span className="flex items-center gap-1">
                        <Calendar className="h-3 w-3" />
                        {new Date(match.createdAt).toLocaleDateString()}
                      </span>
                      <span>{match.documentType}</span>
                    </div>
                  </div>

                  {match.matchedFields.length > 0 && (
                    <div className="text-sm">
                      <span className="font-medium">Matched fields: </span>
                      <span className="text-muted-foreground">
                        {getMatchedFieldsDisplay(match.matchedFields)}
                      </span>
                    </div>
                  )}
                </div>
              ))}
            </div>
          )}

          {/* Action Selection */}
          <div className="space-y-3">
            <h4 className="font-semibold">What would you like to do?</h4>
            
            <div className="space-y-2">
              <label className="flex items-center space-x-3 p-3 border rounded-lg cursor-pointer hover:bg-muted">
                <input
                  type="radio"
                  name="action"
                  value="cancel"
                  checked={selectedAction === 'cancel'}
                  onChange={(e) => setSelectedAction(e.target.value as any)}
                  className="w-4 h-4"
                />
                <div className="flex items-center gap-2">
                  <XCircle className="h-4 w-4 text-red-500" />
                  <div>
                    <div className="font-medium">Cancel Import</div>
                    <div className="text-sm text-muted-foreground">
                      Don't import this document because it's a duplicate
                    </div>
                  </div>
                </div>
              </label>

              <label className="flex items-center space-x-3 p-3 border rounded-lg cursor-pointer hover:bg-muted">
                <input
                  type="radio"
                  name="action"
                  value="proceed"
                  checked={selectedAction === 'proceed'}
                  onChange={(e) => setSelectedAction(e.target.value as any)}
                  className="w-4 h-4"
                />
                <div className="flex items-center gap-2">
                  <CheckCircle2 className="h-4 w-4 text-green-500" />
                  <div>
                    <div className="font-medium">Import Anyway</div>
                    <div className="text-sm text-muted-foreground">
                      Import this document despite the similarity warning
                    </div>
                  </div>
                </div>
              </label>

              {bestMatch && (
                <label className="flex items-center space-x-3 p-3 border rounded-lg cursor-pointer hover:bg-muted">
                  <input
                    type="radio"
                    name="action"
                    value="replace"
                    checked={selectedAction === 'replace'}
                    onChange={(e) => {
                      setSelectedAction(e.target.value as any);
                      setSelectedReplacementId(bestMatch.id);
                    }}
                    className="w-4 h-4"
                  />
                  <div className="flex items-center gap-2">
                    <RefreshCw className="h-4 w-4 text-blue-500" />
                    <div>
                      <div className="font-medium">Replace Existing</div>
                      <div className="text-sm text-muted-foreground">
                        Replace "{bestMatch.fileName}" with this new document
                      </div>
                    </div>
                  </div>
                </label>
              )}
            </div>
          </div>
        </div>

        <DialogFooter>
          <Button 
            variant="outline" 
            onClick={() => onOpenChange(false)}
            disabled={isProcessing}
          >
            Cancel
          </Button>
          <Button 
            onClick={handleAction}
            disabled={isProcessing}
            variant={selectedAction === 'cancel' ? 'destructive' : 'default'}
          >
            {isProcessing ? 'Processing...' : 
             selectedAction === 'cancel' ? 'Cancel Import' :
             selectedAction === 'replace' ? 'Replace Document' :
             'Import Anyway'}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
