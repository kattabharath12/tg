
import { NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { PDFGenerationService } from "@/lib/pdf-generation-service";
import { Form1040Data } from "@/lib/form-1040-types";

export const dynamic = "force-dynamic";

// POST: Generate PDF for Form 1040
export async function POST(
  request: Request,
  { params }: { params: { id: string } }
) {
  console.log("🔍 [1040 PDF] Starting PDF generation for tax return:", params.id);
  
  try {
    const session = await getServerSession(authOptions);
    
    if (!session?.user?.email) {
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
    }

    const user = await prisma.user.findUnique({
      where: { email: session.user.email }
    });

    if (!user) {
      return NextResponse.json({ error: "User not found" }, { status: 404 });
    }

    // Get the request body
    const body = await request.json();
    const { form1040Data, pdfType = 'full' }: { 
      form1040Data: Form1040Data, 
      pdfType?: 'full' | 'summary' 
    } = body;

    // Verify tax return ownership
    const taxReturn = await prisma.taxReturn.findFirst({
      where: { 
        id: params.id,
        userId: user.id 
      }
    });

    if (!taxReturn) {
      return NextResponse.json({ error: "Tax return not found" }, { status: 404 });
    }

    console.log("🔍 [1040 PDF] Generating PDF type:", pdfType);

    // Generate the appropriate PDF
    let pdfBuffer: Buffer;
    let fileName: string;

    if (pdfType === 'summary') {
      pdfBuffer = await PDFGenerationService.generateTaxSummaryPDF(form1040Data);
      fileName = `Tax_Summary_${form1040Data.taxYear}_${form1040Data.firstName}_${form1040Data.lastName}.pdf`;
    } else {
      pdfBuffer = await PDFGenerationService.generateForm1040PDF(form1040Data);
      fileName = `Form_1040_${form1040Data.taxYear}_${form1040Data.firstName}_${form1040Data.lastName}.pdf`;
    }

    console.log("✅ [1040 PDF] PDF generated successfully, size:", pdfBuffer.length, "bytes");

    // Set response headers for PDF download
    const headers = new Headers();
    headers.set('Content-Type', 'application/pdf');
    headers.set('Content-Disposition', `attachment; filename="${fileName}"`);
    headers.set('Content-Length', pdfBuffer.length.toString());

    return new NextResponse(pdfBuffer, {
      status: 200,
      headers
    });

  } catch (error) {
    console.error("💥 [1040 PDF] Error generating PDF:", error);
    return NextResponse.json(
      { error: "Failed to generate PDF" },
      { status: 500 }
    );
  }
}

// GET: Preview Form 1040 as HTML
export async function GET(
  request: Request,
  { params }: { params: { id: string } }
) {
  console.log("🔍 [1040 HTML] Starting HTML preview for tax return:", params.id);
  
  try {
    const session = await getServerSession(authOptions);
    
    if (!session?.user?.email) {
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
    }

    const user = await prisma.user.findUnique({
      where: { email: session.user.email }
    });

    if (!user) {
      return NextResponse.json({ error: "User not found" }, { status: 404 });
    }

    // Get tax return with 1040 data
    const taxReturn = await prisma.taxReturn.findFirst({
      where: { 
        id: params.id,
        userId: user.id 
      },
      include: {
        documents: {
          where: {
            documentType: 'OTHER_TAX_DOCUMENT',
            fileName: { contains: 'Form_1040' }
          }
        }
      }
    });

    if (!taxReturn) {
      return NextResponse.json({ error: "Tax return not found" }, { status: 404 });
    }

    // Find the 1040 form data
    const form1040Document = taxReturn.documents.find((doc: any) => 
      doc.fileName.includes('Form_1040')
    );

    if (!form1040Document?.extractedData) {
      return NextResponse.json({ error: "Form 1040 data not found" }, { status: 404 });
    }

    const form1040Data = form1040Document.extractedData as unknown as Form1040Data;

    // Generate HTML preview (simplified version)
    const htmlPreview = generateForm1040HTMLPreview(form1040Data);

    return new NextResponse(htmlPreview, {
      status: 200,
      headers: {
        'Content-Type': 'text/html'
      }
    });

  } catch (error) {
    console.error("💥 [1040 HTML] Error generating HTML preview:", error);
    return NextResponse.json(
      { error: "Failed to generate preview" },
      { status: 500 }
    );
  }
}

function generateForm1040HTMLPreview(formData: Form1040Data): string {
  const formatCurrency = (amount: number): string => {
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD',
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }).format(amount || 0);
  };

  const isRefund = (formData.line33 || 0) > 0;
  const finalAmount = isRefund ? formData.line33 : formData.line37;

  return `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Form 1040 Preview - ${formData.firstName} ${formData.lastName}</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            line-height: 1.6;
        }
        .header {
            text-align: center;
            border-bottom: 2px solid #333;
            padding-bottom: 20px;
            margin-bottom: 30px;
        }
        .section {
            margin: 20px 0;
            padding: 15px;
            border: 1px solid #ddd;
            border-radius: 5px;
        }
        .section h3 {
            margin-top: 0;
            color: #333;
            border-bottom: 1px solid #eee;
            padding-bottom: 10px;
        }
        .field-row {
            display: flex;
            justify-content: space-between;
            margin: 10px 0;
            padding: 5px 0;
        }
        .field-label {
            font-weight: bold;
        }
        .field-value {
            font-family: monospace;
        }
        .total-row {
            background-color: #f8f9fa;
            font-weight: bold;
            border-top: 2px solid #333;
            border-bottom: 2px solid #333;
        }
        .result-section {
            text-align: center;
            padding: 30px;
            margin: 20px 0;
            border-radius: 10px;
            ${isRefund ? 'background-color: #d4edda; border: 2px solid #28a745;' : 
              finalAmount > 0 ? 'background-color: #f8d7da; border: 2px solid #dc3545;' : 
              'background-color: #d1ecf1; border: 2px solid #17a2b8;'}
        }
        .result-amount {
            font-size: 36px;
            font-weight: bold;
            margin: 15px 0;
            ${isRefund ? 'color: #28a745;' : finalAmount > 0 ? 'color: #dc3545;' : 'color: #17a2b8;'}
        }
        .print-button {
            position: fixed;
            top: 20px;
            right: 20px;
            background: #007bff;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 5px;
            cursor: pointer;
        }
        @media print {
            .print-button {
                display: none;
            }
        }
    </style>
</head>
<body>
    <button class="print-button" onclick="window.print()">Print</button>
    
    <div class="header">
        <h1>Form 1040 - U.S. Individual Income Tax Return</h1>
        <h2>Tax Year ${formData.taxYear}</h2>
        <p><strong>${formData.firstName} ${formData.lastName}</strong></p>
        <p>SSN: ${formData.ssn}</p>
    </div>

    <div class="section">
        <h3>Personal Information</h3>
        <div class="field-row">
            <span class="field-label">Filing Status:</span>
            <span class="field-value">${formData.filingStatus.replace(/_/g, ' ')}</span>
        </div>
        <div class="field-row">
            <span class="field-label">Address:</span>
            <span class="field-value">${formData.address}, ${formData.city}, ${formData.state} ${formData.zipCode}</span>
        </div>
    </div>

    <div class="section">
        <h3>Income</h3>
        <div class="field-row">
            <span class="field-label">Line 1 - Wages from W-2:</span>
            <span class="field-value">${formatCurrency(formData.line1 || 0)}</span>
        </div>
        <div class="field-row">
            <span class="field-label">Line 2b - Taxable Interest:</span>
            <span class="field-value">${formatCurrency(formData.line2b || 0)}</span>
        </div>
        <div class="field-row">
            <span class="field-label">Line 3b - Ordinary Dividends:</span>
            <span class="field-value">${formatCurrency(formData.line3b || 0)}</span>
        </div>
        <div class="field-row total-row">
            <span class="field-label">Line 9 - Total Income:</span>
            <span class="field-value">${formatCurrency(formData.line9 || 0)}</span>
        </div>
    </div>

    <div class="section">
        <h3>Adjusted Gross Income</h3>
        <div class="field-row">
            <span class="field-label">Line 10 - Adjustments:</span>
            <span class="field-value">${formatCurrency(formData.line10 || 0)}</span>
        </div>
        <div class="field-row total-row">
            <span class="field-label">Line 11 - Adjusted Gross Income:</span>
            <span class="field-value">${formatCurrency(formData.line11 || 0)}</span>
        </div>
    </div>

    <div class="section">
        <h3>Deductions and Taxable Income</h3>
        <div class="field-row">
            <span class="field-label">Line 12 - Standard Deduction:</span>
            <span class="field-value">${formatCurrency(formData.line12 || 0)}</span>
        </div>
        <div class="field-row total-row">
            <span class="field-label">Line 15 - Taxable Income:</span>
            <span class="field-value">${formatCurrency(formData.line15 || 0)}</span>
        </div>
    </div>

    <div class="section">
        <h3>Tax and Credits</h3>
        <div class="field-row">
            <span class="field-label">Line 16 - Tax:</span>
            <span class="field-value">${formatCurrency(formData.line16 || 0)}</span>
        </div>
        <div class="field-row">
            <span class="field-label">Line 19 - Child Tax Credit:</span>
            <span class="field-value">${formatCurrency(formData.line19 || 0)}</span>
        </div>
        <div class="field-row total-row">
            <span class="field-label">Line 24 - Total Tax:</span>
            <span class="field-value">${formatCurrency(formData.line24 || 0)}</span>
        </div>
    </div>

    <div class="section">
        <h3>Payments</h3>
        <div class="field-row">
            <span class="field-label">Line 25a - Federal Tax Withheld:</span>
            <span class="field-value">${formatCurrency(formData.line25a || 0)}</span>
        </div>
        <div class="field-row">
            <span class="field-label">Line 25b - Estimated Tax Payments:</span>
            <span class="field-value">${formatCurrency(formData.line25b || 0)}</span>
        </div>
        <div class="field-row total-row">
            <span class="field-label">Line 32 - Total Payments:</span>
            <span class="field-value">${formatCurrency(formData.line32 || 0)}</span>
        </div>
    </div>

    <div class="result-section">
        <h2>${isRefund ? 'Expected Refund' : finalAmount > 0 ? 'Amount Owed' : 'Tax Balance'}</h2>
        <div class="result-amount">${formatCurrency(finalAmount || 0)}</div>
        <p>
            ${isRefund ? 'You overpaid your taxes and are entitled to a refund' : 
              finalAmount > 0 ? `Additional tax is owed. Payment due by April 15, ${formData.taxYear + 1}` : 
              'Your payments exactly match your tax liability'}
        </p>
    </div>

    <div style="text-align: center; margin-top: 40px; color: #666; font-size: 12px;">
        Generated on ${new Date().toLocaleDateString()} | Form 1040 (${formData.taxYear}) Preview
    </div>
</body>
</html>
  `;
}
