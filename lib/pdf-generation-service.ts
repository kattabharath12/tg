
import { Form1040Data, FilingStatus } from './form-1040-types';
import puppeteer from 'puppeteer';

export class PDFGenerationService {
  /**
   * Generates a PDF of the Form 1040 with the provided data
   */
  static async generateForm1040PDF(formData: Form1040Data): Promise<Buffer> {
    const browser = await puppeteer.launch({
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    try {
      const page = await browser.newPage();
      
      // Generate HTML content for the form
      const htmlContent = this.generateForm1040HTML(formData);
      
      // Set content and generate PDF
      await page.setContent(htmlContent, {
        waitUntil: 'networkidle0'
      });
      
      const pdfBuffer = await page.pdf({
        format: 'A4',
        margin: {
          top: '0.5in',
          right: '0.5in',
          bottom: '0.5in',
          left: '0.5in'
        },
        printBackground: true
      });
      
      return Buffer.from(pdfBuffer);
    } finally {
      await browser.close();
    }
  }

  /**
   * Generates HTML content for Form 1040
   */
  private static generateForm1040HTML(formData: Form1040Data): string {
    return `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Form 1040 - ${formData.firstName} ${formData.lastName}</title>
    <style>
        body {
            font-family: 'Times New Roman', serif;
            font-size: 10pt;
            line-height: 1.2;
            margin: 0;
            padding: 20px;
            color: #000;
        }
        
        .form-header {
            text-align: center;
            margin-bottom: 20px;
            border-bottom: 2px solid #000;
            padding-bottom: 10px;
        }
        
        .form-title {
            font-size: 16pt;
            font-weight: bold;
            margin: 0;
        }
        
        .form-subtitle {
            font-size: 12pt;
            margin: 5px 0;
        }
        
        .tax-year {
            font-size: 14pt;
            font-weight: bold;
            margin: 5px 0;
        }
        
        .section {
            margin: 15px 0;
            border: 1px solid #000;
            page-break-inside: avoid;
        }
        
        .section-header {
            background-color: #f0f0f0;
            padding: 5px;
            font-weight: bold;
            font-size: 11pt;
            border-bottom: 1px solid #000;
        }
        
        .section-content {
            padding: 10px;
        }
        
        .two-column {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            gap: 20px;
        }
        
        .column {
            flex: 1;
        }
        
        .field-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin: 8px 0;
            padding: 3px;
        }
        
        .field-label {
            flex: 1;
            font-weight: normal;
        }
        
        .field-value {
            text-align: right;
            font-weight: bold;
            min-width: 80px;
            padding: 2px 5px;
            border-bottom: 1px solid #000;
        }
        
        .line-number {
            font-weight: bold;
            margin-right: 5px;
        }
        
        .currency {
            font-family: monospace;
        }
        
        .total-line {
            border-top: 2px solid #000;
            border-bottom: 2px solid #000;
            background-color: #f8f8f8;
            font-weight: bold;
        }
        
        .signature-section {
            margin-top: 20px;
            border: 1px solid #000;
            padding: 10px;
        }
        
        .signature-line {
            border-bottom: 1px solid #000;
            height: 30px;
            margin: 10px 0;
        }
        
        .date-line {
            border-bottom: 1px solid #000;
            width: 100px;
            display: inline-block;
            margin-left: 10px;
        }
        
        .checkbox {
            display: inline-block;
            width: 12px;
            height: 12px;
            border: 1px solid #000;
            margin-right: 5px;
            text-align: center;
            line-height: 10px;
        }
        
        .checked {
            background-color: #000;
            color: #fff;
        }
        
        @media print {
            .section {
                page-break-inside: avoid;
            }
            
            .signature-section {
                page-break-inside: avoid;
            }
        }
    </style>
</head>
<body>
    <div class="form-header">
        <h1 class="form-title">Form 1040</h1>
        <div class="form-subtitle">U.S. Individual Income Tax Return</div>
        <div class="tax-year">${formData.taxYear}</div>
        <div>Department of the Treasury—Internal Revenue Service</div>
    </div>

    <!-- Personal Information Section -->
    <div class="section">
        <div class="section-header">Filing Status and Personal Information</div>
        <div class="section-content">
            <div class="two-column">
                <div class="column">
                    <div class="field-row">
                        <span class="field-label">First name and initial:</span>
                        <span class="field-value">${formData.firstName}</span>
                    </div>
                    <div class="field-row">
                        <span class="field-label">Last name:</span>
                        <span class="field-value">${formData.lastName}</span>
                    </div>
                    <div class="field-row">
                        <span class="field-label">Social Security Number:</span>
                        <span class="field-value">${formData.ssn}</span>
                    </div>
                </div>
                <div class="column">
                    ${formData.spouseFirstName ? `
                    <div class="field-row">
                        <span class="field-label">Spouse's first name and initial:</span>
                        <span class="field-value">${formData.spouseFirstName}</span>
                    </div>
                    <div class="field-row">
                        <span class="field-label">Spouse's last name:</span>
                        <span class="field-value">${formData.spouseLastName || ''}</span>
                    </div>
                    <div class="field-row">
                        <span class="field-label">Spouse's SSN:</span>
                        <span class="field-value">${formData.spouseSSN || ''}</span>
                    </div>
                    ` : ''}
                </div>
            </div>
            
            <div class="field-row">
                <span class="field-label">Home address:</span>
                <span class="field-value">${formData.address}</span>
            </div>
            <div class="field-row">
                <span class="field-label">City, state, and ZIP code:</span>
                <span class="field-value">${formData.city}, ${formData.state} ${formData.zipCode}</span>
            </div>
            
            <div style="margin-top: 15px;">
                <strong>Filing Status (Check only one):</strong><br>
                <label><span class="checkbox ${formData.filingStatus === FilingStatus.SINGLE ? 'checked' : ''}">✓</span> Single</label><br>
                <label><span class="checkbox ${formData.filingStatus === FilingStatus.MARRIED_FILING_JOINTLY ? 'checked' : ''}">✓</span> Married filing jointly</label><br>
                <label><span class="checkbox ${formData.filingStatus === FilingStatus.MARRIED_FILING_SEPARATELY ? 'checked' : ''}">✓</span> Married filing separately</label><br>
                <label><span class="checkbox ${formData.filingStatus === FilingStatus.HEAD_OF_HOUSEHOLD ? 'checked' : ''}">✓</span> Head of household</label><br>
                <label><span class="checkbox ${formData.filingStatus === FilingStatus.QUALIFYING_SURVIVING_SPOUSE ? 'checked' : ''}">✓</span> Qualifying surviving spouse</label>
            </div>
        </div>
    </div>

    <!-- Income Section -->
    <div class="section">
        <div class="section-header">Income</div>
        <div class="section-content">
            <div class="field-row">
                <span class="field-label"><span class="line-number">1.</span> Total amount from Form(s) W-2, box 1</span>
                <span class="field-value currency">${this.formatCurrency(formData.line1)}</span>
            </div>
            <div class="field-row">
                <span class="field-label"><span class="line-number">2a.</span> Tax-exempt interest</span>
                <span class="field-value currency">${this.formatCurrency(formData.line2a)}</span>
            </div>
            <div class="field-row">
                <span class="field-label"><span class="line-number">2b.</span> Taxable interest</span>
                <span class="field-value currency">${this.formatCurrency(formData.line2b)}</span>
            </div>
            <div class="field-row">
                <span class="field-label"><span class="line-number">3a.</span> Qualified dividends</span>
                <span class="field-value currency">${this.formatCurrency(formData.line3a)}</span>
            </div>
            <div class="field-row">
                <span class="field-label"><span class="line-number">3b.</span> Ordinary dividends</span>
                <span class="field-value currency">${this.formatCurrency(formData.line3b)}</span>
            </div>
            <div class="field-row">
                <span class="field-label"><span class="line-number">7.</span> Capital gain or (loss)</span>
                <span class="field-value currency">${this.formatCurrency(formData.line7)}</span>
            </div>
            <div class="field-row">
                <span class="field-label"><span class="line-number">8.</span> Additional income from Schedule 1, line 10</span>
                <span class="field-value currency">${this.formatCurrency(formData.line8)}</span>
            </div>
            <div class="field-row total-line">
                <span class="field-label"><span class="line-number">9.</span> Add lines 1, 2b, 3b, 4b, 5b, 6b, 7, and 8. This is your total income</span>
                <span class="field-value currency">${this.formatCurrency(formData.line9)}</span>
            </div>
        </div>
    </div>

    <!-- Adjusted Gross Income Section -->
    <div class="section">
        <div class="section-header">Adjusted Gross Income</div>
        <div class="section-content">
            <div class="field-row">
                <span class="field-label"><span class="line-number">10.</span> Adjustments to income from Schedule 1, line 26</span>
                <span class="field-value currency">${this.formatCurrency(formData.line10)}</span>
            </div>
            <div class="field-row total-line">
                <span class="field-label"><span class="line-number">11.</span> Subtract line 10 from line 9. This is your adjusted gross income</span>
                <span class="field-value currency">${this.formatCurrency(formData.line11)}</span>
            </div>
        </div>
    </div>

    <!-- Standard Deduction Section -->
    <div class="section">
        <div class="section-header">Standard Deduction or Itemized Deductions</div>
        <div class="section-content">
            <div class="field-row">
                <span class="field-label"><span class="line-number">12.</span> Standard deduction or itemized deductions</span>
                <span class="field-value currency">${this.formatCurrency(formData.line12)}</span>
            </div>
            <div class="field-row">
                <span class="field-label"><span class="line-number">13.</span> Qualified business income deduction</span>
                <span class="field-value currency">${this.formatCurrency(formData.line13)}</span>
            </div>
            <div class="field-row">
                <span class="field-label"><span class="line-number">14.</span> Add lines 12 and 13</span>
                <span class="field-value currency">${this.formatCurrency(formData.line14)}</span>
            </div>
            <div class="field-row total-line">
                <span class="field-label"><span class="line-number">15.</span> Subtract line 14 from line 11. This is your taxable income</span>
                <span class="field-value currency">${this.formatCurrency(formData.line15)}</span>
            </div>
        </div>
    </div>

    <!-- Tax and Credits Section -->
    <div class="section">
        <div class="section-header">Tax and Credits</div>
        <div class="section-content">
            <div class="field-row">
                <span class="field-label"><span class="line-number">16.</span> Tax</span>
                <span class="field-value currency">${this.formatCurrency(formData.line16)}</span>
            </div>
            <div class="field-row">
                <span class="field-label"><span class="line-number">17.</span> Amount from Schedule 2, line 3</span>
                <span class="field-value currency">${this.formatCurrency(formData.line17)}</span>
            </div>
            <div class="field-row">
                <span class="field-label"><span class="line-number">18.</span> Add lines 16 and 17</span>
                <span class="field-value currency">${this.formatCurrency(formData.line18)}</span>
            </div>
            <div class="field-row">
                <span class="field-label"><span class="line-number">19.</span> Child tax credit and credit for other dependents</span>
                <span class="field-value currency">${this.formatCurrency(formData.line19)}</span>
            </div>
            <div class="field-row">
                <span class="field-label"><span class="line-number">20.</span> Amount from Schedule 3, line 8</span>
                <span class="field-value currency">${this.formatCurrency(formData.line20)}</span>
            </div>
            <div class="field-row">
                <span class="field-label"><span class="line-number">21.</span> Add lines 19 and 20</span>
                <span class="field-value currency">${this.formatCurrency(formData.line21)}</span>
            </div>
            <div class="field-row">
                <span class="field-label"><span class="line-number">22.</span> Subtract line 21 from line 18</span>
                <span class="field-value currency">${this.formatCurrency(formData.line22)}</span>
            </div>
            <div class="field-row">
                <span class="field-label"><span class="line-number">23.</span> Other taxes from Schedule 2, line 21</span>
                <span class="field-value currency">${this.formatCurrency(formData.line23)}</span>
            </div>
            <div class="field-row total-line">
                <span class="field-label"><span class="line-number">24.</span> Add lines 22 and 23. This is your total tax</span>
                <span class="field-value currency">${this.formatCurrency(formData.line24)}</span>
            </div>
        </div>
    </div>

    <!-- Payments Section -->
    <div class="section">
        <div class="section-header">Payments</div>
        <div class="section-content">
            <div class="field-row">
                <span class="field-label"><span class="line-number">25a.</span> Federal income tax withheld from Forms W-2 and 1099</span>
                <span class="field-value currency">${this.formatCurrency(formData.line25a)}</span>
            </div>
            <div class="field-row">
                <span class="field-label"><span class="line-number">25b.</span> 2023 estimated tax payments</span>
                <span class="field-value currency">${this.formatCurrency(formData.line25b)}</span>
            </div>
            <div class="field-row">
                <span class="field-label"><span class="line-number">25c.</span> Earned income credit (EIC)</span>
                <span class="field-value currency">${this.formatCurrency(formData.line25c)}</span>
            </div>
            <div class="field-row">
                <span class="field-label"><span class="line-number">25d.</span> Additional child tax credit</span>
                <span class="field-value currency">${this.formatCurrency(formData.line25d)}</span>
            </div>
            <div class="field-row total-line">
                <span class="field-label"><span class="line-number">32.</span> Add lines 25a through 31. These are your total payments</span>
                <span class="field-value currency">${this.formatCurrency(formData.line32)}</span>
            </div>
        </div>
    </div>

    <!-- Refund or Amount Owed Section -->
    <div class="section">
        <div class="section-header">Refund or Amount You Owe</div>
        <div class="section-content">
            ${formData.line33 > 0 ? `
            <div class="field-row total-line" style="background-color: #e8f5e8;">
                <span class="field-label"><span class="line-number">33.</span> Overpaid amount (Refund)</span>
                <span class="field-value currency">${this.formatCurrency(formData.line33)}</span>
            </div>
            <div class="field-row">
                <span class="field-label"><span class="line-number">34.</span> Amount of line 33 to be refunded to you</span>
                <span class="field-value currency">${this.formatCurrency(formData.line34)}</span>
            </div>
            ${formData.line35a ? `
            <div style="margin-top: 10px;">
                <strong>Direct Deposit Information:</strong><br>
                Routing number: ${formData.line35a}<br>
                Account type: ${formData.line35b}<br>
                Account number: ${formData.line35c}
            </div>
            ` : ''}
            ` : formData.line37 > 0 ? `
            <div class="field-row total-line" style="background-color: #fde8e8;">
                <span class="field-label"><span class="line-number">37.</span> Amount you owe</span>
                <span class="field-value currency">${this.formatCurrency(formData.line37)}</span>
            </div>
            <div style="margin-top: 10px;">
                <strong>Payment Due Date:</strong> April 15, ${formData.taxYear + 1}
            </div>
            ` : `
            <div class="field-row total-line" style="background-color: #f0f8ff;">
                <span class="field-label">Tax Balance</span>
                <span class="field-value currency">$0.00</span>
            </div>
            <div style="margin-top: 10px;">
                Your payments exactly match your tax liability.
            </div>
            `}
        </div>
    </div>

    <!-- Third Party Designee Section -->
    <div class="section">
        <div class="section-header">Third Party Designee</div>
        <div class="section-content">
            <div>
                Do you want to allow another person to discuss this return with the IRS?
                <span class="checkbox ${formData.thirdPartyDesignee ? 'checked' : ''}">✓</span> Yes
                <span class="checkbox ${!formData.thirdPartyDesignee ? 'checked' : ''}">✓</span> No
            </div>
            ${formData.thirdPartyDesignee && formData.designeeName ? `
            <div style="margin-top: 10px;">
                <div>Designee's name: ${formData.designeeName}</div>
                <div>Phone no.: ${formData.designeePhone || ''}</div>
                <div>Personal identification number (PIN): ${formData.designeePin || ''}</div>
            </div>
            ` : ''}
        </div>
    </div>

    <!-- Signature Section -->
    <div class="signature-section">
        <div style="font-weight: bold; margin-bottom: 15px;">Sign Here</div>
        <div>Under penalties of perjury, I declare that I have examined this return and accompanying schedules and statements, and to the best of my knowledge and belief, they are true, correct, and complete.</div>
        
        <div style="margin-top: 20px;">
            <div style="display: flex; justify-content: space-between; align-items: flex-end;">
                <div style="flex: 1;">
                    <div>Your signature</div>
                    <div class="signature-line"></div>
                </div>
                <div style="width: 100px; margin-left: 20px;">
                    <div>Date</div>
                    <div class="date-line"></div>
                </div>
            </div>
            
            ${formData.filingStatus === FilingStatus.MARRIED_FILING_JOINTLY ? `
            <div style="display: flex; justify-content: space-between; align-items: flex-end; margin-top: 20px;">
                <div style="flex: 1;">
                    <div>Spouse's signature</div>
                    <div class="signature-line"></div>
                </div>
                <div style="width: 100px; margin-left: 20px;">
                    <div>Date</div>
                    <div class="date-line"></div>
                </div>
            </div>
            ` : ''}
        </div>
    </div>

    <!-- Paid Preparer Section (if applicable) -->
    ${formData.preparerName ? `
    <div class="signature-section" style="margin-top: 20px;">
        <div style="font-weight: bold; margin-bottom: 15px;">Paid Preparer Use Only</div>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
            <div>
                <div>Preparer's name: ${formData.preparerName}</div>
                <div>Preparer's signature: <div class="signature-line" style="display: inline-block; width: 200px;"></div></div>
                <div>Date: <div class="date-line"></div></div>
            </div>
            <div>
                <div>PTIN: ${formData.preparerSSN || ''}</div>
                <div>Firm's name: ${formData.preparerFirm || ''}</div>
                <div>Firm's address: ${formData.preparerAddress || ''}</div>
                <div>Phone no.: ${formData.preparerPhone || ''}</div>
            </div>
        </div>
    </div>
    ` : ''}

    <div style="text-align: center; margin-top: 30px; font-size: 8pt; color: #666;">
        Form 1040 (${formData.taxYear}) - Generated on ${new Date().toLocaleDateString()}
    </div>
</body>
</html>
    `;
  }

  /**
   * Formats currency values for display
   */
  private static formatCurrency(amount: number): string {
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD',
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }).format(amount || 0);
  }

  /**
   * Generates a simple tax summary document
   */
  static async generateTaxSummaryPDF(formData: Form1040Data): Promise<Buffer> {
    const browser = await puppeteer.launch({
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    try {
      const page = await browser.newPage();
      
      const htmlContent = this.generateTaxSummaryHTML(formData);
      
      await page.setContent(htmlContent, {
        waitUntil: 'networkidle0'
      });
      
      const pdfBuffer = await page.pdf({
        format: 'A4',
        margin: {
          top: '0.5in',
          right: '0.5in',
          bottom: '0.5in',
          left: '0.5in'
        },
        printBackground: true
      });
      
      return Buffer.from(pdfBuffer);
    } finally {
      await browser.close();
    }
  }

  /**
   * Generates HTML for tax summary
   */
  private static generateTaxSummaryHTML(formData: Form1040Data): string {
    const isRefund = formData.line33 > 0;
    const finalAmount = isRefund ? formData.line33 : formData.line37;
    
    return `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tax Summary - ${formData.firstName} ${formData.lastName}</title>
    <style>
        body {
            font-family: 'Arial', sans-serif;
            margin: 0;
            padding: 40px;
            color: #333;
            background-color: #f8f9fa;
        }
        
        .container {
            max-width: 800px;
            margin: 0 auto;
            background: white;
            padding: 40px;
            border-radius: 8px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }
        
        .header {
            text-align: center;
            margin-bottom: 40px;
            padding-bottom: 20px;
            border-bottom: 3px solid #007bff;
        }
        
        .title {
            font-size: 28px;
            font-weight: bold;
            color: #007bff;
            margin: 0;
        }
        
        .subtitle {
            font-size: 16px;
            color: #666;
            margin: 10px 0;
        }
        
        .result-section {
            text-align: center;
            margin: 40px 0;
            padding: 30px;
            border-radius: 8px;
            ${isRefund ? 'background-color: #d4edda; border: 2px solid #28a745;' : finalAmount > 0 ? 'background-color: #f8d7da; border: 2px solid #dc3545;' : 'background-color: #d1ecf1; border: 2px solid #17a2b8;'}
        }
        
        .result-title {
            font-size: 24px;
            font-weight: bold;
            margin-bottom: 15px;
            ${isRefund ? 'color: #155724;' : finalAmount > 0 ? 'color: #721c24;' : 'color: #0c5460;'}
        }
        
        .result-amount {
            font-size: 48px;
            font-weight: bold;
            margin: 20px 0;
            ${isRefund ? 'color: #28a745;' : finalAmount > 0 ? 'color: #dc3545;' : 'color: #17a2b8;'}
        }
        
        .summary-table {
            width: 100%;
            border-collapse: collapse;
            margin: 30px 0;
        }
        
        .summary-table th,
        .summary-table td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }
        
        .summary-table th {
            background-color: #f8f9fa;
            font-weight: bold;
        }
        
        .summary-table .amount {
            text-align: right;
            font-family: monospace;
        }
        
        .total-row {
            font-weight: bold;
            background-color: #f8f9fa;
            border-top: 2px solid #007bff;
        }
        
        .footer {
            text-align: center;
            margin-top: 40px;
            padding-top: 20px;
            border-top: 1px solid #ddd;
            color: #666;
            font-size: 14px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1 class="title">Tax Return Summary</h1>
            <div class="subtitle">Tax Year ${formData.taxYear}</div>
            <div class="subtitle">${formData.firstName} ${formData.lastName}</div>
        </div>

        <div class="result-section">
            <div class="result-title">
                ${isRefund ? 'Expected Refund' : finalAmount > 0 ? 'Amount Owed' : 'Tax Balance'}
            </div>
            <div class="result-amount">
                ${this.formatCurrency(finalAmount)}
            </div>
            <div>
                ${isRefund ? 'You overpaid your taxes and are entitled to a refund' : 
                  finalAmount > 0 ? 'Additional tax is owed. Payment due by April 15, ' + (formData.taxYear + 1) : 
                  'Your payments exactly match your tax liability'}
            </div>
        </div>

        <table class="summary-table">
            <thead>
                <tr>
                    <th>Description</th>
                    <th>Amount</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>Total Income (Line 9)</td>
                    <td class="amount">${this.formatCurrency(formData.line9)}</td>
                </tr>
                <tr>
                    <td>Adjustments to Income (Line 10)</td>
                    <td class="amount">${this.formatCurrency(formData.line10)}</td>
                </tr>
                <tr>
                    <td>Adjusted Gross Income (Line 11)</td>
                    <td class="amount">${this.formatCurrency(formData.line11)}</td>
                </tr>
                <tr>
                    <td>Standard Deduction (Line 12)</td>
                    <td class="amount">${this.formatCurrency(formData.line12)}</td>
                </tr>
                <tr>
                    <td>Taxable Income (Line 15)</td>
                    <td class="amount">${this.formatCurrency(formData.line15)}</td>
                </tr>
                <tr>
                    <td>Tax Liability (Line 16)</td>
                    <td class="amount">${this.formatCurrency(formData.line16)}</td>
                </tr>
                <tr class="total-row">
                    <td>Total Tax (Line 24)</td>
                    <td class="amount">${this.formatCurrency(formData.line24)}</td>
                </tr>
                <tr>
                    <td>Federal Tax Withheld (Line 25a)</td>
                    <td class="amount">${this.formatCurrency(formData.line25a)}</td>
                </tr>
                <tr>
                    <td>Estimated Tax Payments (Line 25b)</td>
                    <td class="amount">${this.formatCurrency(formData.line25b)}</td>
                </tr>
                <tr class="total-row">
                    <td>Total Payments (Line 32)</td>
                    <td class="amount">${this.formatCurrency(formData.line32)}</td>
                </tr>
            </tbody>
        </table>

        <div class="footer">
            Generated on ${new Date().toLocaleDateString()} | Form 1040 (${formData.taxYear})
        </div>
    </div>
</body>
</html>
    `;
  }
}
