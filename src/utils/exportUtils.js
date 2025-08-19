import ExcelJS from 'exceljs';
import jsPDF from 'jspdf';
import 'jspdf-autotable';

export class AMCExportManager {
  constructor(data, settings, paidQuarters) {
    this.data = data;
    this.settings = settings;
    this.paidQuarters = paidQuarters;
  }

  // Export to Excel with multiple sheets and colors
  async exportToExcel(filename = 'AMC_Schedule') {
    const workbook = new ExcelJS.Workbook();

    // Sheet 1: Payment Status (with colors)
    const paymentData = this.preparePaymentStatus();
    const paymentWS = workbook.addWorksheet('Payment Status');
    if (paymentData.length > 0) {
      paymentWS.addRow(Object.keys(paymentData[0]));
      paymentData.forEach(row => {
        const newRow = paymentWS.addRow(Object.values(row));
        if (row.Status === 'PAID') {
          // Find Status column and Payment Date column indices
          const statusColIndex = Object.keys(paymentData[0]).indexOf('Status') + 1;
          const dateColIndex = Object.keys(paymentData[0]).indexOf('Payment Date') + 1;
          
          // Green for PAID status
          newRow.getCell(statusColIndex).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF52E618' }
          };
          
          // Yellow for Payment Date
          newRow.getCell(dateColIndex).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFF00' }
          };
        }
      });
    }

    // Sheet 2: Quarter Summary
    const quarterData = this.prepareQuarterSummary();
    const quarterWS = workbook.addWorksheet('Quarter Summary');
    if (quarterData.length > 0) {
      quarterWS.addRow(Object.keys(quarterData[0]));
      quarterData.forEach(row => {
        const newRow = quarterWS.addRow(Object.values(row));
        if (row['Payment Status'] === 'PAID') {
          // Find Payment Status column and Payment Date column indices
          const statusColIndex = Object.keys(quarterData[0]).indexOf('Payment Status') + 1;
          const dateColIndex = Object.keys(quarterData[0]).indexOf('Payment Date') + 1;
          
          // Green for PAID status
          newRow.getCell(statusColIndex).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF52E618' }
          };
          
          // Yellow for Payment Date
          newRow.getCell(dateColIndex).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFF00' }
          };
        }
      });
    }

    // Sheet 3: AMC Schedule (main data)
    const scheduleData = this.prepareScheduleData();
    const scheduleWS = workbook.addWorksheet('AMC Schedule');
    if (scheduleData.length > 0) {
      scheduleWS.addRow(Object.keys(scheduleData[0]));
      scheduleData.forEach(row => scheduleWS.addRow(Object.values(row)));
    }
  

    // Download the file
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${filename}_${new Date().toISOString().split('T')[0]}.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  // Export to CSV (flattened schedule data)
  exportToCSV(filename = 'AMC_Schedule') {
    const scheduleData = this.prepareScheduleData();
    const csv = this.convertToCSV(scheduleData);
    this.downloadFile(csv, `${filename}_${new Date().toISOString().split('T')[0]}.csv`, 'text/csv');
  }

  // Export to PDF Report
  exportToPDF(filename = 'AMC_Report') {
    const doc = new jsPDF();
    const pageWidth = doc.internal.pageSize.getWidth();
    
    // Header
    doc.setFontSize(18);
    doc.setTextColor(59, 130, 246);
    doc.text('AMC Payment Report', pageWidth / 2, 20, { align: 'center' });
    
    doc.setFontSize(10);
    doc.setTextColor(100);
    doc.text(`Generated on: ${new Date().toLocaleDateString()}`, pageWidth / 2, 30, { align: 'center' });

    let yPosition = 50;

    // Payment Summary
    const summary = this.calculatePaymentSummary();
    doc.setFontSize(14);
    doc.setTextColor(0);
    doc.text('Payment Summary', 20, yPosition);
    yPosition += 15;

    const summaryData = [
      ['Total Amount', `₹${summary.total.toLocaleString()}`],
      ['Paid Amount', `₹${summary.paid.toLocaleString()}`],
      ['Balance Amount', `₹${summary.balance.toLocaleString()}`],
      ['Quarters Paid', `${summary.paidCount}/${summary.totalCount}`],
    ];

    doc.autoTable({
      startY: yPosition,
      head: [['Metric', 'Value']],
      body: summaryData,
      theme: 'grid',
      headStyles: { fillColor: [59, 130, 246] },
      margin: { left: 20, right: 20 },
    });

    yPosition = doc.lastAutoTable.finalY + 20;

    // Quarter Details
    doc.setFontSize(14);
    doc.text('Quarter-wise Details', 20, yPosition);
    yPosition += 10;

    const quarterData = this.prepareQuarterSummary();
    const quarterTableData = quarterData.map(q => [
      q.Quarter,
      `₹${q['Amount (With GST)'].toLocaleString()}`,
      q['Payment Status'],
      q['Payment Date'] || '-'
    ]);

    doc.autoTable({
      startY: yPosition,
      head: [['Quarter', 'Amount', 'Status', 'Paid Date']],
      body: quarterTableData,
      theme: 'grid',
      headStyles: { fillColor: [59, 130, 246] },
      margin: { left: 20, right: 20 },
    });

    // Save the PDF
    doc.save(`${filename}_${new Date().toISOString().split('T')[0]}.pdf`);
  }

  // Export to JSON
  exportToJSON(filename = 'AMC_Data') {
    const exportData = {
      metadata: {
        exportDate: new Date().toISOString(),
        settings: this.settings,
        totalProducts: this.data.length,
      },
      scheduleData: this.prepareScheduleData(),
      quarterSummary: this.prepareQuarterSummary(),
      paymentStatus: this.preparePaymentStatus(),
      settings: this.prepareSettingsData(),
    };

    const json = JSON.stringify(exportData, null, 2);
    this.downloadFile(json, `${filename}_${new Date().toISOString().split('T')[0]}.json`, 'application/json');
  }

  // Helper methods
  prepareScheduleData() {
    return this.data.map(product => {
      const row = {
        'Product Name': product.productName,
        'Location': product.location,
        'Invoice Value': product.invoiceValue,
        'Quantity': product.quantity,
        'AMC Start Date': product.amcStartDate,
        'UAT Date': product.uatDate,
      };

      // Add quarter columns
      Object.keys(product).forEach(key => {
        if (key.match(/^[A-Z]{3}-\d{4}$/)) {
          row[key] = product[key];
        }
      });

      return row;
    });
  }

  preparePaymentStatus() {
  const quarterSummary = this.prepareQuarterSummary();
  return quarterSummary.map(q => {
    let daysOverdue = 0;
    
    if (q['Payment Status'] === 'PENDING') {
      // Calculate days overdue for pending payments
      const [quarterCode, year] = q.Quarter.split('-');
      const quarterMonths = { JFM: 2, AMJ: 5, JAS: 8, OND: 11 }; // End months of quarters
      const quarterEndDate = new Date(parseInt(year), quarterMonths[quarterCode], 0); // Last day of quarter
      const currentDate = new Date();
      
      if (currentDate > quarterEndDate) {
        daysOverdue = Math.floor((currentDate - quarterEndDate) / (1000 * 60 * 60 * 24));
      }
    }
    
    return {
      Quarter: q.Quarter,
      Amount: q['Amount (With GST)'],
      Status: q['Payment Status'],
      'Payment Date': q['Payment Date'],
      'Days Overdue': daysOverdue,
    };
  });}
  
  prepareQuarterSummary() {
    const quarters = {};
    
    // Extract quarters from data
    this.data.forEach(product => {
      Object.keys(product).forEach(key => {
        if (key.match(/^[A-Z]{3}-\d{4}$/)) {
          if (!quarters[key]) {
            quarters[key] = 0;
          }
          quarters[key] += product[key] || 0;
        }
      });
    });

    return Object.entries(quarters)
      .sort(([a], [b]) => {
        const [qA, yA] = a.split('-');
        const [qB, yB] = b.split('-');
        if (yA !== yB) return parseInt(yA) - parseInt(yB);
        const qOrder = { JFM: 0, AMJ: 1, JAS: 2, OND: 3 };
        return qOrder[qA] - qOrder[qB];
      })
      .map(([quarter, amount]) => ({
        Quarter: quarter,
        'Amount (With GST)': amount,
        'Amount (Without GST)': Math.round(amount / (1 + this.settings.gstRate)),
        'Payment Status': this.paidQuarters[quarter]?.paid ? 'PAID' : 'PENDING',
        'Payment Date': this.paidQuarters[quarter]?.date || '',
      }));
  }

  calculatePaymentSummary() {
    const quarterSummary = this.prepareQuarterSummary();
    const total = quarterSummary.reduce((sum, q) => sum + q['Amount (With GST)'], 0);
    const paid = quarterSummary
      .filter(q => q['Payment Status'] === 'PAID')
      .reduce((sum, q) => sum + q['Amount (With GST)'], 0);
    
    return {
      total,
      paid,
      balance: total - paid,
      paidCount: quarterSummary.filter(q => q['Payment Status'] === 'PAID').length,
      totalCount: quarterSummary.length
    }}}