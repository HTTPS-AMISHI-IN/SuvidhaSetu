import ExcelJS from 'exceljs';
import jsPDF from 'jspdf';
import 'jspdf-autotable';

export class AMCExportManager {
  constructor(data, settings, paidQuarters) {
    this.data = data;
    this.settings = settings;
    this.paidQuarters = paidQuarters;
  }

  async exportToExcel(filename = 'AMC_Schedule') {
  const workbook = new ExcelJS.Workbook();

  // Sheet 1: Payment Status (with colors)
  const paymentData = this.preparePaymentStatus();
  const paymentWS = workbook.addWorksheet('Payment Status');
  if (paymentData.length > 0) {
    paymentWS.addRow(Object.keys(paymentData[0]));
    
    paymentData.forEach(row => {
      const newRow = paymentWS.addRow(Object.values(row));
      const statusColIndexRaw = Object.keys(paymentData[0]).indexOf('Status');
      const dateColIndexRaw = Object.keys(paymentData).indexOf('Payment Date');

      if (row.Status === 'PAID') {
        // Neon green fill for PAID status
        if (statusColIndexRaw !== -1) {
          newRow.getCell(statusColIndexRaw + 1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF00FF00' }
          };
        }
      }
      if (
  dateColIndexRaw !== -1 &&
  typeof row['Payment Date'] === 'string' &&
  row['Payment Date'].trim() !== ''
) {
  newRow.getCell(dateColIndexRaw + 1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFFF00' } // Yellow
  };
}
    });
      
    const totalAmount = paymentData.reduce((sum, row) => sum + (parseFloat(row.Amount) || 0), 0);
    const totalPaid = paymentData
      .filter(row => row.Status === 'PAID')
      .reduce((sum, row) => sum + (parseFloat(row.Amount) || 0), 0);
    const totalPending = paymentData
      .filter(row => row.Status === 'PENDING')
      .reduce((sum, row) => sum + (parseFloat(row.Amount) || 0), 0);

    const totalAmountRow = paymentWS.addRow(['Total Amount', totalAmount]);
    totalAmountRow.font = { bold: true };
    totalAmountRow.getCell(2).numFmt = '"₹"#,##0.00';

    const totalPaidRow = paymentWS.addRow(['Total Paid', totalPaid]);
    totalPaidRow.font = { bold: true };
    totalPaidRow.getCell(2).numFmt = '"₹"#,##0.00';

    const totalPendingRow = paymentWS.addRow(['Total Pending', totalPending]);
    totalPendingRow.font = { bold: true };
    totalPendingRow.getCell(2).numFmt = '"₹"#,##0.00';

    paymentWS.addRow([]);
  }

  // Sheet 2: Quarter Summary
  const quarterData = this.prepareQuarterSummary();
  const quarterWS = workbook.addWorksheet('Quarter Summary');
  if (quarterData.length > 0) {
    quarterWS.addRow(Object.keys(quarterData[0]));
    
    quarterData.forEach(row => {
      const newRow = quarterWS.addRow(Object.values(row));
      const statusColIndexRaw = Object.keys(quarterData[0]).indexOf('Payment Status');
      const dateColIndexRaw = Object.keys(quarterData).indexOf('Payment Date');
      
      if (row['Payment Status'] === 'PAID') {
        if (statusColIndexRaw !== -1) {
          newRow.getCell(statusColIndexRaw + 1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF00FF00' }
          };
        }
      }
      
      if (dateColIndexRaw !== -1 && row['Payment Date']) {
        newRow.getCell(dateColIndexRaw + 1).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFFF00' }
        };
      }
    });

    const totalWithGST = quarterData.reduce((sum, row) => sum + (parseFloat(row['Amount (With GST)']) || 0), 0);
    const totalWithoutGST = quarterData.reduce((sum, row) => sum + (parseFloat(row['Amount (Without GST)']) || 0), 0);
    
    const totalRow = quarterWS.addRow(['TOTAL', totalWithGST, totalWithoutGST]);
    totalRow.font = { bold: true };
    totalRow.getCell(2).numFmt = '"₹"#,##0.00';
    totalRow.getCell(3).numFmt = '"₹"#,##0.00';
    
    quarterWS.addRow([]);
  }

  // Sheet 3: AMC Schedule
  const scheduleData = this.prepareScheduleData();
  const scheduleWS = workbook.addWorksheet('AMC Schedule');
  if (scheduleData.length > 0) {
    const headers = Object.keys(scheduleData[0]);
    scheduleWS.addRow(headers);
    scheduleData.forEach(row => scheduleWS.addRow(Object.values(row)));

    // Calculate totals for all numeric columns including quarters
    const totals = {};
    headers.forEach(header => {
      if (header === 'Invoice Value' || header === 'Quantity' || header.match(/^[A-Z]{3}-\d{4}$/)) {
        totals[header] = scheduleData.reduce((sum, row) => sum + (parseFloat(row[header]) || 0), 0);
      }
    });

    // Create totals row with all calculated totals
    const totalRowData = headers.map(header => {
      if (header === 'Product Name') {
        return 'TOTAL';
      } else if (totals[header] !== undefined) {
        return totals[header];
      } else {
        return '';
      }
    });

    const totalRow = scheduleWS.addRow(totalRowData);
    totalRow.font = { bold: true };

    // Apply currency formatting to numeric columns
    headers.forEach((header, index) => {
      if (header === 'Invoice Value' || header.match(/^[A-Z]{3}-\d{4}$/)) {
        totalRow.getCell(index + 1).numFmt = '"₹"#,##0.00';
      }
    });
    
    scheduleWS.addRow([]);
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
    });
  }
  
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
    };
  }
}

// WarrantyExportManager.js
export class WarrantyExportManager {
  constructor({products = [], schedule = [], paidQuarters = {}, settings = {gstRate: 0.18}}) {
    this.products = products;           // warrantyProducts
    this.schedule = schedule;           // calculatedSchedule (array of rows)
    this.paidQuarters = paidQuarters;   // object { quarterKey: {paid, date} }
    this.settings = settings;
  }

  // Helper: Prepare schedule data for export
  prepareScheduleData() {
    return this.schedule
      .filter(row => row.itemName !== "Grand Total")
      .map(row => {
        const { itemName, location, cost, quantity, uatDate, warrantyStart, ...quarters } = row;
        const base = {
          'Item Name': itemName,
          'Location': location,
          'Cost': cost,
          'Quantity': quantity,
          'UAT Date': uatDate,
          'Warranty Start': warrantyStart,
        };
        // Add quarter values
        Object.keys(quarters).forEach(key => {
          if (key.match(/^(JFM|AMJ|JAS|OND) \d{4}$/)) base[key] = row[key];
        });
        return base;
      });
  }

  // Helper: Prepare quarter summary
  prepareQuarterSummary() {
    const quarters = {};
    this.schedule
      .filter(row => row.itemName !== "Grand Total")
      .forEach(row => {
        Object.keys(row).forEach(key => {
          if (key.match(/^(JFM|AMJ|JAS|OND) \d{4}$/)) {
            if (!quarters[key]) quarters[key] = 0;
            quarters[key] += row[key] || 0;
          }
        });
      });
    return Object.entries(quarters)
      .sort(([a], [b]) => {
        const [qA, yA] = a.split(' ');
        const [qB, yB] = b.split(' ');
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

  async exportToExcel(filename = 'Warranty_Tracking') {
    const workbook = new ExcelJS.Workbook();

    // Sheet 1: Payment Status
    const paymentData = this.prepareQuarterSummary();
    const paymentWS = workbook.addWorksheet('Payment Status');
    if (paymentData.length > 0) {
      paymentWS.addRow(Object.keys(paymentData[0]));
      paymentData.forEach(row => {
        const newRow = paymentWS.addRow(Object.values(row));
        const statusColIndex = Object.keys(paymentData[0]).indexOf('Payment Status'); // FIXED
        const dateColIndex = Object.keys(paymentData[0]).indexOf('Payment Date'); // FIXED
        if (row['Payment Status'] === 'PAID' && statusColIndex !== -1) {
          newRow.getCell(statusColIndex + 1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF00FF00' },
          };
        }
        if (dateColIndex !== -1 && row['Payment Date'] && row['Payment Date'].trim() !== '') {
          newRow.getCell(dateColIndex + 1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFF00' } // Yellow
          };
        }
      });
      
      // Summary rows
      const totalAmount = paymentData.reduce((sum, row) => sum + (parseFloat(row['Amount (With GST)']) || 0), 0);
      const totalWithoutGST = paymentData.reduce((sum, row) => sum + (parseFloat(row['Amount (Without GST)']) || 0), 0);
      const totalPaid = paymentData.filter(row => row['Payment Status'] === 'PAID')
        .reduce((sum, row) => sum + (parseFloat(row['Amount (With GST)']) || 0), 0);
      const totalPending = paymentData.filter(row => row['Payment Status'] === 'PENDING')
        .reduce((sum, row) => sum + (parseFloat(row['Amount (With GST)']) || 0), 0);
      paymentWS.addRow(['Total Amount', totalAmount, totalWithoutGST]).font = { bold: true };
      paymentWS.addRow(['Total Paid', totalPaid]).font = { bold: true };
      paymentWS.addRow(['Total Pending', totalPending]).font = { bold: true };
      paymentWS.addRow([]);
    }

    // Sheet 2: Warranty Schedule (quarter-wise)
    const scheduleData = this.prepareScheduleData();
    const scheduleWS = workbook.addWorksheet('Warranty Schedule');
    if (scheduleData.length > 0) {
      scheduleWS.addRow(Object.keys(scheduleData[0]));
      scheduleData.forEach(row => scheduleWS.addRow(Object.values(row)));
      
      // Total row - FIXED: Use correct headers extraction and exclude first row
      const headers = Object.keys(scheduleData[0]); // FIXED: was Object.keys(scheduleData)
      const dataForTotals = scheduleData.slice(1); // Skip first row from totals
      
      const totals = {};
    headers.forEach(header => {
      if (header === 'Cost' || header === 'Quantity' || header.match(/^(JFM|AMJ|JAS|OND) \d{4}$/)) {
        totals[header] = dataForTotals.reduce((sum, row) => sum + (parseFloat(row[header]) || 0), 0);
      }
    });
      const totalRowData = headers.map(header => {
        if (header === 'Item Name') return 'TOTAL';
        if (totals[header] !== undefined) return totals[header];
        return '';
      });
      
      const totalRow = scheduleWS.addRow(totalRowData);
      totalRow.font = { bold: true };
      
      // Format currency columns
      headers.forEach((header, index) => {
        if (header.match(/^(JFM|AMJ|JAS|OND) \d{4}$/)) {
          totalRow.getCell(index + 1).numFmt = '"₹"#,##0.00';
        }
      });
      
      scheduleWS.addRow([]);
    }

    // Download logic
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${filename}_${new Date().toISOString().split('T')[0]}.xlsx`; // FIXED: added 
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  // CSV export (quarterly schedule)
  exportToCSV(filename = 'Warranty_Tracking') {
    const scheduleData = this.prepareScheduleData();
    const keys = Object.keys(scheduleData[0]); // FIXED: was Object.keys(scheduleData)
    const csvRows = [
      keys.join(','),
      ...scheduleData.map(row =>
        keys.map(k => `"${row[k] !== undefined ? row[k] : ''}"`).join(','))
    ];
    const csv = csvRows.join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${filename}_${new Date().toISOString().split('T')[0]}.csv`; // FIXED: added 
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  // PDF report export
  exportToPDF(filename = 'Warranty_Tracking_Report') { // FIXED: removed double underscore
    const doc = new jsPDF();
    doc.setFontSize(16);
    doc.text('Warranty Payment Report', 105, 20, { align: 'center' });
    doc.setFontSize(10);
    doc.text(`Generated on: ${new Date().toLocaleDateString()}`, 105, 30, { align: 'center' });

    // Final Payment Summary
    const paymentSummary = (() => {
      const paymentData = this.prepareQuarterSummary();
      const total = paymentData.reduce((sum, row) => sum + (parseFloat(row['Amount (With GST)']) || 0), 0);
      const paid = paymentData.filter(row => row['Payment Status'] === 'PAID')
        .reduce((sum, row) => sum + (parseFloat(row['Amount (With GST)']) || 0), 0);
      const pending = paymentData.filter(row => row['Payment Status'] === 'PENDING')
        .reduce((sum, row) => sum + (parseFloat(row['Amount (With GST)']) || 0), 0);
      return [
        ['Total', `₹${total.toLocaleString()}`],
        ['Paid', `₹${paid.toLocaleString()}`],
        ['Balance', `₹${pending.toLocaleString()}`]
      ];
    })();

    doc.autoTable({
      startY: 40,
      head: [['Metric', 'Amount']],
      body: paymentSummary,
      theme: 'grid',
      headStyles: { fillColor: [59,130,246] },
      margin: { left: 20, right: 20 }
    });

    // Quarter Status Table
    const quarterData = this.prepareQuarterSummary();
    doc.autoTable({
      startY: doc.lastAutoTable.finalY + 10,
      head: [['Quarter', 'Amount (With GST)', 'Status', 'Payment Date']],
      body: quarterData.map(q => [
        q.Quarter,
        `₹${q['Amount (With GST)'].toLocaleString()}`,
        q['Payment Status'],
        q['Payment Date'] || '-'
      ]),
      theme: 'grid',
      headStyles: { fillColor: [59,130,246] },
      margin: { left: 20, right: 20 }
    });

    doc.save(`${filename}_${new Date().toISOString().split('T')[0]}.pdf`); // FIXED: added 
  }

  // JSON export
  exportToJSON(filename = 'Warranty_Tracking') {
    const json = JSON.stringify({
      exportDate: new Date().toISOString(),
      products: this.products,
      schedule: this.schedule,
      quarters: this.prepareQuarterSummary(),
    }, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${filename}_${new Date().toISOString().split('T')[0]}.json`; // FIXED: added 
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }
}
