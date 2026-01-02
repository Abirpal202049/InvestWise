import type { YearlyData, SWPYearlyData, Region } from '../types';
import { formatCurrency } from './formatCurrency';

/**
 * Convert data to CSV format
 */
function toCSV(headers: string[], rows: string[][]): string {
  const csvContent = [
    headers.join(','),
    ...rows.map((row) => row.map((cell) => `"${cell}"`).join(',')),
  ].join('\n');
  return csvContent;
}

/**
 * Trigger file download
 */
function downloadFile(content: string, filename: string, mimeType: string): void {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/**
 * Open data in Google Sheets in a new tab
 * Creates a new Google Sheet with the data pre-populated
 */
function openInGoogleSheets(headers: string[], rows: string[][], title: string): void {
  // Build tab-separated content (pastes cleanly into Sheets)
  const tsvContent = [
    headers.join('\t'),
    ...rows.map((row) => row.map((cell) => cell.replace(/\t/g, ' ')).join('\t')),
  ].join('\n');

  // Open Google Sheets with a title
  window.open(
    `https://docs.google.com/spreadsheets/create?title=${encodeURIComponent(title)}`,
    '_blank'
  );

  // Copy to clipboard for easy paste
  navigator.clipboard.writeText(tsvContent).then(() => {
    // Show a toast notification
    const notification = document.createElement('div');
    notification.innerHTML = `
      <div style="position:fixed;bottom:20px;right:20px;background:#4CAF50;color:white;padding:16px 24px;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.15);z-index:10000;font-family:system-ui;font-size:14px;display:flex;align-items:center;gap:12px;">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <polyline points="20 6 9 17 4 12"></polyline>
        </svg>
        <span>Data copied! Press <strong>Ctrl+V</strong> (or <strong>Cmd+V</strong>) to paste</span>
      </div>
    `;
    document.body.appendChild(notification);
    setTimeout(() => notification.remove(), 5000);
  });
}

/**
 * Export SIP yearly data
 */
export function exportSIPData(
  data: YearlyData[],
  region: Region,
  format: 'csv' | 'excel',
  showMonthlyAmount: boolean = false,
  showInflation: boolean = false
): void {
  const headers = [
    'Year',
    ...(showMonthlyAmount ? ['Monthly SIP'] : []),
    'Annual Investment',
    'Cumulative Investment',
    'Value at Year-End',
    'Total Returns',
    ...(showInflation ? ['Inflation Adjusted Value'] : []),
  ];

  const rows = data.map((row) => [
    row.year.toString(),
    ...(showMonthlyAmount ? [formatCurrency(row.monthlyAmount || 0, region)] : []),
    formatCurrency(row.annualInvestment, region),
    formatCurrency(row.cumulativeInvestment, region),
    formatCurrency(row.valueAtYearEnd, region),
    formatCurrency(row.totalReturns, region),
    ...(showInflation ? [formatCurrency(row.inflationAdjustedValue || row.valueAtYearEnd, region)] : []),
  ]);

  // Add summary row
  if (data.length > 0) {
    const lastRow = data[data.length - 1];
    rows.push([
      'Total',
      ...(showMonthlyAmount ? ['-'] : []),
      '-',
      formatCurrency(lastRow.cumulativeInvestment, region),
      formatCurrency(lastRow.valueAtYearEnd, region),
      formatCurrency(lastRow.totalReturns, region),
      ...(showInflation ? [formatCurrency(lastRow.inflationAdjustedValue || lastRow.valueAtYearEnd, region)] : []),
    ]);
  }

  const csvContent = toCSV(headers, rows);
  const filename = `sip-breakdown-${new Date().toISOString().split('T')[0]}`;

  if (format === 'csv') {
    downloadFile(csvContent, `${filename}.csv`, 'text/csv;charset=utf-8;');
  } else {
    // For Excel, we use CSV with .xls extension (Excel can open CSV files)
    downloadFile(csvContent, `${filename}.xls`, 'application/vnd.ms-excel;charset=utf-8;');
  }
}

/**
 * Export SIP data to Google Sheets
 */
export function exportSIPToGoogleSheets(
  data: YearlyData[],
  region: Region,
  showMonthlyAmount: boolean = false,
  showInflation: boolean = false
): void {
  const headers = [
    'Year',
    ...(showMonthlyAmount ? ['Monthly SIP'] : []),
    'Annual Investment',
    'Cumulative Investment',
    'Value at Year-End',
    'Total Returns',
    ...(showInflation ? ['Inflation Adjusted Value'] : []),
  ];

  const rows = data.map((row) => [
    row.year.toString(),
    ...(showMonthlyAmount ? [formatCurrency(row.monthlyAmount || 0, region)] : []),
    formatCurrency(row.annualInvestment, region),
    formatCurrency(row.cumulativeInvestment, region),
    formatCurrency(row.valueAtYearEnd, region),
    formatCurrency(row.totalReturns, region),
    ...(showInflation ? [formatCurrency(row.inflationAdjustedValue || row.valueAtYearEnd, region)] : []),
  ]);

  // Add summary row
  if (data.length > 0) {
    const lastRow = data[data.length - 1];
    rows.push([
      'Total',
      ...(showMonthlyAmount ? ['-'] : []),
      '-',
      formatCurrency(lastRow.cumulativeInvestment, region),
      formatCurrency(lastRow.valueAtYearEnd, region),
      formatCurrency(lastRow.totalReturns, region),
      ...(showInflation ? [formatCurrency(lastRow.inflationAdjustedValue || lastRow.valueAtYearEnd, region)] : []),
    ]);
  }

  const title = `SIP Breakdown - ${new Date().toISOString().split('T')[0]}`;
  openInGoogleSheets(headers, rows, title);
}

/**
 * Export SWP yearly data
 */
export function exportSWPData(
  data: SWPYearlyData[],
  region: Region,
  format: 'csv' | 'excel',
  showInflation: boolean = false
): void {
  const headers = [
    'Year',
    'Annual Withdrawal',
    'Cumulative Withdrawal',
    'Corpus at Year-End',
    'Interest Earned',
    ...(showInflation ? ["Corpus (Today's Value)"] : []),
  ];

  const rows = data.map((row) => [
    row.year.toString(),
    formatCurrency(row.annualWithdrawal, region),
    formatCurrency(row.cumulativeWithdrawal, region),
    formatCurrency(row.corpusAtYearEnd, region),
    formatCurrency(row.interestEarned, region),
    ...(showInflation ? [formatCurrency(row.inflationAdjustedCorpus || row.corpusAtYearEnd, region)] : []),
  ]);

  // Add summary row
  if (data.length > 0) {
    const lastRow = data[data.length - 1];
    const totalInterest = data.reduce((sum, row) => sum + row.interestEarned, 0);
    rows.push([
      'Total',
      '-',
      formatCurrency(lastRow.cumulativeWithdrawal, region),
      formatCurrency(lastRow.corpusAtYearEnd, region),
      formatCurrency(totalInterest, region),
      ...(showInflation ? [formatCurrency(lastRow.inflationAdjustedCorpus || lastRow.corpusAtYearEnd, region)] : []),
    ]);
  }

  const csvContent = toCSV(headers, rows);
  const filename = `swp-breakdown-${new Date().toISOString().split('T')[0]}`;

  if (format === 'csv') {
    downloadFile(csvContent, `${filename}.csv`, 'text/csv;charset=utf-8;');
  } else {
    downloadFile(csvContent, `${filename}.xls`, 'application/vnd.ms-excel;charset=utf-8;');
  }
}

/**
 * Export SWP data to Google Sheets
 */
export function exportSWPToGoogleSheets(
  data: SWPYearlyData[],
  region: Region,
  showInflation: boolean = false
): void {
  const headers = [
    'Year',
    'Annual Withdrawal',
    'Cumulative Withdrawal',
    'Corpus at Year-End',
    'Interest Earned',
    ...(showInflation ? ["Corpus (Today's Value)"] : []),
  ];

  const rows = data.map((row) => [
    row.year.toString(),
    formatCurrency(row.annualWithdrawal, region),
    formatCurrency(row.cumulativeWithdrawal, region),
    formatCurrency(row.corpusAtYearEnd, region),
    formatCurrency(row.interestEarned, region),
    ...(showInflation ? [formatCurrency(row.inflationAdjustedCorpus || row.corpusAtYearEnd, region)] : []),
  ]);

  // Add summary row
  if (data.length > 0) {
    const lastRow = data[data.length - 1];
    const totalInterest = data.reduce((sum, row) => sum + row.interestEarned, 0);
    rows.push([
      'Total',
      '-',
      formatCurrency(lastRow.cumulativeWithdrawal, region),
      formatCurrency(lastRow.corpusAtYearEnd, region),
      formatCurrency(totalInterest, region),
      ...(showInflation ? [formatCurrency(lastRow.inflationAdjustedCorpus || lastRow.corpusAtYearEnd, region)] : []),
    ]);
  }

  const title = `SWP Breakdown - ${new Date().toISOString().split('T')[0]}`;
  openInGoogleSheets(headers, rows, title);
}
