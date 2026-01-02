import { useState, useRef, useEffect } from 'react';
import type { YearlyData, SWPYearlyData, Region } from '../../types';
import { formatCurrency } from '../../utils/formatCurrency';
import { exportSIPData, exportSWPData, exportSIPToGoogleSheets, exportSWPToGoogleSheets } from '../../utils/exportData';
import { CardTitle } from '../ui';
import { Table, Download, FileSpreadsheet, FileText, ChevronDown, ExternalLink } from 'lucide-react';

function GoogleSheetsIcon({ className }: { className?: string }) {
  return (
    <svg
      role="img"
      viewBox="0 0 24 24"
      xmlns="http://www.w3.org/2000/svg"
      className={className}
      fill="#34A853"
    >
      <path d="M11.318 12.545H7.91v-1.909h3.41v1.91zM14.728 0v6h6l-6-6zm1.363 10.636h-3.41v1.91h3.41v-1.91zm0 3.273h-3.41v1.91h3.41v-1.91zM20.727 6.5v15.864c0 .904-.732 1.636-1.636 1.636H4.909a1.636 1.636 0 0 1-1.636-1.636V1.636C3.273.732 4.005 0 4.909 0h9.318v6.5h6.5zm-3.273 2.773H6.545v7.909h10.91v-7.91zm-6.136 4.636H7.91v1.91h3.41v-1.91z" />
    </svg>
  );
}

interface SIPYearlyTableProps {
  data: YearlyData[];
  region: Region;
  showMonthlyAmount?: boolean;
  showInflation?: boolean;
}

interface ExportDropdownProps {
  onExportCSV: () => void;
  onExportExcel: () => void;
  onExportGoogleSheets: () => void;
}

function ExportDropdown({ onExportCSV, onExportExcel, onExportGoogleSheets }: ExportDropdownProps) {
  const [isOpen, setIsOpen] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  return (
    <div className="relative" ref={dropdownRef}>
      <button
        onClick={() => setIsOpen(!isOpen)}
        className="flex items-center gap-1.5 px-3 py-1.5 text-sm font-medium text-gray-700 dark:text-slate-200 bg-white dark:bg-slate-700 border border-gray-300 dark:border-slate-600 rounded-lg hover:bg-gray-50 dark:hover:bg-slate-600 transition-colors"
      >
        <Download className="w-4 h-4" />
        Export
        <ChevronDown className={`w-4 h-4 transition-transform ${isOpen ? 'rotate-180' : ''}`} />
      </button>
      {isOpen && (
        <div className="absolute right-0 mt-1 w-60 bg-white dark:bg-slate-700 border border-gray-200 dark:border-slate-600 rounded-lg shadow-lg z-50">
          <button
            onClick={() => {
              onExportCSV();
              setIsOpen(false);
            }}
            className="flex items-center gap-2 w-full px-3 py-2 text-sm text-gray-700 dark:text-slate-200 hover:bg-gray-50 dark:hover:bg-slate-600 transition-colors rounded-t-lg"
          >
            <FileText className="w-4 h-4 text-green-600 dark:text-green-400" />
            Export as CSV
          </button>
          <button
            onClick={() => {
              onExportExcel();
              setIsOpen(false);
            }}
            className="flex items-center gap-2 w-full px-3 py-2 text-sm text-gray-700 dark:text-slate-200 hover:bg-gray-50 dark:hover:bg-slate-600 transition-colors border-t border-gray-100 dark:border-slate-600"
          >
            <FileSpreadsheet className="w-4 h-4 text-emerald-600 dark:text-emerald-400" />
            Export as Excel
          </button>
          <button
            onClick={() => {
              onExportGoogleSheets();
              setIsOpen(false);
            }}
            className="flex items-center gap-2 w-full px-3 py-2 text-sm text-gray-700 dark:text-slate-200 hover:bg-gray-50 dark:hover:bg-slate-600 transition-colors rounded-b-lg border-t border-gray-100 dark:border-slate-600"
          >
            <GoogleSheetsIcon className="w-4 h-4" />
            Open in Google Sheets
            <ExternalLink className='w-4 h-4' />
          </button>
        </div>
      )}
    </div>
  );
}

export function SIPYearlyTable({ data, region, showMonthlyAmount = false, showInflation = true }: SIPYearlyTableProps) {
  const hasInflationData = data.length > 0 && data[0].inflationAdjustedValue !== undefined;
  const displayInflation = showInflation && hasInflationData;

  return (
    <div className="bg-white dark:bg-slate-800 rounded-xl shadow-sm border border-gray-100 dark:border-slate-700">
      {/* Header with Export - separate div to avoid clipping */}
      <div className="px-6 py-4 border-b border-gray-100 dark:border-slate-700 flex flex-row items-center justify-between">
        <div className="flex items-center gap-2">
          <Table className="w-5 h-5 text-indigo-600 dark:text-indigo-400" />
          <CardTitle>Year-on-Year Breakdown</CardTitle>
        </div>
        <ExportDropdown
          onExportCSV={() => exportSIPData(data, region, 'csv', showMonthlyAmount, displayInflation)}
          onExportExcel={() => exportSIPData(data, region, 'excel', showMonthlyAmount, displayInflation)}
          onExportGoogleSheets={() => exportSIPToGoogleSheets(data, region, showMonthlyAmount, displayInflation)}
        />
      </div>
      {/* Table Content */}
      <div className="table-container max-h-96 overflow-auto">
        <table className="w-full min-w-max">
          <thead>
            <tr className="bg-gray-50 dark:bg-slate-700 border-b border-gray-200 dark:border-slate-600">
              <th className="px-4 py-3 text-left text-xs font-semibold text-gray-600 dark:text-slate-300 uppercase tracking-wider">
                Year
              </th>
              {showMonthlyAmount && (
                <th className="px-4 py-3 text-right text-xs font-semibold text-gray-600 dark:text-slate-300 uppercase tracking-wider">
                  Monthly SIP
                </th>
              )}
              <th className="px-4 py-3 text-right text-xs font-semibold text-gray-600 dark:text-slate-300 uppercase tracking-wider">
                Annual Investment
              </th>
              <th className="px-4 py-3 text-right text-xs font-semibold text-gray-600 dark:text-slate-300 uppercase tracking-wider">
                Cumulative Investment
              </th>
              <th className="px-4 py-3 text-right text-xs font-semibold text-gray-600 dark:text-slate-300 uppercase tracking-wider">
                Value at Year-End
              </th>
              <th className="px-4 py-3 text-right text-xs font-semibold text-gray-600 dark:text-slate-300 uppercase tracking-wider">
                Total Returns
              </th>
              <th
                className={`px-4 py-3 text-right text-xs font-semibold text-amber-600 dark:text-amber-400 uppercase tracking-wider transition-all duration-300 ${
                  displayInflation ? 'opacity-100 max-w-xs' : 'opacity-0 max-w-0 overflow-hidden px-0'
                }`}
              >
                {displayInflation && 'Inflation Adjusted'}
              </th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100 dark:divide-slate-700">
            {data.map((row, index) => (
              <tr
                key={row.year}
                className={`
                  ${index % 2 === 0 ? 'bg-white dark:bg-slate-800' : 'bg-gray-50/50 dark:bg-slate-700/50'}
                  ${row.year % 5 === 0 ? 'bg-indigo-50/50 dark:bg-indigo-900/20' : ''}
                  hover:bg-indigo-50/30 dark:hover:bg-indigo-900/30 transition-colors
                `}
              >
                <td className="px-4 py-3 text-sm font-medium text-gray-900 dark:text-white">
                  {row.year}
                </td>
                {showMonthlyAmount && (
                  <td className="px-4 py-3 text-sm text-right text-gray-700 dark:text-slate-300">
                    {formatCurrency(row.monthlyAmount || 0, region)}
                  </td>
                )}
                <td className="px-4 py-3 text-sm text-right text-gray-700 dark:text-slate-300">
                  {formatCurrency(row.annualInvestment, region)}
                </td>
                <td className="px-4 py-3 text-sm text-right text-blue-600 dark:text-blue-400 font-medium">
                  {formatCurrency(row.cumulativeInvestment, region)}
                </td>
                <td className="px-4 py-3 text-sm text-right text-purple-600 dark:text-purple-400 font-medium">
                  {formatCurrency(row.valueAtYearEnd, region)}
                </td>
                <td className="px-4 py-3 text-sm text-right text-green-600 dark:text-green-400 font-medium">
                  {formatCurrency(row.totalReturns, region)}
                </td>
                <td
                  className={`px-4 py-3 text-sm text-right text-amber-600 dark:text-amber-400 font-medium transition-all duration-300 ${
                    displayInflation ? 'opacity-100 max-w-xs' : 'opacity-0 max-w-0 overflow-hidden px-0'
                  }`}
                >
                  {displayInflation && formatCurrency(row.inflationAdjustedValue || row.valueAtYearEnd, region)}
                </td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr className="bg-gray-100 dark:bg-slate-700 border-t-2 border-gray-200 dark:border-slate-600">
              <td className="px-4 py-3 text-sm font-bold text-gray-900 dark:text-white">
                Total
              </td>
              {showMonthlyAmount && <td />}
              <td className="px-4 py-3 text-sm text-right text-gray-500 dark:text-slate-400">-</td>
              <td className="px-4 py-3 text-sm text-right text-blue-700 dark:text-blue-400 font-bold">
                {data.length > 0
                  ? formatCurrency(data[data.length - 1].cumulativeInvestment, region)
                  : '-'}
              </td>
              <td className="px-4 py-3 text-sm text-right text-purple-700 dark:text-purple-400 font-bold">
                {data.length > 0
                  ? formatCurrency(data[data.length - 1].valueAtYearEnd, region)
                  : '-'}
              </td>
              <td className="px-4 py-3 text-sm text-right text-green-700 dark:text-green-400 font-bold">
                {data.length > 0
                  ? formatCurrency(data[data.length - 1].totalReturns, region)
                  : '-'}
              </td>
              <td
                className={`px-4 py-3 text-sm text-right text-amber-700 dark:text-amber-400 font-bold transition-all duration-300 ${
                  displayInflation ? 'opacity-100 max-w-xs' : 'opacity-0 max-w-0 overflow-hidden px-0'
                }`}
              >
                {displayInflation && (data.length > 0
                  ? formatCurrency(data[data.length - 1].inflationAdjustedValue || data[data.length - 1].valueAtYearEnd, region)
                  : '-')}
              </td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  );
}

interface SWPYearlyTableProps {
  data: SWPYearlyData[];
  region: Region;
  showInflation?: boolean;
}

export function SWPYearlyTable({ data, region, showInflation = true }: SWPYearlyTableProps) {
  const hasInflationData = data.length > 0 && data[0].inflationAdjustedCorpus !== undefined;
  const displayInflation = showInflation && hasInflationData;

  return (
    <div className="bg-white dark:bg-slate-800 rounded-xl shadow-sm border border-gray-100 dark:border-slate-700">
      {/* Header with Export - separate div to avoid clipping */}
      <div className="px-6 py-4 border-b border-gray-100 dark:border-slate-700 flex flex-row items-center justify-between">
        <div className="flex items-center gap-2">
          <Table className="w-5 h-5 text-indigo-600 dark:text-indigo-400" />
          <CardTitle>Year-on-Year Breakdown</CardTitle>
        </div>
        <ExportDropdown
          onExportCSV={() => exportSWPData(data, region, 'csv', displayInflation)}
          onExportExcel={() => exportSWPData(data, region, 'excel', displayInflation)}
          onExportGoogleSheets={() => exportSWPToGoogleSheets(data, region, displayInflation)}
        />
      </div>
      {/* Table Content */}
      <div className="table-container max-h-96 overflow-auto">
        <table className="w-full min-w-max">
          <thead>
            <tr className="bg-gray-50 dark:bg-slate-700 border-b border-gray-200 dark:border-slate-600">
              <th className="px-4 py-3 text-left text-xs font-semibold text-gray-600 dark:text-slate-300 uppercase tracking-wider">
                Year
              </th>
              <th className="px-4 py-3 text-right text-xs font-semibold text-gray-600 dark:text-slate-300 uppercase tracking-wider">
                Annual Withdrawal
              </th>
              <th className="px-4 py-3 text-right text-xs font-semibold text-gray-600 dark:text-slate-300 uppercase tracking-wider">
                Cumulative Withdrawal
              </th>
              <th className="px-4 py-3 text-right text-xs font-semibold text-gray-600 dark:text-slate-300 uppercase tracking-wider">
                Corpus at Year-End
              </th>
              <th className="px-4 py-3 text-right text-xs font-semibold text-gray-600 dark:text-slate-300 uppercase tracking-wider">
                Interest Earned
              </th>
              <th
                className={`px-4 py-3 text-right text-xs font-semibold text-amber-600 dark:text-amber-400 uppercase tracking-wider transition-all duration-300 ${
                  displayInflation ? 'opacity-100 max-w-xs' : 'opacity-0 max-w-0 overflow-hidden px-0'
                }`}
              >
                {displayInflation && "Corpus (Today's Value)"}
              </th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100 dark:divide-slate-700">
            {data.map((row, index) => (
              <tr
                key={row.year}
                className={`
                  ${index % 2 === 0 ? 'bg-white dark:bg-slate-800' : 'bg-gray-50/50 dark:bg-slate-700/50'}
                  ${row.year % 5 === 0 ? 'bg-indigo-50/50 dark:bg-indigo-900/20' : ''}
                  hover:bg-indigo-50/30 dark:hover:bg-indigo-900/30 transition-colors
                `}
              >
                <td className="px-4 py-3 text-sm font-medium text-gray-900 dark:text-white">
                  {row.year}
                </td>
                <td className="px-4 py-3 text-sm text-right text-gray-700 dark:text-slate-300">
                  {formatCurrency(row.annualWithdrawal, region)}
                </td>
                <td className="px-4 py-3 text-sm text-right text-red-600 dark:text-red-400 font-medium">
                  {formatCurrency(row.cumulativeWithdrawal, region)}
                </td>
                <td className="px-4 py-3 text-sm text-right text-purple-600 dark:text-purple-400 font-medium">
                  {formatCurrency(row.corpusAtYearEnd, region)}
                </td>
                <td className="px-4 py-3 text-sm text-right text-green-600 dark:text-green-400 font-medium">
                  {formatCurrency(row.interestEarned, region)}
                </td>
                <td
                  className={`px-4 py-3 text-sm text-right text-amber-600 dark:text-amber-400 font-medium transition-all duration-300 ${
                    displayInflation ? 'opacity-100 max-w-xs' : 'opacity-0 max-w-0 overflow-hidden px-0'
                  }`}
                >
                  {displayInflation && formatCurrency(row.inflationAdjustedCorpus || row.corpusAtYearEnd, region)}
                </td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr className="bg-gray-100 dark:bg-slate-700 border-t-2 border-gray-200 dark:border-slate-600">
              <td className="px-4 py-3 text-sm font-bold text-gray-900 dark:text-white">
                Total
              </td>
              <td className="px-4 py-3 text-sm text-right text-gray-500 dark:text-slate-400">-</td>
              <td className="px-4 py-3 text-sm text-right text-red-700 dark:text-red-400 font-bold">
                {data.length > 0
                  ? formatCurrency(data[data.length - 1].cumulativeWithdrawal, region)
                  : '-'}
              </td>
              <td className="px-4 py-3 text-sm text-right text-purple-700 dark:text-purple-400 font-bold">
                {data.length > 0
                  ? formatCurrency(data[data.length - 1].corpusAtYearEnd, region)
                  : '-'}
              </td>
              <td className="px-4 py-3 text-sm text-right text-green-700 dark:text-green-400 font-bold">
                {data.length > 0
                  ? formatCurrency(
                      data.reduce((sum, row) => sum + row.interestEarned, 0),
                      region
                    )
                  : '-'}
              </td>
              <td
                className={`px-4 py-3 text-sm text-right text-amber-700 dark:text-amber-400 font-bold transition-all duration-300 ${
                  displayInflation ? 'opacity-100 max-w-xs' : 'opacity-0 max-w-0 overflow-hidden px-0'
                }`}
              >
                {displayInflation && (data.length > 0
                  ? formatCurrency(data[data.length - 1].inflationAdjustedCorpus || data[data.length - 1].corpusAtYearEnd, region)
                  : '-')}
              </td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  );
}
