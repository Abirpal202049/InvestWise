import { useState } from 'react';
import { Routes, Route, Navigate, useLocation } from 'react-router-dom';
import { Calculator, TrendingUp, ArrowDownUp } from 'lucide-react';
import type { Region } from './types';
import { Tabs, SettingsDropdown } from './components/ui';
import { SIPCalculator, StepUpSIPCalculator, SWPCalculator } from './components/calculator';

const CALCULATOR_TABS = [
  { id: 'sip-calculator', label: 'SIP Calculator', icon: <Calculator className="w-4 h-4" /> },
  { id: 'step-up-sip-calculator', label: 'Step-Up SIP', icon: <TrendingUp className="w-4 h-4" /> },
  { id: 'swp-calculator', label: 'SWP Calculator', icon: <ArrowDownUp className="w-4 h-4" /> },
];

function App() {
  const location = useLocation();
  const [region, setRegion] = useState<Region>('INR');
  const [showInflation, setShowInflation] = useState(true);

  // Get active tab from URL path
  const activeTab = location.pathname.slice(1) || 'sip-calculator';

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100 dark:from-slate-900 dark:to-slate-800 flex flex-col">
      {/* Header */}
      <header className="bg-white dark:bg-slate-800 border-b border-gray-200 dark:border-slate-700 sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex items-center justify-between h-16">
            <div className="flex items-center gap-3">
              <div className="bg-indigo-600 text-white p-2 rounded-lg">
                <Calculator className="w-6 h-6" />
              </div>
              <div>
                <h1 className="text-xl font-bold text-gray-900 dark:text-white">SIP/SWP Calculator</h1>
                <p className="text-xs text-gray-500 dark:text-slate-400 hidden sm:block">
                  Plan your investments smartly
                </p>
              </div>
            </div>
            <SettingsDropdown
              region={region}
              onRegionChange={setRegion}
              showInflation={showInflation}
              onInflationToggle={setShowInflation}
            />
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 w-full">
        {/* Tabs */}
        <div className="bg-white dark:bg-slate-800 rounded-t-xl border border-b-0 border-gray-200 dark:border-slate-700">
          <Tabs
            tabs={CALCULATOR_TABS}
            activeTab={activeTab}
          />
        </div>

        {/* Calculator Content */}
        <div className="bg-white dark:bg-slate-800 rounded-b-xl border border-gray-200 dark:border-slate-700 p-5">
          <Routes>
            <Route path="/" element={<Navigate to="/sip-calculator" replace />} />
            <Route path="/sip-calculator" element={<SIPCalculator region={region} showInflation={showInflation} />} />
            <Route path="/step-up-sip-calculator" element={<StepUpSIPCalculator region={region} showInflation={showInflation} />} />
            <Route path="/swp-calculator" element={<SWPCalculator region={region} showInflation={showInflation} />} />
            <Route path="*" element={<Navigate to="/sip-calculator" replace />} />
          </Routes>
        </div>
      </main>

      {/* Footer */}
      <footer className="bg-white dark:bg-slate-800 border-t border-gray-200 dark:border-slate-700 mt-auto">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
          <p className="text-center text-sm text-gray-500 dark:text-slate-400">
            Disclaimer: The calculations are based on assumed rates of return and are for
            illustrative purposes only. Actual returns may vary based on market conditions.
          </p>
        </div>
      </footer>
    </div>
  );
}

export default App;
