import { useState } from 'react';
import { TrendingUp, PieChart, Maximize2, X } from 'lucide-react';
import type { ChartDataPoint, Region } from '../../types';
import { GrowthChart } from './GrowthChart';
import { DistributionChart } from './DistributionChart';

interface ChartTabsProps {
  chartData: ChartDataPoint[];
  investment: number;
  returns: number;
  region: Region;
  growthTitle?: string;
  distributionTitle?: string;
  showInflation?: boolean;
  isSWP?: boolean;
}

type ChartTab = 'growth' | 'distribution';

export function ChartTabs({
  chartData,
  investment,
  returns,
  region,
  growthTitle = 'Investment Growth Over Time',
  distributionTitle = 'Investment vs Returns',
  showInflation = false,
  isSWP = false,
}: ChartTabsProps) {
  const [activeTab, setActiveTab] = useState<ChartTab>('growth');
  const [isExpanded, setIsExpanded] = useState(false);

  const handleExpand = () => {
    setIsExpanded(true);
    document.body.style.overflow = 'hidden';
  };

  const handleClose = () => {
    setIsExpanded(false);
    document.body.style.overflow = '';
  };

  return (
    <>
      <div className="bg-white dark:bg-slate-800 rounded-xl shadow-sm border border-gray-100 dark:border-slate-700">
        {/* Tab Header */}
        <div className="flex items-center justify-between border-b border-gray-100 dark:border-slate-700 px-4">
          <div className="flex">
            <button
              onClick={() => setActiveTab('growth')}
              className={`flex items-center gap-2 px-4 py-3 text-sm font-medium border-b-2 transition-colors duration-200 ${
                activeTab === 'growth'
                  ? 'border-indigo-500 text-indigo-600 dark:text-indigo-400'
                  : 'border-transparent text-gray-500 dark:text-slate-400 hover:text-gray-700 dark:hover:text-slate-200'
              }`}
            >
              <TrendingUp className="w-4 h-4" />
              Growth Chart
            </button>
            <button
              onClick={() => setActiveTab('distribution')}
              className={`flex items-center gap-2 px-4 py-3 text-sm font-medium border-b-2 transition-colors duration-200 ${
                activeTab === 'distribution'
                  ? 'border-indigo-500 text-indigo-600 dark:text-indigo-400'
                  : 'border-transparent text-gray-500 dark:text-slate-400 hover:text-gray-700 dark:hover:text-slate-200'
              }`}
            >
              <PieChart className="w-4 h-4" />
              Distribution
            </button>
          </div>
          {activeTab === 'growth' && (
            <button
              onClick={handleExpand}
              className="p-2 text-gray-500 dark:text-slate-400 hover:text-indigo-600 dark:hover:text-indigo-400 hover:bg-indigo-50 dark:hover:bg-slate-700 rounded-lg transition-colors"
              title="Expand to fullscreen"
            >
              <Maximize2 className="w-4 h-4" />
            </button>
          )}
        </div>

        {/* Tab Content */}
        <div className="p-5">
          {activeTab === 'growth' && (
            <div>
              <h3 className="text-sm font-semibold text-gray-700 dark:text-slate-300 mb-4">
                {growthTitle}
              </h3>
              <GrowthChart data={chartData} region={region} showInflation={showInflation} isSWP={isSWP} />
            </div>
          )}
          {activeTab === 'distribution' && (
            <div>
              <h3 className="text-sm font-semibold text-gray-700 dark:text-slate-300 mb-4">
                {distributionTitle}
              </h3>
              <DistributionChart
                investment={investment}
                returns={returns}
                region={region}
              />
            </div>
          )}
        </div>
      </div>

      {/* Fullscreen Modal */}
      {isExpanded && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-6 bg-black/30 dark:bg-black/50 backdrop-blur-sm">
          {/* Modal Container */}
          <div className="w-[95vw] h-full max-h-[90vh] bg-white dark:bg-slate-900/80 backdrop-blur-xl rounded-xl shadow-2xl border border-white/40 dark:border-slate-700/40 flex flex-col overflow-hidden">
            {/* Header */}
            <div className="flex items-center justify-between px-6 py-4 border-b border-gray-200/50 dark:border-slate-700/50">
              <h2 className="text-lg font-semibold text-gray-900 dark:text-white">{growthTitle}</h2>
              <button
                onClick={handleClose}
                className="p-2 text-gray-500 dark:text-slate-400 hover:text-gray-700 dark:hover:text-white hover:bg-gray-100/50 dark:hover:bg-slate-700/50 rounded-lg transition-colors"
                title="Close"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            {/* Chart Container */}
            <div className="flex-1 p-6 overflow-hidden">
              <div className="w-full h-full">
                <GrowthChart data={chartData} region={region} fullHeight showInflation={showInflation} isSWP={isSWP} />
              </div>
            </div>
          </div>
        </div>
      )}
    </>
  );
}
