import { useState, useEffect } from 'react';
import {
  AreaChart,
  Area,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  Legend,
} from 'recharts';
import type { ChartDataPoint, Region } from '../../types';
import { formatCurrencyCompact, formatCurrency } from '../../utils/formatCurrency';
import { CHART_COLORS } from '../../utils/constants';
import { useTheme } from '../../contexts/ThemeContext';

// Hook to detect mobile viewport
function useIsMobile(breakpoint = 768) {
  const [isMobile, setIsMobile] = useState(false);

  useEffect(() => {
    const checkMobile = () => setIsMobile(window.innerWidth < breakpoint);
    checkMobile();
    window.addEventListener('resize', checkMobile);
    return () => window.removeEventListener('resize', checkMobile);
  }, [breakpoint]);

  return isMobile;
}

interface GrowthChartProps {
  data: ChartDataPoint[];
  region: Region;
  fullHeight?: boolean;
  showInflation?: boolean;
  isSWP?: boolean;
}

interface TooltipPayload {
  color: string;
  name: string;
  value: number;
  dataKey: string;
}

interface CustomTooltipProps {
  active?: boolean;
  payload?: TooltipPayload[];
  label?: number;
  region: Region;
  isDark: boolean;
  showInflation?: boolean;
  isSWP?: boolean;
}

function CustomTooltip({ active, payload, label, region, isDark, showInflation, isSWP }: CustomTooltipProps) {
  if (!active || !payload || !payload.length) return null;

  const investmentData = payload.find((p) => p.dataKey === 'investment');
  const valueData = payload.find((p) => p.dataKey === 'value');
  const inflationData = payload.find((p) => p.dataKey === 'inflationAdjustedValue');

  const investment = investmentData?.value || 0;
  const value = valueData?.value || 0;
  const inflationAdjustedValue = inflationData?.value || 0;
  const returns = value - investment;

  return (
    <div className={`p-3 md:p-4 rounded-lg shadow-lg border ${isDark ? 'bg-slate-800 border-slate-700' : 'bg-white border-gray-100'}`}>
      <p className={`font-semibold mb-1.5 md:mb-2 text-sm md:text-base ${isDark ? 'text-white' : 'text-gray-900'}`}>Year {label}</p>
      <div className="space-y-0.5 md:space-y-1 text-xs md:text-sm">
        <div className="flex justify-between gap-4">
          <span className={isDark ? 'text-slate-400' : 'text-gray-600'}>
            {isSWP ? 'Withdrawn:' : 'Investment:'}
          </span>
          <span className="font-medium text-blue-500">
            {formatCurrency(investment, region)}
          </span>
        </div>
        {!isSWP && (
          <div className="flex justify-between gap-4">
            <span className={isDark ? 'text-slate-400' : 'text-gray-600'}>Returns:</span>
            <span className="font-medium text-green-500">
              {formatCurrency(returns, region)}
            </span>
          </div>
        )}
        <div className={`flex justify-between gap-4 pt-1 border-t ${isDark ? 'border-slate-700' : 'border-gray-100'}`}>
          <span className={isDark ? 'text-slate-400' : 'text-gray-600'}>
            {isSWP ? 'Remaining Corpus:' : 'Total Value:'}
          </span>
          <span className="font-medium text-green-500">
            {formatCurrency(value, region)}
          </span>
        </div>
        {showInflation && inflationAdjustedValue > 0 && (
          <div className="flex justify-between gap-4">
            <span className={isDark ? 'text-slate-400' : 'text-gray-600'}>Today's Value:</span>
            <span className="font-medium text-amber-500">
              {formatCurrency(inflationAdjustedValue, region)}
            </span>
          </div>
        )}
      </div>
    </div>
  );
}

export function GrowthChart({ data, region, fullHeight = false, showInflation = false, isSWP = false }: GrowthChartProps) {
  const { isDark } = useTheme();
  const isMobile = useIsMobile();
  const formatYAxis = (value: number) => formatCurrencyCompact(value, region);

  const gridColor = isDark ? '#334155' : '#e5e7eb';
  const tickColor = isDark ? '#94a3b8' : '#6b7280';

  // Check if we have inflation data
  const hasInflationData = data.length > 0 && data.some(d => d.inflationAdjustedValue !== undefined && d.inflationAdjustedValue > 0);
  const displayInflation = showInflation && hasInflationData;

  return (
    <div className={`w-full ${fullHeight ? 'h-full' : 'h-64 md:h-80'} outline-none`} tabIndex={-1}>
      <ResponsiveContainer width="100%" height="100%">
        <AreaChart
          data={data}
          margin={{ top: 5, right: isMobile ? 5 : 5, left: isMobile ? 5 : -10, bottom: 0 }}
        >
          <defs>
            <linearGradient id="colorInvestment" x1="0" y1="0" x2="0" y2="1">
              <stop offset="5%" stopColor={CHART_COLORS.investment} stopOpacity={0.8} />
              <stop offset="95%" stopColor={CHART_COLORS.investment} stopOpacity={0.2} />
            </linearGradient>
            <linearGradient id="colorValue" x1="0" y1="0" x2="0" y2="1">
              <stop offset="5%" stopColor={CHART_COLORS.returns} stopOpacity={0.8} />
              <stop offset="95%" stopColor={CHART_COLORS.returns} stopOpacity={0.2} />
            </linearGradient>
            <linearGradient id="colorInflation" x1="0" y1="0" x2="0" y2="1">
              <stop offset="5%" stopColor={CHART_COLORS.inflationAdjusted} stopOpacity={0.8} />
              <stop offset="95%" stopColor={CHART_COLORS.inflationAdjusted} stopOpacity={0.2} />
            </linearGradient>
          </defs>
          <CartesianGrid strokeDasharray="3 3" stroke={gridColor} />
          <XAxis
            dataKey="year"
            axisLine={false}
            tickLine={false}
            tick={{ fill: tickColor, fontSize: 10 }}
            tickFormatter={(value) => `Y${value}`}
            interval="preserveStartEnd"
          />
          {!isMobile && (
            <YAxis
              axisLine={false}
              tickLine={false}
              tick={{ fill: tickColor, fontSize: 10 }}
              tickFormatter={formatYAxis}
              width={55}
            />
          )}
          <Tooltip content={<CustomTooltip region={region} isDark={isDark} showInflation={displayInflation} isSWP={isSWP} />} />
          <Legend
            verticalAlign="top"
            height={32}
            wrapperStyle={{ fontSize: '10px' }}
            formatter={(value) => {
              const labels: Record<string, string> = {
                investment: isSWP ? 'Withdrawn' : 'Investment',
                value: isSWP ? 'Corpus' : 'Value',
                inflationAdjustedValue: "Today's",
              };
              return (
                <span className={`text-xs md:text-sm ${isDark ? 'text-slate-400' : 'text-gray-600'}`}>
                  {labels[value] || value}
                </span>
              );
            }}
          />
          <Area
            type="monotone"
            dataKey="investment"
            stackId="1"
            stroke={CHART_COLORS.investment}
            fill="url(#colorInvestment)"
            name="investment"
          />
          <Area
            type="monotone"
            dataKey="value"
            stackId="2"
            stroke={CHART_COLORS.returns}
            fill="url(#colorValue)"
            name="value"
          />
          {displayInflation && (
            <Area
              type="monotone"
              dataKey="inflationAdjustedValue"
              stackId="3"
              stroke={CHART_COLORS.inflationAdjusted}
              fill="url(#colorInflation)"
              name="inflationAdjustedValue"
              strokeDasharray="5 5"
            />
          )}
        </AreaChart>
      </ResponsiveContainer>
    </div>
  );
}
