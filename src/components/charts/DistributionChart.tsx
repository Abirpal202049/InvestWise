import { useState, useEffect } from 'react';
import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip } from 'recharts';
import type { Region } from '../../types';
import { formatCurrency } from '../../utils/formatCurrency';
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

interface DistributionChartProps {
  investment: number;
  returns: number;
  region: Region;
}

interface TooltipPayload {
  name: string;
  value: number;
  payload: {
    color: string;
  };
}

interface CustomTooltipProps {
  active?: boolean;
  payload?: TooltipPayload[];
  region: Region;
  isDark: boolean;
}

function CustomTooltip({ active, payload, region, isDark }: CustomTooltipProps) {
  if (!active || !payload || !payload.length) return null;

  const data = payload[0];
  return (
    <div className={`p-3 rounded-lg shadow-lg border ${isDark ? 'bg-slate-800 border-slate-700' : 'bg-white border-gray-100'}`}>
      <p className={`font-medium ${isDark ? 'text-white' : 'text-gray-900'}`}>{data.name}</p>
      <p className="text-sm" style={{ color: data.payload.color }}>
        {formatCurrency(data.value, region)}
      </p>
    </div>
  );
}

interface CustomLegendProps {
  investment: number;
  returns: number;
  region: Region;
}

function CustomLegend({ investment, returns, region }: CustomLegendProps) {
  const total = investment + returns;
  const investmentPercent = ((investment / total) * 100).toFixed(1);
  const returnsPercent = ((returns / total) * 100).toFixed(1);

  return (
    <div className="flex flex-col gap-2 md:gap-3 mt-3 md:mt-4">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-1.5 md:gap-2">
          <div
            className="w-2.5 h-2.5 md:w-3 md:h-3 rounded-full"
            style={{ backgroundColor: CHART_COLORS.investment }}
          />
          <span className="text-xs md:text-sm text-gray-600 dark:text-slate-400">Investment</span>
        </div>
        <div className="text-right">
          <span className="text-xs md:text-sm font-medium text-gray-900 dark:text-white">
            {formatCurrency(investment, region)}
          </span>
          <span className="text-[10px] md:text-xs text-gray-500 dark:text-slate-500 ml-1 md:ml-2">({investmentPercent}%)</span>
        </div>
      </div>
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-1.5 md:gap-2">
          <div
            className="w-2.5 h-2.5 md:w-3 md:h-3 rounded-full"
            style={{ backgroundColor: CHART_COLORS.returns }}
          />
          <span className="text-xs md:text-sm text-gray-600 dark:text-slate-400">Returns</span>
        </div>
        <div className="text-right">
          <span className="text-xs md:text-sm font-medium text-gray-900 dark:text-white">
            {formatCurrency(returns, region)}
          </span>
          <span className="text-[10px] md:text-xs text-gray-500 dark:text-slate-500 ml-1 md:ml-2">({returnsPercent}%)</span>
        </div>
      </div>
    </div>
  );
}

export function DistributionChart({ investment, returns, region }: DistributionChartProps) {
  const { isDark } = useTheme();
  const isMobile = useIsMobile();
  const data = [
    { name: 'Total Investment', value: investment, color: CHART_COLORS.investment },
    { name: 'Expected Returns', value: returns, color: CHART_COLORS.returns },
  ];

  // Responsive pie dimensions
  const innerRadius = isMobile ? 45 : 60;
  const outerRadius = isMobile ? 70 : 90;

  return (
    <div className="w-full">
      <div className="h-48 md:h-64 outline-none" tabIndex={-1}>
        <ResponsiveContainer width="100%" height="100%">
          <PieChart>
            <Pie
              data={data}
              cx="50%"
              cy="50%"
              innerRadius={innerRadius}
              outerRadius={outerRadius}
              paddingAngle={2}
              dataKey="value"
            >
              {data.map((entry, index) => (
                <Cell key={`cell-${index}`} fill={entry.color} />
              ))}
            </Pie>
            <Tooltip content={<CustomTooltip region={region} isDark={isDark} />} />
          </PieChart>
        </ResponsiveContainer>
      </div>
      <CustomLegend
        investment={investment}
        returns={returns}
        region={region}
      />
    </div>
  );
}
