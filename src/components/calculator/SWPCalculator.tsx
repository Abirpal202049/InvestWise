import { useState, useMemo, useCallback, useEffect } from 'react';
import { Wallet, TrendingUp, Landmark, RotateCcw, TrendingDown } from 'lucide-react';
import type { Region } from '../../types';
import { calculateSWP, toSWPChartData } from '../../utils/calculations';
import { formatCurrency } from '../../utils/formatCurrency';
import { DEFAULT_VALUES, INPUT_RANGES } from '../../utils/constants';
import { Card, CardContent, SliderInput, SummaryCard, Button } from '../ui';
import { ChartTabs } from '../charts';
import { SWPYearlyTable } from './YearlyTable';

interface SWPCalculatorProps {
  region: Region;
  showInflation: boolean;
}

export function SWPCalculator({ region, showInflation }: SWPCalculatorProps) {
  const [initialCorpus, setInitialCorpus] = useState(DEFAULT_VALUES.initialCorpus[region]);
  const [monthlyWithdrawal, setMonthlyWithdrawal] = useState(
    DEFAULT_VALUES.monthlyWithdrawal[region]
  );
  const [expectedReturn, setExpectedReturn] = useState(DEFAULT_VALUES.expectedReturn);
  const [tenure, setTenure] = useState(DEFAULT_VALUES.tenure);
  const [inflationRate, setInflationRate] = useState(DEFAULT_VALUES.inflationRate[region]);

  // Update values when region changes
  useEffect(() => {
    setInitialCorpus(DEFAULT_VALUES.initialCorpus[region]);
    setMonthlyWithdrawal(DEFAULT_VALUES.monthlyWithdrawal[region]);
    setInflationRate(DEFAULT_VALUES.inflationRate[region]);
  }, [region]);

  const result = useMemo(
    () =>
      calculateSWP({
        initialCorpus,
        monthlyWithdrawal,
        expectedReturn,
        tenure,
        region,
      }, inflationRate),
    [initialCorpus, monthlyWithdrawal, expectedReturn, tenure, region, inflationRate]
  );

  const chartData = useMemo(
    () => toSWPChartData(result.yearlyData, initialCorpus),
    [result.yearlyData, initialCorpus]
  );

  const handleReset = useCallback(() => {
    setInitialCorpus(DEFAULT_VALUES.initialCorpus[region]);
    setMonthlyWithdrawal(DEFAULT_VALUES.monthlyWithdrawal[region]);
    setExpectedReturn(DEFAULT_VALUES.expectedReturn);
    setTenure(DEFAULT_VALUES.tenure);
    setInflationRate(DEFAULT_VALUES.inflationRate[region]);
  }, [region]);

  const currencySymbol = region === 'INR' ? '₹' : '$';
  const corpusRanges = INPUT_RANGES.initialCorpus[region];
  const withdrawalRanges = INPUT_RANGES.monthlyWithdrawal[region];

  // Calculate total interest earned
  const totalInterestEarned = result.yearlyData.reduce(
    (sum, row) => sum + row.interestEarned,
    0
  );

  return (
    <div className="space-y-5">
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">
        {/* Input Panel */}
        <div className="lg:col-span-1">
          <Card>
            <CardContent className="space-y-4">
              <div className="flex items-center justify-between">
                <h2 className="text-lg font-semibold text-gray-900 dark:text-white">SWP Parameters</h2>
                <Button variant="ghost" size="sm" onClick={handleReset}>
                  <RotateCcw className="w-4 h-4" />
                  Reset
                </Button>
              </div>

              <SliderInput
                label="Initial Corpus"
                value={initialCorpus}
                onChange={setInitialCorpus}
                min={corpusRanges.min}
                max={corpusRanges.max}
                step={corpusRanges.step}
                prefix={currencySymbol}
              />

              <SliderInput
                label="Monthly Withdrawal"
                value={monthlyWithdrawal}
                onChange={setMonthlyWithdrawal}
                min={withdrawalRanges.min}
                max={withdrawalRanges.max}
                step={withdrawalRanges.step}
                prefix={currencySymbol}
              />

              <SliderInput
                label="Expected Return Rate"
                value={expectedReturn}
                onChange={setExpectedReturn}
                min={INPUT_RANGES.expectedReturn.min}
                max={INPUT_RANGES.expectedReturn.max}
                step={INPUT_RANGES.expectedReturn.step}
                suffix="%"
                hint="Common range: 8-15% for equity, 6-9% for debt"
              />

              <SliderInput
                label="Withdrawal Tenure"
                value={tenure}
                onChange={setTenure}
                min={INPUT_RANGES.tenure.min}
                max={INPUT_RANGES.tenure.max}
                step={INPUT_RANGES.tenure.step}
                suffix=" years"
              />

              {showInflation && (
                <SliderInput
                  label="Expected Inflation Rate"
                  value={inflationRate}
                  onChange={setInflationRate}
                  min={INPUT_RANGES.inflationRate.min}
                  max={INPUT_RANGES.inflationRate.max}
                  step={INPUT_RANGES.inflationRate.step}
                  suffix="%"
                  hint={`Typical: ${region === 'INR' ? '5-7%' : '2-4%'} - Shows corpus value in today's money`}
                />
              )}

              {/* Corpus depletion warning */}
              {result.remainingCorpus <= 0 && (
                <div className="bg-red-50 dark:bg-red-900/30 border border-red-200 dark:border-red-700 rounded-md p-3">
                  <p className="text-sm text-red-800 dark:text-red-200">
                    <span className="font-medium">Warning:</span> Your corpus will be depleted
                    before the end of the tenure. Consider reducing monthly withdrawal or
                    increasing the expected return rate.
                  </p>
                </div>
              )}
            </CardContent>
          </Card>
        </div>

        {/* Results Panel */}
        <div className="lg:col-span-2 space-y-5">
          {/* Summary Cards */}
          <div className={`grid grid-cols-1 sm:grid-cols-2 ${showInflation ? 'lg:grid-cols-4' : 'lg:grid-cols-3'} gap-3 transition-all duration-300`}>
            <SummaryCard
              title="Total Withdrawn"
              value={formatCurrency(result.totalWithdrawn, region)}
              icon={<Wallet className="w-5 h-5" />}
              variant="investment"
            />
            <SummaryCard
              title="Interest Earned"
              value={formatCurrency(totalInterestEarned, region)}
              icon={<TrendingUp className="w-5 h-5" />}
              variant="returns"
            />
            <SummaryCard
              title="Remaining Corpus"
              value={formatCurrency(result.remainingCorpus, region)}
              icon={<Landmark className="w-5 h-5" />}
              variant="maturity"
            />
            <div
              className={`transition-all duration-300 ease-in-out ${
                showInflation
                  ? 'opacity-100 scale-100 max-w-full'
                  : 'opacity-0 scale-95 max-w-0 overflow-hidden'
              }`}
            >
              {showInflation && (
                <SummaryCard
                  title="Inflation Adjusted"
                  value={formatCurrency(result.inflationAdjustedCorpus || result.remainingCorpus, region)}
                  icon={<TrendingDown className="w-5 h-5" />}
                  variant="inflation"
                  subtitle="Corpus in today's value"
                />
              )}
            </div>
          </div>

          {/* Charts */}
          <ChartTabs
            chartData={chartData}
            investment={result.totalWithdrawn}
            returns={result.remainingCorpus}
            region={region}
            growthTitle="Corpus Depletion Over Time"
            distributionTitle="Withdrawn vs Remaining"
            showInflation={showInflation && inflationRate > 0}
            isSWP
          />
        </div>
      </div>

      {/* Yearly Table - Full Width */}
      <SWPYearlyTable data={result.yearlyData} region={region} showInflation={showInflation} />
    </div>
  );
}
