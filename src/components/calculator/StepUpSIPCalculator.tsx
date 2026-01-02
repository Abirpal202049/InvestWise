import { useState, useMemo, useCallback, useEffect } from 'react';
import { Wallet, TrendingUp, PiggyBank, RotateCcw, TrendingDown } from 'lucide-react';
import type { Region, StepUpFrequency } from '../../types';
import { calculateStepUpSIP, toChartData } from '../../utils/calculations';
import { formatCurrency } from '../../utils/formatCurrency';
import { DEFAULT_VALUES, INPUT_RANGES, STEP_UP_FREQUENCY_OPTIONS } from '../../utils/constants';
import { Card, CardContent, SliderInput, SummaryCard, Button, RadioGroup } from '../ui';
import { ChartTabs } from '../charts';
import { SIPYearlyTable } from './YearlyTable';

interface StepUpSIPCalculatorProps {
  region: Region;
  showInflation: boolean;
}

export function StepUpSIPCalculator({ region, showInflation }: StepUpSIPCalculatorProps) {
  const [monthlyInvestment, setMonthlyInvestment] = useState(
    DEFAULT_VALUES.monthlyInvestment[region]
  );
  const [expectedReturn, setExpectedReturn] = useState(DEFAULT_VALUES.expectedReturn);
  const [tenure, setTenure] = useState(DEFAULT_VALUES.tenure);
  const [stepUpPercentage, setStepUpPercentage] = useState(DEFAULT_VALUES.stepUpPercentage);
  const [stepUpFrequency, setStepUpFrequency] = useState<StepUpFrequency>(
    DEFAULT_VALUES.stepUpFrequency
  );
  const [inflationRate, setInflationRate] = useState(DEFAULT_VALUES.inflationRate[region]);

  // Update values when region changes
  useEffect(() => {
    setMonthlyInvestment(DEFAULT_VALUES.monthlyInvestment[region]);
    setInflationRate(DEFAULT_VALUES.inflationRate[region]);
  }, [region]);

  const result = useMemo(
    () =>
      calculateStepUpSIP({
        monthlyInvestment,
        expectedReturn,
        tenure,
        region,
        stepUpPercentage,
        stepUpFrequency,
      }, inflationRate),
    [monthlyInvestment, expectedReturn, tenure, region, stepUpPercentage, stepUpFrequency, inflationRate]
  );

  const chartData = useMemo(() => toChartData(result.yearlyData), [result.yearlyData]);

  const handleReset = useCallback(() => {
    setMonthlyInvestment(DEFAULT_VALUES.monthlyInvestment[region]);
    setExpectedReturn(DEFAULT_VALUES.expectedReturn);
    setTenure(DEFAULT_VALUES.tenure);
    setStepUpPercentage(DEFAULT_VALUES.stepUpPercentage);
    setStepUpFrequency(DEFAULT_VALUES.stepUpFrequency);
    setInflationRate(DEFAULT_VALUES.inflationRate[region]);
  }, [region]);

  const currencySymbol = region === 'INR' ? '₹' : '$';
  const ranges = INPUT_RANGES.monthlyInvestment[region];

  // Get final monthly SIP amount
  const finalMonthlyAmount =
    result.yearlyData.length > 0
      ? result.yearlyData[result.yearlyData.length - 1].monthlyAmount
      : monthlyInvestment;

  return (
    <div className="space-y-5">
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">
        {/* Input Panel */}
        <div className="lg:col-span-1">
          <Card>
            <CardContent className="space-y-4">
              <div className="flex items-center justify-between">
                <h2 className="text-lg font-semibold text-gray-900 dark:text-white">Step-Up SIP Parameters</h2>
                <Button variant="ghost" size="sm" onClick={handleReset}>
                  <RotateCcw className="w-4 h-4" />
                  Reset
                </Button>
              </div>

              <SliderInput
                label="Initial Monthly Investment"
                value={monthlyInvestment}
                onChange={setMonthlyInvestment}
                min={ranges.min}
                max={ranges.max}
                step={ranges.step}
                prefix={currencySymbol}
              />

              <SliderInput
                label="Annual Step-Up"
                value={stepUpPercentage}
                onChange={setStepUpPercentage}
                min={INPUT_RANGES.stepUpPercentage.min}
                max={INPUT_RANGES.stepUpPercentage.max}
                step={INPUT_RANGES.stepUpPercentage.step}
                suffix="%"
                hint="Typical range: 5-10%"
              />

              <RadioGroup
                label="Step-Up Frequency"
                options={STEP_UP_FREQUENCY_OPTIONS}
                value={stepUpFrequency}
                onChange={(value) => setStepUpFrequency(value as StepUpFrequency)}
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
                label="Investment Tenure"
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
                  hint={`Typical: ${region === 'INR' ? '5-7%' : '2-4%'} - Adjusts future value to today's money`}
                />
              )}

              {/* Final SIP Amount Info */}
              <div className="bg-amber-50 dark:bg-amber-900/30 border border-amber-200 dark:border-amber-700 rounded-md p-3">
                <p className="text-sm text-amber-800 dark:text-amber-200">
                  <span className="font-medium">Final Monthly SIP (Year {tenure}):</span>
                  <br />
                  <span className="text-lg font-bold">
                    {formatCurrency(finalMonthlyAmount || 0, region)}
                  </span>
                </p>
              </div>
            </CardContent>
          </Card>
        </div>

        {/* Results Panel */}
        <div className="lg:col-span-2 space-y-5">
          {/* Summary Cards */}
          <div className={`grid grid-cols-1 sm:grid-cols-2 ${showInflation ? 'lg:grid-cols-4' : 'lg:grid-cols-3'} gap-3 transition-all duration-300`}>
            <SummaryCard
              title="Total Investment"
              value={formatCurrency(result.totalInvestment, region)}
              icon={<Wallet className="w-5 h-5" />}
              variant="investment"
            />
            <SummaryCard
              title="Expected Returns"
              value={formatCurrency(result.expectedReturns, region)}
              icon={<TrendingUp className="w-5 h-5" />}
              variant="returns"
            />
            <SummaryCard
              title="Maturity Value"
              value={formatCurrency(result.maturityValue, region)}
              icon={<PiggyBank className="w-5 h-5" />}
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
                  value={formatCurrency(result.inflationAdjustedMaturity || result.maturityValue, region)}
                  icon={<TrendingDown className="w-5 h-5" />}
                  variant="inflation"
                  subtitle="Today's value"
                />
              )}
            </div>
          </div>

          {/* Charts */}
          <ChartTabs
            chartData={chartData}
            investment={result.totalInvestment}
            returns={result.expectedReturns}
            region={region}
            showInflation={showInflation && inflationRate > 0}
          />
        </div>
      </div>

      {/* Yearly Table - Full Width */}
      <SIPYearlyTable data={result.yearlyData} region={region} showMonthlyAmount showInflation={showInflation} />
    </div>
  );
}
