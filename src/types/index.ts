export type Region = 'INR' | 'USD';

export type CalculatorType = 'sip' | 'stepup' | 'swp';

export type StepUpFrequency = 'annual' | 'semi-annual' | 'quarterly';

export interface SIPInputs {
  monthlyInvestment: number;
  expectedReturn: number;
  tenure: number;
  region: Region;
}

export interface LumpsumInputs {
  principalAmount: number;
  expectedReturn: number;
  tenure: number;
  region: Region;
}

export interface StepUpSIPInputs extends SIPInputs {
  stepUpPercentage: number;
  stepUpFrequency: StepUpFrequency;
}

export interface SWPInputs {
  initialCorpus: number;
  monthlyWithdrawal: number;
  expectedReturn: number;
  tenure: number;
  region: Region;
}

export interface YearlyData {
  year: number;
  monthlyAmount?: number;
  annualInvestment: number;
  cumulativeInvestment: number;
  valueAtYearEnd: number;
  totalReturns: number;
  // Inflation-adjusted values (in today's money)
  inflationAdjustedValue?: number;
  inflationAdjustedReturns?: number;
}

export interface SWPYearlyData {
  year: number;
  annualWithdrawal: number;
  cumulativeWithdrawal: number;
  corpusAtYearEnd: number;
  interestEarned: number;
  // Inflation-adjusted values (in today's money)
  inflationAdjustedCorpus?: number;
  inflationAdjustedWithdrawal?: number;
}

export interface CalculationResult {
  totalInvestment: number;
  expectedReturns: number;
  maturityValue: number;
  yearlyData: YearlyData[];
  // Inflation-adjusted final values
  inflationAdjustedMaturity?: number;
  inflationAdjustedReturns?: number;
}

export interface SWPCalculationResult {
  totalWithdrawn: number;
  remainingCorpus: number;
  yearlyData: SWPYearlyData[];
  // Inflation-adjusted final values
  inflationAdjustedCorpus?: number;
}

export interface ChartDataPoint {
  year: number;
  investment: number;
  value: number;
  returns: number;
  // Inflation-adjusted values for chart
  inflationAdjustedValue?: number;
}

export interface PieChartData {
  name: string;
  value: number;
  color: string;
}
