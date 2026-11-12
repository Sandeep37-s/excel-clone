using System;
using System.Collections.Generic;
using System.Linq;
using MathNet.Numerics.Statistics;
using MathNet.Numerics.LinearRegression;
using Syncfusion.XlsIO;

namespace ExcelClone.Services
{
    public class DescriptiveStatisticsResult
    {
        public string VariableName { get; set; } = "Variable";
        public double Mean { get; set; }
        public double StandardError { get; set; }
        public double Median { get; set; }
        public double Mode { get; set; }
        public double StandardDeviation { get; set; }
        public double SampleVariance { get; set; }
        public double Kurtosis { get; set; }
        public double Skewness { get; set; }
        public double Range { get; set; }
        public double Minimum { get; set; }
        public double Maximum { get; set; }
        public double Sum { get; set; }
        public int Count { get; set; }
    }

    public class RegressionResult
    {
        public double Slope { get; set; }
        public double Intercept { get; set; }
        public double RSquared { get; set; }
        public double StandardError { get; set; }
        public int Observations { get; set; }
    }

    public static class NumericalAnalysisService
    {
        public static DescriptiveStatisticsResult CalculateDescriptiveStatistics(IEnumerable<double> values, string name = "Data")
        {
            var list = values.Where(v => !double.IsNaN(v) && !double.IsInfinity(v)).ToList();
            if (list.Count == 0)
                throw new InvalidOperationException("Input data contains no valid numeric values.");

            double count = list.Count;
            double mean = list.Mean();
            double stdDev = count > 1 ? list.StandardDeviation() : 0.0;
            double stdErr = count > 1 ? stdDev / Math.Sqrt(count) : 0.0;
            double median = list.Median();
            double min = list.Min();
            double max = list.Max();
            double sum = list.Sum();
            double variance = count > 1 ? list.Variance() : 0.0;
            double skewness = count > 2 ? list.Skewness() : 0.0;
            double kurtosis = count > 3 ? list.Kurtosis() : 0.0;

            // Mode calculation
            var modeGroup = list.GroupBy(v => v).OrderByDescending(g => g.Count()).FirstOrDefault();
            double mode = modeGroup?.Key ?? mean;

            return new DescriptiveStatisticsResult
            {
                VariableName = name,
                Mean = mean,
                StandardError = stdErr,
                Median = median,
                Mode = mode,
                StandardDeviation = stdDev,
                SampleVariance = variance,
                Kurtosis = kurtosis,
                Skewness = skewness,
                Range = max - min,
                Minimum = min,
                Maximum = max,
                Sum = sum,
                Count = (int)count
            };
        }

        public static RegressionResult CalculateLinearRegression(IList<double> x, IList<double> y)
        {
            if (x.Count != y.Count || x.Count < 2)
                throw new ArgumentException("X and Y datasets must have equal lengths and at least 2 observations.");

            var (intercept, slope) = SimpleRegression.Fit(x.ToArray(), y.ToArray());

            // Compute R^2
            double r = Correlation.Pearson(x, y);
            double rSquared = r * r;

            // Compute Residual Standard Error
            double sse = 0.0;
            for (int i = 0; i < x.Count; i++)
            {
                double pred = intercept + slope * x[i];
                double diff = y[i] - pred;
                sse += diff * diff;
            }
            double se = x.Count > 2 ? Math.Sqrt(sse / (x.Count - 2)) : 0.0;

            return new RegressionResult
            {
                Slope = slope,
                Intercept = intercept,
                RSquared = rSquared,
                StandardError = se,
                Observations = x.Count
            };
        }

        public static double CalculatePearsonCorrelation(IList<double> x, IList<double> y)
        {
            if (x.Count != y.Count || x.Count < 2)
                throw new ArgumentException("Datasets must be of equal size and have at least 2 points.");
            return Correlation.Pearson(x, y);
        }

        public static List<double> CalculateSimpleMovingAverage(IList<double> data, int windowSize)
        {
            if (windowSize <= 0 || windowSize > data.Count)
                throw new ArgumentException("Window size must be greater than 0 and less than or equal to data count.");

            var result = new List<double>();
            for (int i = 0; i < data.Count; i++)
            {
                if (i < windowSize - 1)
                {
                    result.Add(double.NaN);
                }
                else
                {
                    double sum = 0;
                    for (int j = i - windowSize + 1; j <= i; j++)
                    {
                        sum += data[j];
                    }
                    result.Add(sum / windowSize);
                }
            }
            return result;
        }

        /// <summary>
        /// Solves Goal Seek using Secant / Bisection numerical method on an active worksheet.
        /// </summary>
        public static bool SolveGoalSeek(IWorksheet sheet, string targetCellAddress, double targetValue, string changingCellAddress, out double resultValue, int maxIterations = 100, double tolerance = 0.0001)
        {
            var targetCell = sheet.Range[targetCellAddress];
            var changingCell = sheet.Range[changingCellAddress];

            double x0 = changingCell.Number;
            if (Math.Abs(x0) < 1e-9) x0 = 1.0;

            changingCell.Number = x0;
            sheet.EnableSheetCalculations();
            double y0 = ParseNumericValue(targetCell.CalculatedValue) - targetValue;
            if (Math.Abs(y0) <= tolerance)
            {
                resultValue = x0;
                return true;
            }

            // Step 2: Perturb x slightly to find slope
            double x1 = x0 + (Math.Abs(x0) > 1.0 ? 0.05 * x0 : 0.1);
            changingCell.Number = x1;
            sheet.EnableSheetCalculations();
            double y1 = ParseNumericValue(targetCell.CalculatedValue) - targetValue;

            for (int iter = 0; iter < maxIterations; iter++)
            {
                if (Math.Abs(y1) <= tolerance)
                {
                    resultValue = x1;
                    return true;
                }

                if (Math.Abs(y1 - y0) < 1e-12)
                {
                    // Degenerate slope: nudge x1
                    x1 += 0.5;
                    changingCell.Number = x1;
                    sheet.EnableSheetCalculations();
                    y1 = ParseNumericValue(targetCell.CalculatedValue) - targetValue;
                    continue;
                }

                // Secant update
                double xNext = x1 - y1 * ((x1 - x0) / (y1 - y0));
                x0 = x1;
                y0 = y1;
                x1 = xNext;

                changingCell.Number = x1;
                sheet.EnableSheetCalculations();
                y1 = ParseNumericValue(targetCell.CalculatedValue) - targetValue;
            }

            resultValue = x1;
            return Math.Abs(y1) <= (tolerance * 10);
        }

        private static double ParseNumericValue(object? val)
        {
            if (val == null) return 0.0;
            if (val is double d) return d;
            if (val is int i) return i;
            if (double.TryParse(val.ToString(), out double parsed)) return parsed;
            return 0.0;
        }
    }
}

