using System;
using System.Collections.Generic;
using System.Linq;
using MathNet.Numerics;
using MathNet.Numerics.Statistics;
using MathNet.Numerics.LinearRegression;
using unvell.ReoGrid;
using ClosedXML.Excel;

namespace ExcelClone.Services
{
    public class DescriptiveStatisticsResult
    {
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

    public class LinearRegressionResult
    {
        public double Intercept { get; set; }
        public double Slope { get; set; }
        public double RSquared { get; set; }
        public double StandardError { get; set; }
    }

    public static class NumericalAnalysisService
    {
        public static DescriptiveStatisticsResult CalculateDescriptiveStatistics(IEnumerable<double> values)
        {
            var data = values.Where(v => !double.IsNaN(v) && !double.IsInfinity(v)).ToList();
            if (data.Count == 0)
            {
                throw new ArgumentException("Dataset contains no valid numeric values.");
            }

            double mean = data.Mean();
            double variance = data.Variance();
            double stdDev = data.StandardDeviation();
            double median = data.Median();
            double min = data.Minimum();
            double max = data.Maximum();
            double sum = data.Sum();
            double count = data.Count;
            double stdError = count > 1 ? stdDev / Math.Sqrt(count) : 0;
            double kurtosis = count > 3 ? data.Kurtosis() : 0;
            double skewness = count > 2 ? data.Skewness() : 0;

            // Mode calculation
            var grouped = data.GroupBy(v => v).OrderByDescending(g => g.Count()).ToList();
            double mode = grouped.First().Key;

            return new DescriptiveStatisticsResult
            {
                Mean = mean,
                StandardError = stdError,
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

        public static LinearRegressionResult CalculateLinearRegression(IEnumerable<double> xVals, IEnumerable<double> yVals)
        {
            var x = xVals.ToArray();
            var y = yVals.ToArray();

            if (x.Length != y.Length || x.Length < 2)
            {
                throw new ArgumentException("X and Y arrays must have identical length of at least 2 observations.");
            }

            var p = SimpleRegression.Fit(x, y);
            double rSquared = GoodnessOfFit.RSquared(x.Select(val => p.Item1 + p.Item2 * val), y);

            // Compute standard error of estimate
            double sumSqResiduals = 0;
            for (int i = 0; i < x.Length; i++)
            {
                double pred = p.Item1 + p.Item2 * x[i];
                sumSqResiduals += Math.Pow(y[i] - pred, 2);
            }
            double se = x.Length > 2 ? Math.Sqrt(sumSqResiduals / (x.Length - 2)) : 0;

            return new LinearRegressionResult
            {
                Intercept = p.Item1,
                Slope = p.Item2,
                RSquared = rSquared,
                StandardError = se
            };
        }

        public static List<double> CalculateExponentialSmoothing(IEnumerable<double> values, double alpha)
        {
            var data = values.ToList();
            if (data.Count == 0) return new List<double>();
            if (alpha <= 0 || alpha > 1) throw new ArgumentOutOfRangeException(nameof(alpha), "Damping factor alpha must be in (0, 1].");

            var result = new List<double>(data.Count);
            double s = data[0];
            result.Add(s);

            for (int i = 1; i < data.Count; i++)
            {
                s = alpha * data[i] + (1 - alpha) * s;
                result.Add(s);
            }
            return result;
        }

        public static List<double> CalculateMovingAverage(IEnumerable<double> values, int windowSize)
        {
            var data = values.ToList();
            if (windowSize <= 0 || windowSize > data.Count)
                throw new ArgumentOutOfRangeException(nameof(windowSize), "Interval must be > 0 and <= sample size.");

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

        public static bool SolveGoalSeek(Worksheet sheet, string targetCellAddress, double targetValue, string changingCellAddress, out double resultValue, int maxIterations = 100, double tolerance = 0.0001)
        {
            var targetCell = sheet.Cells[targetCellAddress];
            var changingCell = sheet.Cells[changingCellAddress];

            double x0 = ParseNumericValue(changingCell?.Data);
            if (Math.Abs(x0) < 1e-9) x0 = 1.0;

            if (changingCell != null) changingCell.Data = x0;
            sheet.Recalculate();
            double y0 = ParseNumericValue(targetCell?.Data) - targetValue;
            if (Math.Abs(y0) <= tolerance)
            {
                resultValue = x0;
                return true;
            }

            double x1 = x0 + (Math.Abs(x0) > 1.0 ? 0.05 * x0 : 0.1);
            if (changingCell != null) changingCell.Data = x1;
            sheet.Recalculate();
            double y1 = ParseNumericValue(targetCell?.Data) - targetValue;

            for (int iter = 0; iter < maxIterations; iter++)
            {
                if (Math.Abs(y1) <= tolerance)
                {
                    resultValue = x1;
                    return true;
                }

                if (Math.Abs(y1 - y0) < 1e-12)
                {
                    x1 += 0.5;
                    if (changingCell != null) changingCell.Data = x1;
                    sheet.Recalculate();
                    y1 = ParseNumericValue(targetCell?.Data) - targetValue;
                    continue;
                }

                double xNext = x1 - y1 * ((x1 - x0) / (y1 - y0));
                x0 = x1;
                y0 = y1;
                x1 = xNext;

                if (changingCell != null) changingCell.Data = x1;
                sheet.Recalculate();
                y1 = ParseNumericValue(targetCell?.Data) - targetValue;
            }

            resultValue = x1;
            return Math.Abs(y1) <= (tolerance * 10);
        }

        public static bool SolveGoalSeek(IXLWorksheet sheet, string targetCellAddress, double targetValue, string changingCellAddress, out double resultValue, int maxIterations = 100, double tolerance = 0.0001)
        {
            var targetCell = sheet.Cell(targetCellAddress);
            var changingCell = sheet.Cell(changingCellAddress);

            double x0 = ParseNumericValue(changingCell.Value);
            if (Math.Abs(x0) < 1e-9) x0 = 1.0;

            changingCell.Value = x0;
            double y0 = ParseNumericValue(targetCell.Value) - targetValue;
            if (Math.Abs(y0) <= tolerance)
            {
                resultValue = x0;
                return true;
            }

            double x1 = x0 + (Math.Abs(x0) > 1.0 ? 0.05 * x0 : 0.1);
            changingCell.Value = x1;
            double y1 = ParseNumericValue(targetCell.Value) - targetValue;

            for (int iter = 0; iter < maxIterations; iter++)
            {
                if (Math.Abs(y1) <= tolerance)
                {
                    resultValue = x1;
                    return true;
                }

                if (Math.Abs(y1 - y0) < 1e-12)
                {
                    x1 += 0.5;
                    changingCell.Value = x1;
                    y1 = ParseNumericValue(targetCell.Value) - targetValue;
                    continue;
                }

                double xNext = x1 - y1 * ((x1 - x0) / (y1 - y0));
                x0 = x1;
                y0 = y1;
                x1 = xNext;

                changingCell.Value = x1;
                y1 = ParseNumericValue(targetCell.Value) - targetValue;
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
