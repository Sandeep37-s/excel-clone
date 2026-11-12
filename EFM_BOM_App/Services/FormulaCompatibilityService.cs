using System;
using System.Text.RegularExpressions;

namespace EFM_BOM_App.Services
{
    public static class FormulaCompatibilityService
    {
        private static readonly Regex VarPRegex = new Regex(@"\bVAR\.P\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex VarSRegex = new Regex(@"\bVAR\.S\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex StdevPRegex = new Regex(@"\bSTDEV\.P\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex StdevSRegex = new Regex(@"\bSTDEV\.S\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex PercentileIncRegex = new Regex(@"\bPERCENTILE\.INC\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex QuartileIncRegex = new Regex(@"\bQUARTILE\.INC\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex EmptySumRegex = new Regex(@"\bSUM\s*\(\s*\)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex EmptyCountRegex = new Regex(@"\bCOUNT\s*\(\s*\)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex ConvertMiKmRegex = new Regex(@"\bCONVERT\s*\(\s*([^,]+)\s*,\s*""mi""\s*,\s*""km""\s*\)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Normalizes modern Excel 365 and Excel 2010+ formula syntax into OpenXML / CalcEngine compatible expressions.
        /// Resolves dot-notation aliases, empty argument edge-cases, and unit discrepancies.
        /// </summary>
        public static string NormalizeFormula(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula)) return formula;

            string normalized = formula;

            // 1. Map modern dot notation to universal OpenXML formula engine names
            normalized = VarPRegex.Replace(normalized, "VARP(");
            normalized = VarSRegex.Replace(normalized, "VAR(");
            normalized = StdevPRegex.Replace(normalized, "STDEVP(");
            normalized = StdevSRegex.Replace(normalized, "STDEV(");
            normalized = PercentileIncRegex.Replace(normalized, "PERCENTILE(");
            normalized = QuartileIncRegex.Replace(normalized, "QUARTILE(");

            // 2. Normalize empty function arguments to prevent parser exceptions
            normalized = EmptySumRegex.Replace(normalized, "SUM(0)");
            normalized = EmptyCountRegex.Replace(normalized, "COUNT(\"\")");

            // 3. Convert mile to km unit normalization
            normalized = ConvertMiKmRegex.Replace(normalized, "($1 * 1.609344)");

            // 4. Normalize 3-argument AVERAGEIF to universally supported AVERAGEIFS
            // Format: AVERAGEIF(criteria_range, criteria, sum_range) -> AVERAGEIFS(sum_range, criteria_range, criteria)
            normalized = NormalizeAverageIf(normalized);

            return normalized;
        }

        private static string NormalizeAverageIf(string formula)
        {
            var match = Regex.Match(formula, @"AVERAGEIF\s*\(([^,]+),\s*([^,]+),\s*([^\)]+)\)", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                string critRange = match.Groups[1].Value.Trim();
                string criteria = match.Groups[2].Value.Trim();
                string avgRange = match.Groups[3].Value.Trim();
                return formula.Replace(match.Value, $"AVERAGEIFS({avgRange}, {critRange}, {criteria})");
            }
            return formula;
        }
    }
}

