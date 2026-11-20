using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    public class FormulaEvaluator
    {
        public static readonly string[] AvailableKeywords =
        {
        "minp",
        "allowz",
        "weekly0",
        "weekly1",
        "weekly2",
        "weekly3",
        "weekly4",
        };

        private class Rule
        {
            public int ReturnValue { get; set; }
            public List<string> Conditions { get; set; } = new List<string>();
        }

        private class Rule2
        { 
            public string Signature { get; set; }
            public List<string> Conditions { get; set; } = new List<string>();
        }

        public static int Evaluate(
    EvaluateType type,
    string formula,
    CheckedData data,
    int defaultFallback = -1)
        {
            if (type == EvaluateType.signature)
                throw new ArgumentException("Use a string fallback for signature formulas.");

            string result = EvaluateCore(
                formula, data, defaultFallback.ToString(CultureInfo.InvariantCulture));

            return int.Parse(result, CultureInfo.InvariantCulture);
        }

        public static string Evaluate(
            EvaluateType type,
            string formula,
            CheckedData data,
            string defaultFallback)
        {
            if (type != EvaluateType.signature)
                throw new ArgumentException("Use an integer fallback for numeric formulas.");

            return EvaluateCore(formula, data, defaultFallback);
        }

        private static string EvaluateCore(string formula, CheckedData data, string defaultFallback)
        {
            var matches = Regex.Matches(formula, @"([^()]+)\s*\(([^()]*)\)")
                .Cast<Match>()
                .OrderBy(match => string.IsNullOrWhiteSpace(match.Groups[2].Value));

            foreach (Match match in matches)
            {
                var conditions = match.Groups[2].Value
                    .Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(condition => condition.Trim());

                if (conditions.All(condition => EvaluateSingleCondition(condition, data)))
                    return match.Groups[1].Value.Trim().Trim(',', ';').Trim();
            }

            return defaultFallback;
        }

        private static bool EvaluateSingleCondition(string condition, CheckedData data)
        {
            var varMatch = Regex.Match(condition, @"^([a-zA-Z_][a-zA-Z_0-9]*)");
            if (!varMatch.Success) return false;

            string keyword = varMatch.Groups[1].Value.ToLower();
            string remainder = condition.Substring(keyword.Length).Replace(',', '.');

            var comparisons = Regex.Matches(remainder, @"([<>=!]+)\s*(\d+(?:\.\d+)?)");
            if (comparisons.Count == 0) return false;

            foreach (Match comp in comparisons)
            {
                string op = comp.Groups[1].Value;
                double targetValue = double.Parse(comp.Groups[2].Value, CultureInfo.InvariantCulture);

                if (!EvaluateKeyword(keyword, op, targetValue, data))
                    return false;
            }

            return true;
        }

        private static bool EvaluateKeyword(string keyword, string op, double targetValue, CheckedData data)
        {
            switch (keyword)
            {
                case "minp":
                    return CompareNumbers(data.PointPercentage, op, targetValue);

                case "allowz":
                    if (op == "=" || op == "==")
                        return data.ZeroCount <= targetValue;

                    return CompareNumbers(data.ZeroCount, op, targetValue);
                default:
                    var weeklyGradeMatch = Regex.Match(keyword, @"^weekly([0-4])$");
                    if (!weeklyGradeMatch.Success)
                        return false;

                    int weeklyGrade = int.Parse(weeklyGradeMatch.Groups[1].Value);
                    int count = data.WeeklyPerformance.Count(value => value == weeklyGrade);

                    return CompareNumbers(count, op, targetValue);
            }
        }

        private static bool CompareNumbers(double actual, string op, double target)
        {
            switch (op)
            {
                case "<":
                    return actual < target;
                case ">":
                    return actual > target;
                case "=":
                case "==":
                    return Math.Abs(actual - target) < 0.0001;
                case "<=":
                case "=<":
                    return actual <= target;
                case ">=":
                case "=>":
                    return actual >= target;
                case "!=":
                    return Math.Abs(actual - target) >= 0.0001;
                default:
                    return false;
            }
        }
    }

}
