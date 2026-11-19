using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        bool kiszh = false;
        bool nagyzh = false;
        bool beadando = false;
        string filePath = "";
        string filePath2 = "";
        Workbook workbook;
        Workbook workbook2;
        Worksheet worksheet;
        Worksheet worksheet2;
        Microsoft.Office.Interop.Excel.Application excelApp;
        private readonly Dictionary<string, StudentData> studentsByNeptun = new Dictionary<string, StudentData>();
        private double maxPointsGlobal = 0;
        private int startingRow;
        public Form1()
        {
            InitializeComponent();
            FormClosing += (sender, e) => ShutdownExcel();
        }

        private void EnsureExcelApp()
        {
            if (excelApp == null)
                excelApp = new Microsoft.Office.Interop.Excel.Application();
        }

        private object GetCellValue(Worksheet sheet, int row, int col)
        {
            Range cell = (Range)sheet.Cells[row, col];

            if ((bool)cell.MergeCells)
            {
                // Grab the top-left cell of the merged group block
                Range firstCell = (Range)cell.MergeArea.Cells[1, 1];
                return firstCell.Value2;
            }

            return cell.Value2;
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            //gombok eltuntetese, megjelenitese
            numericUpDown1.Visible = !numericUpDown1.Visible;
            numericUpDown2.Visible = !numericUpDown2.Visible;
            label1.Visible = !label1.Visible;
            label4.Visible = !label4.Visible;
            textBox2.Visible = !textBox2.Visible;
            checkBox4.Visible = !checkBox4.Visible;
            kiszh = !kiszh;
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            //gombok eltuntetese, megjelenitese
            numericUpDown5.Visible = !numericUpDown5.Visible;
            numericUpDown6.Visible = !numericUpDown6.Visible;
            label3.Visible = !label3.Visible;
            label6.Visible = !label6.Visible;
            textBox4.Visible = !textBox4.Visible;
            checkBox6.Visible = !checkBox6.Visible;
            beadando = !beadando;


        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

            numericUpDown3.Visible = !numericUpDown3.Visible;
            numericUpDown4.Visible = !numericUpDown4.Visible;
            label2.Visible = !label2.Visible;
            label5.Visible = !label5.Visible;
            textBox3.Visible = !textBox3.Visible;
            checkBox5.Visible = !checkBox5.Visible;
            nagyzh = !nagyzh;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //fajl kivalasztas, csak excel
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog1.FileName;
                readInitialExcelFile(filePath);
                MessageBox.Show("Az adatok beolvasása megtörtént!");
            }
            
        }
        private void button2_Click_1(object sender, EventArgs e)
        {
            //fajl kivalasztas, csak excel
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                filePath2 = openFileDialog2.FileName;
                readEndExcelFile(filePath2);
                MessageBox.Show("Az adatok beolvasása megtörtént!");
            }
            
        }
        private void readInitialExcelFile(string path)
        {
            EnsureExcelApp();
            workbook = excelApp.Workbooks.Open(path);
            worksheet = (Worksheet)workbook.Worksheets[1];
        }
        private void readEndExcelFile(string path2)
        {
            EnsureExcelApp();
            workbook2 = excelApp.Workbooks.Open(path2);
            worksheet2 = (Worksheet)workbook2.Worksheets[1];
        }

        private int GetLastPopulatedRow(Worksheet sheet, int column)
        {
            return sheet.Cells[sheet.Rows.Count, column].End(XlDirection.xlUp).Row;
        }

        private void PerformFunction(NumericUpDown startC, NumericUpDown endC, System.Windows.Forms.TextBox point, System.Windows.Forms.CheckBox percent, Func<StudentData, List<object>> targetListSelector)
        {
            int startColumn = (int)startC.Value;
            int endColumn = (int)endC.Value;
            int[] maxPoints = point.Text.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries).Select(int.Parse).ToArray();
            if (maxPoints.Length == 1&& percent.Checked)
            {
                maxPointsGlobal += maxPoints[0] * (((endColumn - startColumn)/2) + 1);
            }
            else
            {
                maxPointsGlobal += maxPoints.Sum();
            }

            if (percent.Checked)
            {
                copyFunction(targetListSelector, startColumn, endColumn, maxPoints, true);
            }
            else
            {
                copyFunction(targetListSelector, startColumn, endColumn, maxPoints);
            }

        }

        private void copyFunction(Func<StudentData, List<object>> getValues, int startColumn, int endColumn, int[] maxPoints, bool percentagePresent = false)
        {
            //beolvasom az oszlop szamait, majd feldolgozom az ertekeket


            for (int i = startingRow; i <= GetLastPopulatedRow(worksheet, 3); i++)
            {
                //double currentPoints = 0;
                object neptunCodeValue = worksheet.Cells[i, 3].Value2;
                if (neptunCodeValue == null)
                {
                    continue;
                }
                //elmentem a neptun kodot, hogyha mas sorrendbe lenne a kimenetbe, akkor is meg lehessen talalni mindent
                string neptunCode = neptunCodeValue.ToString();

                if (!studentsByNeptun.TryGetValue(neptunCode, out StudentData student))
                {
                    student = new StudentData();
                    studentsByNeptun[neptunCode] = student;
                }

                var targetList = getValues(student);

                if (percentagePresent)
                {
                    int ind = 0;
                    for (int j = startColumn; j <= endColumn; j = j + 2)
                    {
                        
                        var cellValue = GetCellValue(worksheet, i, j);
                        try
                        {
                            targetList.Add($"{cellValue}/{maxPoints[ind]}");
                        }
                        catch (IndexOutOfRangeException)
                        {

                            targetList.Add($"{cellValue}/{maxPoints[0]}");
                        }
                        student.TotalEarnedPoints += Convert.ToDouble(cellValue);
                        if (cellValue == null || cellValue.ToString() == "0")
                        {
                            student.ZeroCount++;
                        }
                        ind++;
                    }
                }
                else
                {
                    for (int j = startColumn; j <= endColumn; j++)
                    {
                        var cellValue = GetCellValue(worksheet, i, j);
                        try
                        {
                            targetList.Add($"{cellValue}/{maxPoints[j - startColumn]}");
                        }
                        catch (IndexOutOfRangeException)
                        {

                            targetList.Add($"{cellValue}/{maxPoints[0]}");
                        }
                        ;
                        student.TotalEarnedPoints += Convert.ToDouble(cellValue);
                        if (cellValue == null || cellValue.ToString() == "0")
                        {
                            student.ZeroCount++;
                        }
                    }
                }
                student.PointPercentage = Math.Round((student.TotalEarnedPoints / maxPointsGlobal) * 100, 2);
            }
            
            var destinationRowsByNeptunCode = new Dictionary<string, int>();
            int lastDestinationRow = GetLastPopulatedRow(worksheet2, 2);
            for (int i = 2; i <= lastDestinationRow; i++)
            {
                object neptunCodeValue = worksheet2.Cells[i, 2].Value2;
                if (neptunCodeValue != null)
                {
                    string neptunCode = neptunCodeValue.ToString();
                    if (!destinationRowsByNeptunCode.ContainsKey(neptunCode))
                    {
                        destinationRowsByNeptunCode[neptunCode] = i;
                    }
                }
            }
            int columnOffset = 15;
            foreach (var kvp in studentsByNeptun)
            {
                if (destinationRowsByNeptunCode.TryGetValue(kvp.Key, out int destinationRow))
                {
                    List<object> values = getValues(kvp.Value);
                    for (int j = 0; j < values.Count; j++)
                    {
                        while (worksheet2.Cells[destinationRow, columnOffset + j].Value2 != null)
                        {
                            columnOffset++;
                        }
                        worksheet2.Cells[destinationRow, columnOffset + j].Value2 = values[j];
                    }
                }
                columnOffset = 15;
            }

        }

        private void kiszhFunction()
        {
            PerformFunction(numericUpDown1, numericUpDown2, textBox2, checkBox4, s => s.KisZhValues);

        }


        private void nagyzhFunction()
        {
           PerformFunction(numericUpDown4, numericUpDown3, textBox3, checkBox5, s => s.NagyZhValues);
        }
        private void beadandoFunction()
        {
            PerformFunction(numericUpDown6, numericUpDown5, textBox4, checkBox6, s => s.BeadandoValues);


        }

        private void button3_Click(object sender, EventArgs e)
        {
            startingRow = (int)numericUpDown7.Value;
            studentsByNeptun.Clear();
            maxPointsGlobal = 0;
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            if (string.IsNullOrEmpty(filePath) || string.IsNullOrEmpty(filePath2))
            {
                MessageBox.Show("Kérlek töltsd be mind a két filet!");
                return;
            }
            bool isAnySelected = panel1.Controls.OfType<System.Windows.Forms.RadioButton>().Any(rb => rb.Checked);

            if (!isAnySelected){ MessageBox.Show("Kérlek válaszd ki a hetet!"); return; }
            if (checkBox1.Checked)
            {
                kiszhFunction();
            }
            if (checkBox2.Checked)
            {
                nagyzhFunction();
            }
            if (checkBox3.Checked)
            {
                beadandoFunction();
            }

            

            CalcIdk();

            excelApp.DisplayAlerts = false;
            
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Files|*.xlsx;*.xlsm";
                saveFileDialog.Title = @"Save Excel File" + " - " + timestamp;
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    workbook2.SaveAs(saveFileDialog.FileName);
                }
            }
            finally
            {
                CloseCurrentWorkbooks();
                MessageBox.Show("Az adatok írása megtörtént!");
            }
            
        }

        private void CalcIdk()
        {
            for (int i = 2; i <= GetLastPopulatedRow(worksheet2, 2); i++)
            {
                object neptunCodeValue = worksheet2.Cells[i, 2].Value2;
                if (neptunCodeValue == null)
                    continue;

                string neptunCode = neptunCodeValue.ToString();


                if (!studentsByNeptun.TryGetValue(neptunCode, out StudentData student))
                    continue;

                
                var eval = new CheckedData
                {
                    PointPercentage = student.PointPercentage,
                    ZeroCount = student.ZeroCount
                };

                var result = FormulaEvaluator.Evaluate(textBox1.Text, eval, -1);

                System.Windows.Forms.RadioButton radioButton = panel1.Controls.OfType<System.Windows.Forms.RadioButton>().FirstOrDefault(rb => rb.Checked);
                switch (radioButton?.Name)
                {
                    case "radioButton1":
                        worksheet2.Cells[i, 9].Value2 = result;
                        break;
                    case "radioButton2":
                        worksheet2.Cells[i, 10].Value2 = result;
                        break;
                    case "radioButton3":
                        worksheet2.Cells[i, 11].Value2 = result;
                        break;
                    case "radioButton4":
                        worksheet2.Cells[i, 12].Value2 = result;
                        break;
                    default:
                        break;
                }
            }
        }

        private void CloseCurrentWorkbooks()
        {
            if (workbook2 != null) workbook2.Close(false);
            if (workbook != null) workbook.Close(false);

            ReleaseComObject(worksheet2);
            ReleaseComObject(worksheet);
            ReleaseComObject(workbook2);
            ReleaseComObject(workbook);

            worksheet2 = null;
            worksheet = null;
            workbook2 = null;
            workbook = null;
        }

        private void ShutdownExcel()
        {
            CloseCurrentWorkbooks();

            if (excelApp != null)
            {
                excelApp.Quit();
                ReleaseComObject(excelApp);
                excelApp = null;
            }
        }

        private static void ReleaseComObject(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
                Marshal.ReleaseComObject(comObject);
        }
    }
}






public class CheckedData
{
    public double PointPercentage { get; set; }
    public int ZeroCount { get; set; }
}

public class FormulaEvaluator
{
    public static readonly string[] AvailableKeywords =
    {
        "minp",
        "allowz",
    };

    private class Rule
    {
        public int ReturnValue { get; set; }
        public List<string> Conditions { get; set; } = new List<string>();
    }

    public static int Evaluate(string formula, CheckedData data, int defaultFallback = -1)
    {
        var rules = new List<Rule>();
        var ruleMatches = Regex.Matches(formula, @"(\d+)\s*\(([^)]*)\)");

        foreach (Match match in ruleMatches)
        {
            var rule = new Rule { ReturnValue = int.Parse(match.Groups[1].Value) };
            string rawConditions = match.Groups[2].Value.Trim();

            if (!string.IsNullOrEmpty(rawConditions))
            {
                rule.Conditions = rawConditions
                    .Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(c => c.Trim())
                    .ToList();
            }
            rules.Add(rule);
        }

        foreach (var rule in rules.OrderBy(r => r.Conditions.Count == 0))
        {
            if (rule.Conditions.All(cond => EvaluateSingleCondition(cond, data)))
                return rule.ReturnValue;
        }

        return defaultFallback;
    }

    private static bool EvaluateSingleCondition(string condition, CheckedData data)
    {
        var varMatch = Regex.Match(condition, @"^([a-zA-Z_]+)");
        if (!varMatch.Success) return false;

        string keyword = varMatch.Groups[1].Value.ToLower();
        string remainder = condition.Substring(keyword.Length);

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
                return false;
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
