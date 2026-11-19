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

            

            CalcWeek();

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

        private void CalcWeek()
        {
            for (int i = 2; i <= GetLastPopulatedRow(worksheet2, 2); i++)
            {
                object neptunCodeValue = GetCellValue(worksheet2, i, 2);
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

                var result = FormulaEvaluator.Evaluate(EvaluateType.weekly, textBox1.Text, eval, -1);

                System.Windows.Forms.RadioButton radioButton = panel1.Controls.OfType<System.Windows.Forms.RadioButton>().FirstOrDefault(rb => rb.Checked);
                switch (radioButton?.Name)
                {
                    case "radioButton1":
                        GetCellValue(worksheet2, i, 9) = result;
                        student.WeeklyPerformance[0] = result;
                        break;
                    case "radioButton2":
                        GetCellValue(worksheet2, i, 10) = result;
                        student.WeeklyPerformance[1] = result;
                        break;
                    case "radioButton3":
                        GetCellValue(worksheet2, i, 11) = result;
                        student.WeeklyPerformance[2] = result;
                        break;
                    case "radioButton4":
                        GetCellValue(worksheet2, i, 12) = result;
                        student.WeeklyPerformance[3] = result;
                        break;
                    default:
                        break;
                }
            }
        }

        public void CalcSignature()
        {
            for (int i = 2; i <= GetLastPopulatedRow(worksheet2, 2); i++)
            {
                object neptunCodeValue = GetCellValue(worksheet2, i, 2);
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
                var result = FormulaEvaluator.Evaluate(EvaluateType.signature, textBox7.Text, eval, -1);
                GetCellValue(worksheet2,i,13) = result;
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

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            label8.Visible = !label8.Visible;
            label9.Visible = !label9.Visible;
            textBox5.Visible = !textBox5.Visible;
            textBox6.Visible = !textBox6.Visible;
        }
    }
}