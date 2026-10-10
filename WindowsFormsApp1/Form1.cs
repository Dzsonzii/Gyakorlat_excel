using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        bool kiszh = false;
        bool nagyzh = false;
        bool beadando = false;
        string filePath = "";
        Workbook workbook;
        Workbook workbook2;
        Worksheet worksheet;
        Worksheet worksheet2;
        Microsoft.Office.Interop.Excel.Application excelApp;
        private readonly Dictionary<string, List<object>> kisZhValuesByNeptunCode = new Dictionary<string, List<object>>();
        private readonly Dictionary<string, List<object>> nagyZhValuesByNeptunCode = new Dictionary<string, List<object>>();
        private readonly Dictionary<string, List<object>> beadandoValuesByNeptunCode = new Dictionary<string, List<object>>();
        public Form1()
        {
            InitializeComponent();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            numericUpDown1.Visible = !numericUpDown1.Visible;
            numericUpDown2.Visible = !numericUpDown2.Visible;
            label1.Visible = !label1.Visible;
            kiszh = !kiszh;
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            numericUpDown3.Visible = !numericUpDown3.Visible;
            numericUpDown4.Visible = !numericUpDown4.Visible;
            label2.Visible = !label2.Visible;
            nagyzh = !nagyzh;

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            numericUpDown5.Visible = !numericUpDown5.Visible;
            numericUpDown6.Visible = !numericUpDown6.Visible;
            label3.Visible = !label3.Visible;
            beadando = !beadando;

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //open file dialog, search for excel file, and get the path of the file
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog1.FileName;
                readInitialExcelFile(filePath);
            }
        }
        private void button2_Click_1(object sender, EventArgs e)
        {
            //open file dialog, search for excel file, and get the path of the file
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                string filePath2 = openFileDialog2.FileName;
                readEndExcelFile(filePath2);
            }
        }
        private void readInitialExcelFile(string filePath)
        {
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            workbook = excelApp.Workbooks.Open(filePath);
            worksheet = workbook.Sheets[1];


            //workbook.Close(false);
            //excelApp.Quit();
        }
        private void readEndExcelFile(string filePath2)
        {
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            workbook2 = excelApp.Workbooks.Open(filePath2);
            worksheet2 = workbook2.Sheets[1];


            //workbook.Close(false);
            //excelApp.Quit();
        }

        private int GetLastPopulatedRow(Worksheet sheet, int column)
        {
            return sheet.Cells[sheet.Rows.Count, column].End(XlDirection.xlUp).Row;
        }

        private void kiszhFunction()
        {
            //read the values from the numericUpDowns and write them to the excel file
            int startColumn = (int)numericUpDown1.Value;
            int endColumn = (int)numericUpDown2.Value;

            for (int i = 4; i <= GetLastPopulatedRow(worksheet, 3); i++)
            {
                object neptunCodeValue = worksheet.Cells[i, 3].Value2;
                if (neptunCodeValue == null)
                {
                    continue;
                }

                string neptunCode = neptunCodeValue.ToString();
                var values = new List<object>();
                for (int column = startColumn; column <= endColumn; column++)
                {
                    values.Add(worksheet.Cells[i, column].Value2);
                }

                kisZhValuesByNeptunCode[neptunCode] = values;

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

            foreach (var kvp in kisZhValuesByNeptunCode)
            {
                if (destinationRowsByNeptunCode.TryGetValue(kvp.Key, out int destinationRow))
                {
                    List<object> values = kvp.Value;
                    for (int j = 0; j < values.Count; j++)
                    {
                        worksheet2.Cells[destinationRow, 15 + j].Value2 = values[j];
                    }
                }
            }

        }
        private void nagyzhFunction()
        {
            //read the values from the numericUpDowns and write them to the excel file
            int startColumn = (int)numericUpDown3.Value;
            int endColumn = (int)numericUpDown4.Value;
            for (int i = 4; i <= GetLastPopulatedRow(worksheet, 3); i++)
            {
                object neptunCodeValue = worksheet.Cells[i, 3].Value2;
                if (neptunCodeValue == null)
                {
                    continue;
                }

                string neptunCode = neptunCodeValue.ToString();
                var values = new List<object>();
                for (int column = startColumn; column <= endColumn; column++)
                {
                    values.Add(worksheet.Cells[i, column].Value2);
                }

                nagyZhValuesByNeptunCode[neptunCode] = values;

            }
        }
        private void beadandoFunction()
        {
            //read the values from the numericUpDowns and write them to the excel file
            int startColumn = (int)numericUpDown5.Value;
            int endColumn = (int)numericUpDown6.Value;
            for (int i = 4; i <= GetLastPopulatedRow(worksheet, 3); i++)
            {
                object neptunCodeValue = worksheet.Cells[i, 3].Value2;
                if (neptunCodeValue == null)
                {
                    continue;
                }

                string neptunCode = neptunCodeValue.ToString();
                var values = new List<object>();
                for (int column = startColumn; column <= endColumn; column++)
                {
                    values.Add(worksheet.Cells[i, column].Value2);
                }

                beadandoValuesByNeptunCode[neptunCode] = values;

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("Please select an Excel file first.");
                return;
            }
            if (workbook2 == null || worksheet2 == null)
            {
                MessageBox.Show("Please select the output Excel file first.");
                return;
            }
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
            excelApp.DisplayAlerts = false;
            workbook2.SaveAs(@"D:\Letöltések\Monitoring - Python_programozas_ANONIM_2.xlsx");
            workbook2.Close(false);
            excelApp.Quit();
        }
        
    }
}
