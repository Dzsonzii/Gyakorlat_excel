using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
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
        string filePath2 = "";
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
            //gombok eltuntetese, megjelenitese
            numericUpDown1.Visible = !numericUpDown1.Visible;
            numericUpDown2.Visible = !numericUpDown2.Visible;
            label1.Visible = !label1.Visible;
            label4.Visible = !label4.Visible;
            numericUpDown7.Visible = !numericUpDown7.Visible;
            kiszh = !kiszh;
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            //gombok eltuntetese, megjelenitese
            numericUpDown5.Visible = !numericUpDown5.Visible;
            numericUpDown6.Visible = !numericUpDown6.Visible;
            label3.Visible = !label3.Visible;
            label6.Visible = !label6.Visible;
            numericUpDown9.Visible = !numericUpDown9.Visible;
            beadando = !beadando;


        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            
            numericUpDown3.Visible = !numericUpDown3.Visible;
            numericUpDown4.Visible = !numericUpDown4.Visible;
            label2.Visible = !label2.Visible;
            label5.Visible = !label5.Visible;
            numericUpDown8.Visible = !numericUpDown8.Visible;
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
            }
            MessageBox.Show("Az adatok beolvasása megtörtént!");
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
            }
            MessageBox.Show("Az adatok beolvasása megtörtént!");
        }
        private void readInitialExcelFile(string filePath)
        {
            //fajlbol beolvasom a munkalapot
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            workbook = excelApp.Workbooks.Open(filePath);
            worksheet = workbook.Sheets[1];

;
        }
        private void readEndExcelFile(string filePath2)
        {
            //fajlbol beolvasom a munkalapot
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            workbook2 = excelApp.Workbooks.Open(filePath2);
            worksheet2 = workbook2.Sheets[1];


        }

        private int GetLastPopulatedRow(Worksheet sheet, int column)
        {
            return sheet.Cells[sheet.Rows.Count, column].End(XlDirection.xlUp).Row;
        }

        private void copyFunction(Dictionary<string, List<object>> sourceValuesByNeptunCode, int startColumn, int endColumn, int maxPoints, bool isNagyZh = false)
        {
            //beolvasom az oszlop szamait, majd feldolgozom az ertekeket


            for (int i = 4; i <= GetLastPopulatedRow(worksheet, 3); i++)
            {
                object neptunCodeValue = worksheet.Cells[i, 3].Value2;
                if (neptunCodeValue == null)
                {
                    continue;
                }
                //elmentem a neptun kodot, hogyha mas sorrendbe lenne a kimenetbe, akkor is meg lehessen talalni mindent
                string neptunCode = neptunCodeValue.ToString();
                var values = new List<object>();
                if (isNagyZh)
                {
                    for (int j = startColumn; j <= endColumn; j=j+2)
                    {
                        values.Add($"{worksheet.Cells[i, j].Value2}/{maxPoints}");
                    }
                    sourceValuesByNeptunCode[neptunCode] = values;
                }
                else
                {
                    for (int j = startColumn; j <= endColumn; j++)
                    {
                        values.Add($"{worksheet.Cells[i, j].Value2}/{maxPoints}");
                    }
                    sourceValuesByNeptunCode[neptunCode] = values;
                }

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
            foreach (var kvp in sourceValuesByNeptunCode)
            {
                if (destinationRowsByNeptunCode.TryGetValue(kvp.Key, out int destinationRow))
                {
                    List<object> values = kvp.Value;
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
            //beolvasom az oszlop szamait, majd feldolgozom az ertekeket
            int startColumn = (int)numericUpDown1.Value;
            int endColumn = (int)numericUpDown2.Value;
            int maxPoints = (int)numericUpDown7.Value;
            copyFunction(kisZhValuesByNeptunCode, startColumn, endColumn, maxPoints);
        }
            
        
        private void nagyzhFunction()
        {
            //same as kiszh
            int startColumn = (int)numericUpDown4.Value;
            int endColumn = (int)numericUpDown3.Value;
            int maxPoints = (int)numericUpDown8.Value;
            copyFunction(nagyZhValuesByNeptunCode, startColumn, endColumn, maxPoints, true);
        }
        private void beadandoFunction()
        {
            //same as kiszh
            int startColumn = (int)numericUpDown6.Value;
            int endColumn = (int)numericUpDown5.Value;
            int maxPoints = (int)numericUpDown9.Value;
            copyFunction(beadandoValuesByNeptunCode, startColumn, endColumn, maxPoints);


        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(filePath) || string.IsNullOrEmpty(filePath2))
            {
                MessageBox.Show("Kérlek töltsd be mind a két filet!");
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
            workbook.Close(false);
            excelApp.Quit();
            MessageBox.Show("Az adatok írása megtörtént!");
        }
        
    }
}
