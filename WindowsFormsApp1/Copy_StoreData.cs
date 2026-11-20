using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1
    {
        private void PerformFunction(NumericUpDown startC, NumericUpDown endC, System.Windows.Forms.TextBox point, System.Windows.Forms.CheckBox percent, Func<StudentData, List<object>> targetListSelector)
        {
            int startColumn = (int)startC.Value;
            int endColumn = (int)endC.Value;
            int[] maxPoints = point.Text.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries).Select(int.Parse).ToArray();
            if (maxPoints.Length == 1 && percent.Checked)
            {
                maxPointsGlobal += maxPoints[0] * (((endColumn - startColumn) / 2) + 1);
            }
            else if(maxPoints.Length>1)
            {
                int columnCount = percent.Checked ? ((endColumn - startColumn) / 2) + 1 : (endColumn - startColumn) + 1;
                if (maxPoints.Length != columnCount)
                {
                    MessageBox.Show("Nincs minden oszzlopnak maximális pontszám megadva");
                    return;
                }
                else
                {
                    maxPointsGlobal += maxPoints.Sum();
                }
            }
            else
            {
                maxPointsGlobal += maxPoints[0] * (endColumn - startColumn + 1);
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
    }
}
