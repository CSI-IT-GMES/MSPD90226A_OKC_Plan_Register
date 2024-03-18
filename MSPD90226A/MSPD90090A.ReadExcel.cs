using System;
using System.Data;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace CSI.GMES.PD
{
    internal class ReadExcel
    {
        public DataTable GetData(string argPathOfExcelFile)
        {
            DataTable dt = new DataTable();

            try
            {
                Excel.Application excelApp = new Excel.Application();

                excelApp.DisplayAlerts = false; //Don't want Excel to display error messageboxes

                Excel.Workbook workbook = excelApp.Workbooks.Open(argPathOfExcelFile); //This opens the file
                                                                                       //  workbook.ActiveSheet

                string sheetName = workbook.ActiveSheet.Name;
                Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets.get_Item(sheetName); //Get the first sheet in the file

                int lastRow = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                int lastColumn = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                Excel.Range c1 = sheet.Cells[1, 1];
                Excel.Range c2 = sheet.Cells[lastRow, lastColumn];

                Excel.Range oRange = sheet.get_Range(c1, c2);

                oRange.EntireColumn.AutoFit();

                object[,] cellValues = (object[,])oRange.Value2;
                object[] values = new object[lastColumn];

                for (int i = 0; i < oRange.Columns.Count; i++)
                {
                    dt.Columns.Add("a" + i.ToString());
                }

                for (int i = 1; i <= lastRow; i++)
                {
                    if (cellValues[i, 1] == null) continue;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        values[j] = cellValues[i, j + 1];
                    }
                    dt.Rows.Add(values);
                }

                workbook.Close(false);
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            return dt;
        }

    }
}
