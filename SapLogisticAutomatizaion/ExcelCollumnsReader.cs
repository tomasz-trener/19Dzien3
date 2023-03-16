using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace SapLogisticAutomatizaion
{
    internal class ExcelCollumnsReader
    {
        public string[][] ReadExcelFile(string path)
        {
            // Create an instance of Excel application
            Excel.Application excelApp = new Excel.Application();

            try
            {
                // Open the Excel file
                Excel.Workbook workbook = excelApp.Workbooks.Open(path);

                // Select the first worksheet
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

                // Find the last row and column with data
                Excel.Range last = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastRow = last.Row;
                int lastCol = last.Column;

                // Create a jagged array to store the data
                string[][] data = new string[lastRow][];
                for (int i = 0; i < lastRow; i++)
                {
                    data[i] = new string[2];
                }

                // Read data from columns A and B
                for (int i = 1; i <= lastRow; i++)
                {
                    data[i - 1][0] = ((Excel.Range)worksheet.Cells[i, 1]).Value2.ToString();
                    data[i - 1][1] = ((Excel.Range)worksheet.Cells[i, 2]).Value2.ToString();
                }

                // Close the Excel file
                workbook.Close(false);
                return data;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                // Quit Excel application
                excelApp.Quit();
            }
        }
    }
}