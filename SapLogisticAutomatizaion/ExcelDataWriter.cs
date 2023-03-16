using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace SapLogisticAutomatizaion
{
    internal class ExcelDataWriter
    {
        public void WriteToExcel(string path, Product product)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(path);
            Excel.Worksheet worksheet = workbook.Sheets["Sheet1"];

            // Find the last row in the worksheet
            int lastRow = worksheet.Cells.Find("*", System.Reflection.Missing.Value,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            // Get the values from the Product
            string partNumber = product.PartNumber;
            string materialDesc = product.MaterialDescription;
            string serialNumber = product.SerialNumber;
            DateTime manufacturingDate = product.ManufacturingDate;
            DateTime receiptDate = product.ReceiptDate;
            string additionalData = product.AdditionalData;
            string containerCondt = product.ContainerCondition;
            string customsStatus = product.CustomsStatus;
            string containerDetails = product.ContainerDetails;

            // Write the values to the worksheet
            worksheet.Cells[lastRow + 1, 1] = partNumber;
            worksheet.Cells[lastRow + 1, 2] = materialDesc;
            worksheet.Cells[lastRow + 1, 3] = serialNumber;
            worksheet.Cells[lastRow + 1, 4] = manufacturingDate;
            worksheet.Cells[lastRow + 1, 5] = receiptDate;
            worksheet.Cells[lastRow + 1, 6] = additionalData;
            worksheet.Cells[lastRow + 1, 7] = containerCondt;
            worksheet.Cells[lastRow + 1, 8] = customsStatus;
            worksheet.Cells[lastRow + 1, 9] = containerDetails;

            // Save and close the workbook
            workbook.Save();
            workbook.Close();
            excelApp.Quit();
        }
    }
}