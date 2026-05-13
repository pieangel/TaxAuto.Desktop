using Excel = Microsoft.Office.Interop.Excel;
using TaxAuto.Desktop.Models;

namespace TaxAuto.Desktop.Services
{
    public class PurchaseExcelExporter
    {
        public void Export(string excelPath, List<PurchaseOcrResult> results)
        {
            Excel.Application? excel = null;
            Excel.Workbook? workbook = null;

            try
            {
                excel = new Excel.Application();
                excel.DisplayAlerts = false;

                workbook = excel.Workbooks.Open(excelPath);

                Excel.Worksheet templateSheet = (Excel.Worksheet)workbook.Worksheets[1];

                string sheetName = MakeSafeSheetName(DateTime.Today.ToString("yyyy-MM-dd"), workbook);

                templateSheet.Copy(After: workbook.Worksheets[workbook.Worksheets.Count]);

                Excel.Worksheet newSheet = (Excel.Worksheet)workbook.Worksheets[workbook.Worksheets.Count];
                newSheet.Name = sheetName;

                ClearDataArea(newSheet);
                WriteHeaderDate(newSheet, DateTime.Today);
                WritePurchaseResults(newSheet, results);

                workbook.Save();
            }
            finally
            {
                workbook?.Close(SaveChanges: true);
                excel?.Quit();
            }
        }

        private void WriteHeaderDate(Excel.Worksheet sheet, DateTime date)
        {
            sheet.Range["B3"].Value = date.Year;
            sheet.Range["C3"].Value = $"{date.Month}월 {date.Day}일";
        }

        private void ClearDataArea(Excel.Worksheet sheet)
        {
            sheet.Range["A6:H65"].ClearContents();
        }

        private void WritePurchaseResults(Excel.Worksheet sheet, List<PurchaseOcrResult> results)
        {
            int row = 6;

            foreach (var result in results)
            {
                int startRow = row;

                foreach (var item in result.Items)
                {
                    ((Excel.Range)sheet.Cells[row, 1]).Value2 = DateTime.Today.ToString("M/d");
                    ((Excel.Range)sheet.Cells[row, 2]).Value2 = result.Supplier;
                    ((Excel.Range)sheet.Cells[row, 3]).Value2 = item.ItemNameRaw;
                    ((Excel.Range)sheet.Cells[row, 4]).Value2 = item.Quantity;
                    ((Excel.Range)sheet.Cells[row, 5]).Value2 = "";
                    ((Excel.Range)sheet.Cells[row, 6]).Value2 = item.UnitPrice;
                    ((Excel.Range)sheet.Cells[row, 7]).Value2 = item.Amount;
                    ((Excel.Range)sheet.Cells[row, 8]).Value2 = "";

                    row++;
                }

                int subtotal = result.Items.Sum(x => x.Amount);

                ((Excel.Range)sheet.Cells[row, 1]).Value2 = "";
                ((Excel.Range)sheet.Cells[row, 2]).Value2 = result.Supplier;
                ((Excel.Range)sheet.Cells[row, 3]).Value2 = "소계";
                ((Excel.Range)sheet.Cells[row, 4]).Value2 = "";
                ((Excel.Range)sheet.Cells[row, 5]).Value2 = "";
                ((Excel.Range)sheet.Cells[row, 6]).Value2 = "";
                ((Excel.Range)sheet.Cells[row, 7]).Value2 = subtotal;
                ((Excel.Range)sheet.Cells[row, 8]).Value2 = "";

                // 보기 좋게 소계 행 굵게
                Excel.Range subtotalRange = sheet.Range[$"A{row}:H{row}"];
                subtotalRange.Font.Bold = true;

                row++;

                // 업체 사이 한 줄 띄우기
                row++;
            }
        }

        private string MakeSafeSheetName(string baseName, Excel.Workbook workbook)
        {
            string name = baseName;
            int index = 2;

            while (SheetExists(workbook, name))
            {
                name = $"{baseName} ({index})";
                index++;
            }

            return name;
        }

        private bool SheetExists(Excel.Workbook workbook, string sheetName)
        {
            foreach (object obj in workbook.Worksheets)
            {
                var sheet = (Excel.Worksheet)obj;

                if (sheet.Name == sheetName)
                    return true;
            }

            return false;
        }
    }
}
