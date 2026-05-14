using ClosedXML.Excel;
using TaxAuto.Desktop.Models;
using System.IO;

namespace TaxAuto.Desktop.Services
{
    public class SalesExcelExporter
    {
        public string Export(
            string excelPath,
            List<SalesExcelResultDto> results,
            Action<string>? log = null)
        {
            var outputPath = Path.Combine(
                Path.GetDirectoryName(excelPath)!,
                Path.GetFileNameWithoutExtension(excelPath) + "_신규매출현황.xlsx"
            );

            using var workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add("입력");

            BuildTemplate(sheet);
            WriteResults(sheet, results, log);

            workbook.SaveAs(outputPath);

            log?.Invoke($"신규 매출현황 저장 완료: {outputPath}");

            return outputPath;
        }

        private void BuildTemplate(IXLWorksheet sheet)
        {
            sheet.Style.Font.FontName = "굴림체";
            sheet.Style.Font.FontSize = 9;

            sheet.Cell(1, 2).Value = "일 자";
            sheet.Cell(1, 3).Value = "업 체";
            sheet.Cell(1, 4).Value = "담 당";
            sheet.Cell(1, 5).Value = "저 장 번 호";
            sheet.Cell(1, 6).Value = "매출금액";
            sheet.Cell(1, 7).Value = "비 고";

            var header = sheet.Range(1, 2, 1, 7);
            header.Style.Font.Bold = true;
            header.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            header.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            sheet.Column(2).Width = 10;
            sheet.Column(3).Width = 18;
            sheet.Column(4).Width = 10;
            sheet.Column(5).Width = 42;
            sheet.Column(6).Width = 15;
            sheet.Column(7).Width = 12;

            sheet.Column(6).Style.NumberFormat.Format = "#,##0";
        }

        private void WriteResults(
            IXLWorksheet sheet,
            List<SalesExcelResultDto> results,
            Action<string>? log)
        {
            int row = 2;
            bool wroteToday = false;

            foreach (var result in results)
            {
                foreach (var item in result.Items)
                {
                    if (!wroteToday)
                    {
                        sheet.Cell(row, 1).Value = DateTime.Today.ToString("M/d");
                        wroteToday = true;
                    }

                    sheet.Cell(row, 2).Value = NormalizeDate(result.Date);
                    sheet.Cell(row, 3).Value = result.CustomerName ?? "";
                    sheet.Cell(row, 4).Value = result.Manager ?? "";
                    sheet.Cell(row, 5).Value = item.ItemName ?? item.ItemNameRaw ?? "";
                    sheet.Cell(row, 6).Value = item.CalculatedAmount ?? item.Amount ?? 0;
                    sheet.Cell(row, 7).Value = "수주";

                    var range = sheet.Range(row, 1, row, 7);
                    range.Style.Font.FontName = "굴림체";
                    range.Style.Font.FontSize = 9;
                    range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    sheet.Cell(row, 6).Style.NumberFormat.Format = "#,##0";
                    sheet.Cell(row, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    row++;
                }

                row++;

                log?.Invoke(
                    $"매출 상세 입력: {result.CustomerName} / 품목 {result.Items.Count}개 / 총액 {result.TotalAmount:N0}원");
            }
        }

        private string NormalizeDate(string? raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return "";

            var digits = new string(raw.Where(char.IsDigit).ToArray());

            if (digits.Length >= 8)
            {
                int month = int.Parse(digits.Substring(4, 2));
                int day = int.Parse(digits.Substring(6, 2));

                return $"{month}/{day}";
            }

            return raw;
        }
    }
}