using ClosedXML.Excel;
using TaxAuto.Desktop.Models;

namespace TaxAuto.Desktop.Services
{
    public class PurchaseExcelExporter
    {
        public void Export(
            string excelPath,
            List<PurchaseOcrResult> results,
            Action<string>? log = null)
        {
            log?.Invoke($"엑셀 파일 열기: {excelPath}");
            log?.Invoke($"OCR 결과 개수: {results.Count}");

            using var workbook = new XLWorkbook(excelPath);

            var templateSheet = workbook.Worksheets
                .FirstOrDefault(x => x.Name.StartsWith("날짜"))
                ?? workbook.Worksheets.First();

            log?.Invoke($"템플릿 Sheet: {templateSheet.Name}");

            string sheetName = MakeSafeSheetName(
                DateTime.Today.ToString("yyyy-MM-dd"),
                workbook);

            var sheet = templateSheet.CopyTo(sheetName);
            log?.Invoke($"새 Sheet 생성: {sheetName}");

            ClearDataArea(sheet);
            WriteHeaderDate(sheet, DateTime.Today);
            WritePurchaseResults(sheet, results, log);

            workbook.SaveAs(excelPath);
            log?.Invoke($"엑셀 저장 완료: {excelPath}");
        }

        private void WriteHeaderDate(IXLWorksheet sheet, DateTime date)
        {
            sheet.Cell("B3").Value = date.Year;
            sheet.Cell("C3").Value = date.Month;
            sheet.Cell("D3").Value = date.Day;
        }

        private void ClearDataArea(IXLWorksheet sheet)
        {
            var range = sheet.Range("A6:I200");

            range.Unmerge();

            foreach (var cell in range.Cells())
            {
                cell.FormulaA1 = null;
                cell.Value = "";
            }

            range.Style.Fill.BackgroundColor = XLColor.NoColor;
            range.Style.Font.Bold = false;

            // 기존 소계 행의 2줄 선 제거
            range.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            range.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            range.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        private void WritePurchaseResults(
            IXLWorksheet sheet,
            List<PurchaseOcrResult> results,
            Action<string>? log)
        {
            int row = 6;

            foreach (var result in results)
            {
                string supplier = result.Supplier ?? "(업체명 확인 필요)";
                string dateText = GetDisplayDate(result.Date);

                int startRow = row;

                foreach (var item in result.Items)
                {
                    WriteItemRow(sheet, row, item);
                    row++;
                }

                int endItemRow = row - 1;

                if (endItemRow >= startRow)
                {
                    MergeVendorCells(
                        sheet,
                        startRow,
                        endItemRow,
                        dateText,
                        supplier,
                        result.NeedsReview ? "검토 필요" : "");
                }

                // 여기 row가 바로 소계 행
                WriteSubtotalRow(sheet, row, result);
                row++;
            }

            log?.Invoke($"최종 기록 종료 행: {row}");
        }

        private void WriteItemRow(
    IXLWorksheet sheet,
    int row,
    PurchaseItemDto item)
        {
            sheet.Cell(row, 3).Value = item.ItemNameRaw ?? "";
            sheet.Cell(row, 4).Value = item.Quantity ?? 0;
            sheet.Cell(row, 5).Value = "";
            sheet.Cell(row, 6).Value = item.UnitPrice ?? 0;
            sheet.Cell(row, 7).Value = item.Amount ?? 0;

            // I열은 품목행에서는 반드시 비움
            sheet.Cell(row, 9).FormulaA1 = null;
            sheet.Cell(row, 9).Value = "";

            var range = sheet.Range(row, 1, row, 9);

            // 스타일 초기화
            range.Style.Fill.BackgroundColor = XLColor.NoColor;
            range.Style.Font.Bold = false;

            // 정렬
            range.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            // ★ 핵심: 기존 소계의 Double Border 제거
            range.Style.Border.TopBorder =
                XLBorderStyleValues.Thin;

            range.Style.Border.BottomBorder =
                XLBorderStyleValues.Thin;

            range.Style.Border.LeftBorder =
                XLBorderStyleValues.Thin;

            range.Style.Border.RightBorder =
                XLBorderStyleValues.Thin;

            range.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            range.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            // 숫자 포맷
            sheet.Cell(row, 6).Style.NumberFormat.Format = "#,##0";
            sheet.Cell(row, 7).Style.NumberFormat.Format = "#,##0";
        }


        private void MergeVendorCells(
            IXLWorksheet sheet,
            int startRow,
            int endRow,
            string dateText,
            string supplier,
            string memo)
        {
            var dateRange = sheet.Range(startRow, 1, endRow, 1);
            dateRange.Merge();
            dateRange.Value = dateText;

            var supplierRange = sheet.Range(startRow, 2, endRow, 2);
            supplierRange.Merge();
            supplierRange.Value = supplier;

            var memoRange = sheet.Range(startRow, 8, endRow, 8);
            memoRange.Merge();
            memoRange.Value = memo;

            dateRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            supplierRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            memoRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            dateRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            supplierRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            memoRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        private void WriteSubtotalRow(
            IXLWorksheet sheet,
            int row,
            PurchaseOcrResult result)
        {
            int subtotal = result.Items.Sum(x => x.Amount ?? 0);

            // 기존 병합/수식/값 정리
            var rowRange = sheet.Range(row, 1, row, 9);
            rowRange.Unmerge();

            foreach (var cell in rowRange.Cells())
            {
                cell.FormulaA1 = null;
                cell.Value = "";
            }

            // A~C 병합 후 소계 입력
            var labelRange = sheet.Range(row, 1, row, 3);
            labelRange.Merge();
            labelRange.Value = "소    계";

            // 금액 입력
            sheet.Cell(row, 7).Value = subtotal; // G열
            sheet.Cell(row, 9).Value = subtotal; // I열

            // 기본 스타일
            rowRange.Style.Fill.BackgroundColor = XLColor.LightGray;
            rowRange.Style.Font.Bold = true;
            rowRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            labelRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            sheet.Cell(row, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            sheet.Cell(row, 7).Style.NumberFormat.Format = "#,##0";

            sheet.Cell(row, 9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            sheet.Cell(row, 9).Style.NumberFormat.Format = "#,##0";

            // 테두리: 먼저 기본 얇은 선
            rowRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            rowRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // 마지막에 아래쪽만 두 줄
            rowRange.Style.Border.BottomBorder = XLBorderStyleValues.Double;
        }


        private string GetDisplayDate(string? dateText)
        {
            if (DateTime.TryParse(dateText, out var date))
                return date.ToString("M/d");

            return DateTime.Today.ToString("M/d");
        }

        private string MakeSafeSheetName(string baseName, XLWorkbook workbook)
        {
            string name = baseName;
            int index = 2;

            while (workbook.Worksheets.Any(x => x.Name == name))
            {
                name = $"{baseName} ({index})";
                index++;
            }

            return name;
        }
    }
}
