using ClosedXML.Excel;
using TaxAuto.Desktop.Models;
using System.IO;

namespace TaxAuto.Desktop.Services
{
    public class SalesExcelExporter
    {
        public void Export(
    string excelPath,
    List<SalesExcelResultDto> results,
    Action<string>? log = null)
        {
            try
            {
                log?.Invoke($"엑셀 파일 열기: {excelPath}");
                log?.Invoke($"매출 결과 개수: {results.Count}");

                using var workbook = new XLWorkbook(excelPath);

                var sheet = workbook.Worksheet("입력");

                int row = FindNextRow(sheet);
                log?.Invoke($"매출 입력 시작 행: {row}");

                foreach (var result in results)
                {
                    try
                    {
                        //sheet.Cell(row, 2).Value = NormalizeDate(result.Date);      // B: 일자
                        sheet.Cell(row, 3).Value = result.CustomerName ?? "";       // C: 업체

                        //CopyManagerFormula(sheet, row);                             // D: 담당

                        sheet.Cell(row, 5).Value = result.StorageNo ?? "";          // E: 저장 번호
                        sheet.Cell(row, 6).Value = result.TotalAmount ?? 0;          // F: 매출금액
                        sheet.Cell(row, 7).Value = "수주";                           // G: 비고

                        log?.Invoke($"매출 입력: row={row}, {result.CustomerName} / {result.TotalAmount:N0}원");

                        row++;
                    }
                    catch (Exception ex)
                    {
                        log?.Invoke($"매출 행 입력 오류: row={row}, 업체={result.CustomerName}, 오류={ex}");
                        throw;
                    }
                }

                var outputPath = Path.Combine(
                    Path.GetDirectoryName(excelPath)!,
                    Path.GetFileNameWithoutExtension(excelPath) + "_매출입력완료.xlsx"
                );

                log?.Invoke($"엑셀 저장 시도: {outputPath}");
                RemoveAllAutoFilters(workbook, log);

                workbook.SaveAs(outputPath);

                log?.Invoke($"엑셀 저장 완료: {outputPath}");
            }
            catch (Exception ex)
            {
                log?.Invoke($"SalesExcelExporter.Export 오류: {ex}");
                throw;
            }
        }

        private void RemoveAllAutoFilters(XLWorkbook workbook, Action<string>? log = null)
        {
            foreach (var ws in workbook.Worksheets)
            {
                try
                {
                    ws.AutoFilter.Clear();
                    ws.RangeUsed()?.SetAutoFilter(false);

                    foreach (var table in ws.Tables)
                    {
                        table.ShowAutoFilter = false;
                    }

                    log?.Invoke($"AutoFilter 제거: {ws.Name}");
                }
                catch (Exception ex)
                {
                    log?.Invoke($"AutoFilter 제거 실패: {ws.Name} / {ex.Message}");
                }
            }
        }

        private int FindNextRow(IXLWorksheet sheet)
        {
            int row = 4;

            while (true)
            {
                var customer = sheet.Cell(row, 3).GetString();

                if (string.IsNullOrWhiteSpace(customer))
                    return row;

                row++;
            }
        }

        private void CopyManagerFormula(IXLWorksheet sheet, int row)
        {
            if (row <= 4)
                return;

            var source = sheet.Cell(row - 1, 4);
            var target = sheet.Cell(row, 4);

            if (!string.IsNullOrWhiteSpace(source.FormulaA1))
                target.FormulaA1 = source.FormulaA1;
        }

        private string NormalizeDate(string? raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return "";

            var digits = new string(raw.Where(char.IsDigit).ToArray());

            if (digits.Length >= 8)
                return $"{digits[..4]}-{digits.Substring(4, 2)}-{digits.Substring(6, 2)}";

            return raw;
        }
    }
}