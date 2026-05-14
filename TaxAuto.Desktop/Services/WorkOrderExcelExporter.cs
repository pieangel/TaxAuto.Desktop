using ClosedXML.Excel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using TaxAuto.Desktop.Models;

namespace TaxAuto.Desktop.Services
{
    public class WorkOrderExcelExporter
    {
        private const string SheetNameFallback = "작업기록";

        public async Task ExportAsync(
            string excelPath,
            List<string> resultJsonPaths,
            Action<string>? log = null)
        {
            if (!File.Exists(excelPath))
                throw new FileNotFoundException("엑셀 파일을 찾지 못했습니다.", excelPath);

            if (resultJsonPaths == null || resultJsonPaths.Count == 0)
            {
                log?.Invoke("작업지시 엑셀 내보내기: OCR 결과 JSON이 없습니다.");
                return;
            }

            string exePath = Path.Combine(
                AppContext.BaseDirectory,
                "tools",
                "workorder_excel_exporter",
                "workorder_excel_exporter.exe");

            if (!File.Exists(exePath))
                throw new FileNotFoundException("작업지시 엑셀 내보내기 EXE를 찾지 못했습니다.", exePath);

            var args = new StringBuilder();
            args.Append($"\"{excelPath}\" ");

            foreach (var jsonPath in resultJsonPaths)
            {
                if (!File.Exists(jsonPath))
                {
                    log?.Invoke($"JSON 파일 없음. 제외: {jsonPath}");
                    continue;
                }

                args.Append($"\"{jsonPath}\" ");
            }

            log?.Invoke($"Excel Export EXE: {exePath}");
            log?.Invoke($"EXCEL: {excelPath}");

            var psi = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = args.ToString(),
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
            };

            using var process = new Process();
            process.StartInfo = psi;

            process.OutputDataReceived += (_, e) =>
            {
                if (!string.IsNullOrWhiteSpace(e.Data))
                    log?.Invoke(e.Data);
            };

            process.ErrorDataReceived += (_, e) =>
            {
                if (!string.IsNullOrWhiteSpace(e.Data))
                    log?.Invoke("[ERR] " + e.Data);
            };

            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            await process.WaitForExitAsync();

            if (process.ExitCode != 0)
                throw new Exception($"작업지시 엑셀 내보내기 실패. ExitCode={process.ExitCode}");

            log?.Invoke("작업지시 엑셀 내보내기 완료");
        }
    

        public void Export(
            string excelPath,
            List<WorkOrderExcelResultDto> results,
            Action<string>? log = null)
        {
            if (!File.Exists(excelPath))
                throw new FileNotFoundException("엑셀 파일을 찾지 못했습니다.", excelPath);

            if (results == null || results.Count == 0)
            {
                log?.Invoke("작업지시 엑셀 내보내기: 결과가 없습니다.");
                return;
            }

            log?.Invoke($"엑셀 파일 열기: {excelPath}");
            log?.Invoke($"작업지시 결과 개수: {results.Count}");

            BackupExcel(excelPath, log);

            using var workbook = new XLWorkbook(excelPath);

            var sheet = FindCurrentMonthSheet(workbook, log);

            foreach (var result in results)
            {
                WriteWorkOrder(sheet, result, log);
            }

            workbook.Save();

            log?.Invoke($"작업지시 엑셀 저장 완료: {excelPath}");
        }

        private IXLWorksheet FindCurrentMonthSheet(
            XLWorkbook workbook,
            Action<string>? log)
        {
            string currentMonthText = $"{DateTime.Today.Month}월";

            var sheet = workbook.Worksheets
                .FirstOrDefault(x => x.Name.Contains(currentMonthText));

            if (sheet != null)
            {
                log?.Invoke($"현재 월 Sheet 선택: {sheet.Name}");
                return sheet;
            }

            sheet = workbook.Worksheets
                .FirstOrDefault(x => x.Name == SheetNameFallback);

            if (sheet != null)
            {
                log?.Invoke($"Fallback Sheet 선택: {sheet.Name}");
                return sheet;
            }

            sheet = workbook.Worksheets.First();

            log?.Invoke($"첫 번째 Sheet 선택: {sheet.Name}");
            return sheet;
        }

        private void WriteWorkOrder(
            IXLWorksheet sheet,
            WorkOrderExcelResultDto result,
            Action<string>? log)
        {
            if (string.IsNullOrWhiteSpace(result.JobNo))
            {
                log?.Invoke("접수번호가 없어 건너뜁니다.");
                return;
            }

            int targetRow = FindTargetRow(sheet, result.JobNo)
                ?? GetAppendRow(sheet);

            log?.Invoke($"작업지시 기록: {result.JobNo} -> Row {targetRow}");

            // A 접수번호
            // B 구역
            // C 업체명
            // E 납기일
            // F 시간
            // K 호기/담당자
            // U 재질 목록

            sheet.Cell(targetRow, 1).Value = result.JobNo;
            sheet.Cell(targetRow, 2).Value = result.WorkType ?? "";
            sheet.Cell(targetRow, 3).Value = result.Company ?? "";
            sheet.Cell(targetRow, 5).Value = result.DueDate ?? "";
            sheet.Cell(targetRow, 6).Value = result.DueTime ?? "";
            sheet.Cell(targetRow, 11).Value = result.Manager ?? "";
            sheet.Cell(targetRow, 21).Value = string.Join(", ", result.Materials ?? new List<string>());

            if (result.NeedsReview)
            {
                log?.Invoke(
                    $"검토 필요: {result.JobNo} / {string.Join(", ", result.ReviewReason ?? new List<string>())}");
            }
        }

        private int? FindTargetRow(
            IXLWorksheet sheet,
            string ocrJobNo)
        {
            string ocrBase = GetBaseJobNo(ocrJobNo);

            int lastRow = sheet.LastRowUsed()?.RowNumber() ?? 1;

            for (int row = 2; row <= lastRow; row++)
            {
                string excelNo = sheet.Cell(row, 1).GetString().Trim();

                if (string.IsNullOrWhiteSpace(excelNo))
                    continue;

                string excelBase = GetBaseJobNo(excelNo);

                if (ocrJobNo.StartsWith(excelNo, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(ocrBase, excelBase, StringComparison.OrdinalIgnoreCase))
                {
                    return row;
                }
            }

            return null;
        }

        private string GetBaseJobNo(string jobNo)
        {
            if (string.IsNullOrWhiteSpace(jobNo))
                return "";

            var match = Regex.Match(
                jobNo.Trim(),
                @"^([A-Za-z]+\d+)",
                RegexOptions.IgnoreCase);

            return match.Success
                ? match.Groups[1].Value.ToUpperInvariant()
                : jobNo.Trim().ToUpperInvariant();
        }

        private int GetAppendRow(IXLWorksheet sheet)
        {
            int lastRow = sheet.LastRowUsed()?.RowNumber() ?? 1;
            int lastDataRow = 1;

            for (int row = 2; row <= lastRow; row++)
            {
                string jobNo = sheet.Cell(row, 1).GetString().Trim();

                if (!string.IsNullOrWhiteSpace(jobNo))
                    lastDataRow = row;
            }

            return lastDataRow + 1;
        }

        private void BackupExcel(
            string excelPath,
            Action<string>? log)
        {
            string dir = Path.Combine(
                Path.GetDirectoryName(excelPath)!,
                "excel_backups");

            Directory.CreateDirectory(dir);

            string fileName = Path.GetFileNameWithoutExtension(excelPath);
            string ext = Path.GetExtension(excelPath);
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            string backupPath = Path.Combine(
                dir,
                $"{fileName}_{timestamp}{ext}");

            File.Copy(excelPath, backupPath, overwrite: true);

            log?.Invoke($"엑셀 백업 생성: {backupPath}");
        }
    }
}