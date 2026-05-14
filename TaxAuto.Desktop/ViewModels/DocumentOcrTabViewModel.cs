using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using TaxAuto.Desktop.Infrastructure;
using TaxAuto.Desktop.Models;
using TaxAuto.Desktop.Services;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace TaxAuto.Desktop.ViewModels
{
    public class DocumentOcrTabViewModel : INotifyPropertyChanged
    {
        private readonly DocumentKind _documentKind;

        public ObservableCollection<OcrJobItem> Jobs { get; } = new();

        public RelayCommand SelectFilesCommand { get; }
        public RelayCommand RunOcrCommand { get; }
        public RelayCommand ClearCommand { get; }

        public AsyncRelayCommand ExportExcelCommand { get; }


        private bool _isRunning;
        public bool IsRunning
        {
            get => _isRunning;
            set
            {
                _isRunning = value;
                OnPropertyChanged();

                SelectFilesCommand?.RaiseCanExecuteChanged();
                RunOcrCommand?.RaiseCanExecuteChanged();
                ClearCommand?.RaiseCanExecuteChanged();
                ExportExcelCommand?.RaiseCanExecuteChanged();
            }
        }

        private string _logText = "";
        public string LogText
        {
            get => _logText;
            set
            {
                _logText = value;
                OnPropertyChanged();
            }
        }

        public DocumentOcrTabViewModel(DocumentKind documentKind)
        {
            _documentKind = documentKind;

            SelectFilesCommand = new RelayCommand(SelectFiles, () => !IsRunning);
            RunOcrCommand = new RelayCommand(async () => await RunOcrAsync(), () => !IsRunning && Jobs.Count > 0);
            ClearCommand = new RelayCommand(Clear, () => !IsRunning);

            ExportExcelCommand = new AsyncRelayCommand(
                ExportExcelAsync,
                () => !IsRunning && Jobs.Any(x => !string.IsNullOrWhiteSpace(x.ResultJsonPath))
            );
        }


        private async Task ExportExcelAsync_old()
        {
            try
            {
                IsRunning = true;

                var dialog = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "매입 내역을 기록할 엑셀 파일을 선택하세요",
                    Filter = "Excel 파일 (*.xlsx)|*.xlsx|모든 파일 (*.*)|*.*"
                };

                if (dialog.ShowDialog() != true)
                {
                    AppendLog("엑셀 내보내기가 취소되었습니다.");
                    return;
                }

                var loader = new OcrResultLoader();

                var results = Jobs
                    .Where(x => !string.IsNullOrWhiteSpace(x.ResultJsonPath))
                    .Select(x => loader.LoadPurchase(x.ResultJsonPath!))
                    .ToList();

                if (results.Count == 0)
                {
                    AppendLog("엑셀로 내보낼 매입 OCR 결과가 없습니다.");
                    return;
                }

                AppendLog($"엑셀 내보내기 시작: {dialog.FileName}");

                var exporter = new PurchaseExcelExporter();
                exporter.Export(
                    dialog.FileName,
                    results,
                    AppendLog);

                AppendLog($"엑셀 내보내기 완료: {dialog.FileName}");

                //                 System.Windows.MessageBox.Show(
                //                     "엑셀 내보내기가 완료되었습니다.",
                //                     "완료",
                //                     System.Windows.MessageBoxButton.OK,
                //                     System.Windows.MessageBoxImage.Information);

                Process.Start(new ProcessStartInfo
                {
                    FileName = dialog.FileName,
                    UseShellExecute = true
                });

                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                AppendLog($"엑셀 내보내기 오류: {ex.Message}");
            }
            finally
            {
                IsRunning = false;
            }
        }


        private async Task ExportExcelAsync()
        {
            try
            {
                IsRunning = true;

                var dialog = new OpenFileDialog
                {
                    Title = GetExcelDialogTitle(),
                    Filter = "Excel 파일 (*.xlsx)|*.xlsx|모든 파일 (*.*)|*.*"
                };

                if (dialog.ShowDialog() != true)
                {
                    AppendLog("엑셀 내보내기가 취소되었습니다.");
                    return;
                }

                string? outputPath = null;

                switch (_documentKind)
                {
                    case DocumentKind.Purchase:
                        ExportPurchaseExcel(dialog.FileName);
                        outputPath = dialog.FileName;
                        break;

                    case DocumentKind.Sales:
                        outputPath = ExportSalesExcel(dialog.FileName);
                        break;

                    case DocumentKind.WorkOrder:
                        await ExportWorkOrderExcelAsync(dialog.FileName);
                        outputPath = dialog.FileName;
                        break;
                }

                if (!string.IsNullOrWhiteSpace(outputPath))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = outputPath,
                        UseShellExecute = true
                    });
                }

                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                AppendLog($"엑셀 내보내기 오류: {ex.Message}");
            }
            finally
            {
                IsRunning = false;
            }
        }

        private string GetExcelDialogTitle()
        {
            return _documentKind switch
            {
                DocumentKind.Purchase => "매입 내역을 기록할 엑셀 파일을 선택하세요",
                DocumentKind.Sales => "매출 내역을 기록할 엑셀 파일을 선택하세요",
                DocumentKind.WorkOrder => "작업지시 내역을 기록할 엑셀 파일을 선택하세요",
                _ => "엑셀 파일을 선택하세요"
            };
        }

        private void ExportPurchaseExcel(string excelPath)
        {
            var loader = new OcrResultLoader();

            var results = Jobs
                .Where(x => !string.IsNullOrWhiteSpace(x.ResultJsonPath))
                .Select(x => loader.LoadPurchase(x.ResultJsonPath!))
                .ToList();

            if (results.Count == 0)
            {
                AppendLog("엑셀로 내보낼 매입 OCR 결과가 없습니다.");
                return;
            }

            var exporter = new PurchaseExcelExporter();

            exporter.Export(
                excelPath,
                results,
                AppendLog);
        }

        private string? ExportSalesExcel(string excelPath)
        {
            var loader = new OcrResultLoader();

            var results = Jobs
                .Where(x => !string.IsNullOrWhiteSpace(x.ResultJsonPath))
                .Select(x => loader.LoadSales(x.ResultJsonPath!))
                .ToList();

            if (results.Count == 0)
            {
                AppendLog("엑셀로 내보낼 매출 OCR 결과가 없습니다.");
                return null;
            }

            var exporter = new SalesExcelExporter();

            return exporter.Export(
                excelPath,
                results,
                AppendLog);
        }

        private void ExportWorkOrderExcel(string excelPath)
        {
            var loader = new OcrResultLoader();

            var results = Jobs
                .Where(x => !string.IsNullOrWhiteSpace(x.ResultJsonPath))
                .SelectMany(x => loader.LoadWorkOrders(x.ResultJsonPath!))
                .ToList();

            if (results.Count == 0)
            {
                AppendLog("엑셀로 내보낼 작업지시 OCR 결과가 없습니다.");
                return;
            }

            var exporter = new WorkOrderExcelExporter();

            exporter.Export(
                excelPath,
                results,
                AppendLog);
        }


        private async Task ExportWorkOrderExcelAsync(string excelPath)
        {
            var jsonPaths = Jobs
                .Where(x => !string.IsNullOrWhiteSpace(x.ResultJsonPath))
                .Select(x => x.ResultJsonPath!)
                .Where(File.Exists)
                .Distinct()
                .ToList();

            if (jsonPaths.Count == 0)
            {
                AppendLog("엑셀로 내보낼 작업지시 OCR 결과 JSON이 없습니다.");
                return;
            }

            AppendLog($"작업지시 엑셀 내보내기 JSON 개수: {jsonPaths.Count}");

            var exporter = new WorkOrderExcelExporter();

            await exporter.ExportAsync(
                excelPath,
                jsonPaths,
                AppendLog);
        }



        private void SelectFiles()
        {
            var dialog = new OpenFileDialog
            {
                Title = "OCR 처리할 파일 선택",
                Filter = "문서 파일|*.pdf;*.jpg;*.jpeg;*.png|PDF 파일|*.pdf|이미지 파일|*.jpg;*.jpeg;*.png|모든 파일|*.*",
                Multiselect = true
            };

            if (dialog.ShowDialog() != true)
                return;

            foreach (var file in dialog.FileNames)
            {
                if (Jobs.Any(x => x.FilePath.Equals(file, StringComparison.OrdinalIgnoreCase)))
                    continue;

                Jobs.Add(new OcrJobItem(file));
            }

            AppendLog($"파일 {dialog.FileNames.Length}개 선택됨");

            RunOcrCommand.RaiseCanExecuteChanged();
            ClearCommand.RaiseCanExecuteChanged();
        }

        private async Task RunOcrAsync()
        {
            IsRunning = true;

            try
            {
                foreach (var job in Jobs.Where(x => x.Status is OcrJobStatus.Waiting or OcrJobStatus.Failed))
                {
                    await RunSingleJobAsync(job);
                }
            }
            finally
            {
                IsRunning = false;
            }
        }

        private async Task RunSingleJobAsync(OcrJobItem job)
        {
            job.Status = OcrJobStatus.Running;
            job.Message = "OCR 실행 중...";
            job.Elapsed = null;
            job.ResultJsonPath = null;

            AppendLog($"시작: {job.FileName}");

            var sw = Stopwatch.StartNew();

            try
            {
                string exePath = GetOcrExePath();

                AppendLog($"OCR EXE: {exePath}");
                AppendLog($"INPUT: {job.FilePath}");

                if (!File.Exists(exePath))
                    throw new FileNotFoundException("OCR 실행 파일을 찾을 수 없습니다.", exePath);

                var psi = new ProcessStartInfo
                {
                    FileName = exePath,
                    Arguments = $"\"{job.FilePath}\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };

                using var process = new Process
                {
                    StartInfo = psi,
                    EnableRaisingEvents = true
                };

                var stdout = new StringBuilder();
                var stderr = new StringBuilder();

                process.OutputDataReceived += (_, e) =>
                {
                    if (e.Data == null) return;

                    stdout.AppendLine(e.Data);
                    AppendLog(e.Data);

                    var resultPath = TryExtractResultJsonPath(e.Data);
                    if (!string.IsNullOrWhiteSpace(resultPath))
                        job.ResultJsonPath = resultPath;
                };

                process.ErrorDataReceived += (_, e) =>
                {
                    if (e.Data == null) return;

                    stderr.AppendLine(e.Data);
                    AppendLog("[ERR] " + e.Data);
                };

                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();

                await process.WaitForExitAsync();

                sw.Stop();
                job.Elapsed = sw.Elapsed;

                if (process.ExitCode == 0)
                {
                    job.Status = OcrJobStatus.Completed;
                    job.Message = "완료";
                    AppendLog($"완료: {job.FileName} / {job.ElapsedText}");
                }
                else
                {
                    job.Status = OcrJobStatus.Failed;
                    job.Message = $"실패 ExitCode={process.ExitCode}";
                    AppendLog($"실패: {job.FileName}");
                }
            }
            catch (Exception ex)
            {
                sw.Stop();

                job.Elapsed = sw.Elapsed;
                job.Status = OcrJobStatus.Failed;
                job.Message = ex.Message;

                AppendLog($"예외: {ex.Message}");
            }
        }

        private string GetOcrExePath()
        {
            return _documentKind switch
            {
                DocumentKind.Sales => Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    "tools",
                    "sales_ocr",
                    "sales_ocr.exe"
                ),

                DocumentKind.Purchase => Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    "tools",
                    "purchase_ocr",
                    "purchase_ocr.exe"
                ),

                DocumentKind.WorkOrder => Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    "tools",
                    "workorder_ocr",
                    "workorder_ocr.exe"
                ),

                _ => throw new NotSupportedException("지원하지 않는 문서 타입입니다.")
            };
        }

        private static string? TryExtractResultJsonPath(string line)
        {
            const string prefix = "RESULT_JSON=";

            if (!line.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                return null;

            return line[prefix.Length..].Trim();
        }

        private void Clear()
        {
            Jobs.Clear();
            LogText = "";

            RunOcrCommand.RaiseCanExecuteChanged();
            ClearCommand.RaiseCanExecuteChanged();
        }

        private void AppendLog(string message)
        {
            App.Current.Dispatcher.Invoke(() =>
            {
                LogText += $"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}";
            });
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
