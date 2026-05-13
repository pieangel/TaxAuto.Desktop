using TaxAuto.Desktop.Infrastructure;
using TaxAuto.Desktop.Models;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Input;

namespace TaxAuto.Desktop.ViewModels
{
    public class DocumentOcrTabViewModel : INotifyPropertyChanged
    {
        private readonly DocumentKind _documentKind;

        public ObservableCollection<OcrJobItem> Jobs { get; } = new();

        public ICommand SelectFilesCommand { get; }
        public ICommand RunOcrCommand { get; }
        public ICommand ClearCommand { get; }

        private bool _isRunning;
        public bool IsRunning
        {
            get => _isRunning;
            set
            {
                _isRunning = value;
                OnPropertyChanged();
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
                DocumentKind.Sales => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sales_ocr", "sales_ocr.exe"),

                DocumentKind.Purchase => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "purchase_ocr", "purchase_ocr.exe"),

                DocumentKind.WorkOrder => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "workorder_ocr", "workorder_ocr.exe"),

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
