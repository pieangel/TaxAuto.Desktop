using System.Diagnostics;
using System.IO;
using System.Text;
using TaxAuto.Desktop.Models;

namespace TaxAuto.Desktop.Services
{
    public class OcrRunnerService : IOcrRunnerService
    {
        public async Task<OcrRunResult> RunAsync(string inputFile, CancellationToken cancellationToken = default)
        {
            var exePath = Path.Combine(
                AppContext.BaseDirectory,
                "tools",
                "sales_ocr",
                "sales_ocr.exe"
            );

            if (!File.Exists(exePath))
                throw new FileNotFoundException("sales_ocr.exe를 찾을 수 없습니다.", exePath);

            var psi = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = $"\"{inputFile}\"",
                WorkingDirectory = Path.GetDirectoryName(exePath)!,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            using var process = new Process { StartInfo = psi };

            process.Start();

            var stdoutTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
            var stderrTask = process.StandardError.ReadToEndAsync(cancellationToken);

            await process.WaitForExitAsync(cancellationToken);

            var stdout = await stdoutTask;
            var stderr = await stderrTask;

            return new OcrRunResult
            {
                InputFile = inputFile,
                ExitCode = process.ExitCode,
                StandardOutput = stdout,
                StandardError = stderr,
                FullOutput = stdout + Environment.NewLine + stderr,
                ResultJsonPath = ExtractResultJsonPath(stdout)
            };
        }

        private static string? ExtractResultJsonPath(string stdout)
        {
            foreach (var line in stdout.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n'))
            {
                if (line.StartsWith("RESULT_JSON="))
                    return line["RESULT_JSON=".Length..].Trim();
            }

            return null;
        }
    }
}
