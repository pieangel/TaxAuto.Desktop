using TaxAuto.Desktop.Models;

namespace TaxAuto.Desktop.Services
{
    public interface IOcrRunnerService
    {
        Task<OcrRunResult> RunAsync(string inputFile, CancellationToken cancellationToken = default);
    }
}
