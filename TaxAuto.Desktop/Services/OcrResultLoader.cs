using System.Text.Json;
using System.IO;
using TaxAuto.Desktop.Models;

namespace TaxAuto.Desktop.Services
{
    public class OcrResultLoader
    {
        public PurchaseOcrResult LoadPurchase(string jsonPath)
        {
            var json = File.ReadAllText(jsonPath);

            return JsonSerializer.Deserialize<PurchaseOcrResult>(json)!
                   ?? throw new Exception("JSON deserialize 실패");
        }
    }
}
