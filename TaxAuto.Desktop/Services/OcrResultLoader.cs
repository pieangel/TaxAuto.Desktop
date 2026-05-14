using System.Text.Json;
using System.IO;
using TaxAuto.Desktop.Models;
using System.Text;

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

        public SalesExcelResultDto LoadSales(string jsonPath)
        {
            var json = File.ReadAllText(jsonPath, Encoding.UTF8);

            var result = JsonSerializer.Deserialize<SalesExcelResultDto>(json);

            if (result == null)
                throw new InvalidOperationException("매출 OCR 결과 JSON을 읽지 못했습니다.");

            return result;
        }

        public WorkOrderExcelResultDto LoadWorkOrder(string jsonPath)
        {
            throw new NotSupportedException("작업지시 OCR 결과 로더는 아직 구현되지 않았습니다.");
        }

        public List<WorkOrderExcelResultDto> LoadWorkOrders(string jsonPath)
        {
            if (!File.Exists(jsonPath))
                throw new FileNotFoundException("작업지시 OCR JSON 파일을 찾지 못했습니다.", jsonPath);

            string json = File.ReadAllText(jsonPath);

            var result = JsonSerializer.Deserialize<WorkOrderOcrResultDto>(
                json,
                new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });

            if (result == null)
                throw new InvalidOperationException("작업지시 OCR JSON 파싱에 실패했습니다.");

            return result.Cards ?? new List<WorkOrderExcelResultDto>();
        }
    }
}
