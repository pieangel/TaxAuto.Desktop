using System.Text.Json.Serialization;

namespace TaxAuto.Desktop.Models
{
    public class PurchaseOcrResult
    {
        [JsonPropertyName("supplier")]
        public string? Supplier { get; set; }

        [JsonPropertyName("receiver")]
        public string? Receiver { get; set; }

        [JsonPropertyName("date")]
        public string? Date { get; set; }

        [JsonPropertyName("items")]
        public List<PurchaseItemDto> Items { get; set; } = new();

        [JsonPropertyName("validation")]
        public ValidationDto? Validation { get; set; }
    }
}
