using System.Text.Json.Serialization;

namespace TaxAuto.Desktop.Models
{
    public class PurchaseOcrResult
    {
        [JsonPropertyName("document_type")]
        public string? DocumentType { get; set; }

        [JsonPropertyName("supplier")]
        public string? Supplier { get; set; }

        [JsonPropertyName("receiver")]
        public string? Receiver { get; set; }

        [JsonPropertyName("date")]
        public string? Date { get; set; }

        [JsonPropertyName("items")]
        public List<PurchaseItemDto> Items { get; set; } = new();

        [JsonPropertyName("total_amount")]
        public int? TotalAmount { get; set; }

        [JsonPropertyName("validation")]
        public ValidationDto? Validation { get; set; }

        [JsonPropertyName("needs_review")]
        public bool NeedsReview { get; set; }
    }
}
