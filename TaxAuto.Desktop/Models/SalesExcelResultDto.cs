using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace TaxAuto.Desktop.Models
{
    public class SalesExcelResultDto
    {
        [JsonPropertyName("input_file")]
        public string? InputFile { get; set; }

        [JsonPropertyName("date")]
        public string? Date { get; set; }

        [JsonPropertyName("customer_name")]
        public string? CustomerName { get; set; }

        [JsonPropertyName("manager")]
        public string? Manager { get; set; }

        [JsonPropertyName("storage_no")]
        public string? StorageNo { get; set; }

        [JsonPropertyName("job_no")]
        public string? JobNo { get; set; }

        [JsonPropertyName("total_amount")]
        public int? TotalAmount { get; set; }

        [JsonPropertyName("calculated_total")]
        public int? CalculatedTotal { get; set; }

        [JsonPropertyName("discount_amount")]
        public int? DiscountAmount { get; set; }

        [JsonPropertyName("total_check")]
        public bool? TotalCheck { get; set; }

        [JsonPropertyName("needs_review")]
        public bool NeedsReview { get; set; }

        [JsonPropertyName("review_reason")]
        public List<string> ReviewReason { get; set; } = new();

        [JsonPropertyName("items")]
        public List<SalesItemDto> Items { get; set; } = new();
    }
}