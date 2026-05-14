using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace TaxAuto.Desktop.Models
{
    public class SalesItemDto
    {
        [JsonPropertyName("item_name_raw")]
        public string? ItemNameRaw { get; set; }

        [JsonPropertyName("spec_raw")]
        public string? SpecRaw { get; set; }

        [JsonPropertyName("quantity_raw")]
        public string? QuantityRaw { get; set; }

        [JsonPropertyName("unit_price_raw")]
        public string? UnitPriceRaw { get; set; }

        [JsonPropertyName("amount_raw")]
        public string? AmountRaw { get; set; }

        [JsonPropertyName("item_name")]
        public string? ItemName { get; set; }

        [JsonPropertyName("spec")]
        public string? Spec { get; set; }

        [JsonPropertyName("quantity")]
        public int? Quantity { get; set; }

        [JsonPropertyName("unit_price")]
        public int? UnitPrice { get; set; }

        [JsonPropertyName("amount")]
        public int? Amount { get; set; }

        [JsonPropertyName("calculated_amount")]
        public int? CalculatedAmount { get; set; }

        [JsonPropertyName("amount_check")]
        public bool? AmountCheck { get; set; }

        [JsonPropertyName("needs_review")]
        public bool NeedsReview { get; set; }

        [JsonPropertyName("review_reason")]
        public List<string> ReviewReason { get; set; } = new();
    }
}