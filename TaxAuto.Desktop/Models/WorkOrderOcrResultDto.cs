using System.Text.Json.Serialization;

namespace TaxAuto.Desktop.Models
{
    public class WorkOrderOcrResultDto
    {
        [JsonPropertyName("input_file")]
        public string? InputFile { get; set; }

        [JsonPropertyName("card_count")]
        public int CardCount { get; set; }

        [JsonPropertyName("cards")]
        public List<WorkOrderExcelResultDto> Cards { get; set; } = new();

        [JsonPropertyName("needs_review")]
        public bool NeedsReview { get; set; }
    }
}