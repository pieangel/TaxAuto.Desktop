using System.Text.Json.Serialization;

namespace TaxAuto.Desktop.Models
{
    public class WorkOrderExcelResultDto
    {
        [JsonPropertyName("card_index")]
        public int CardIndex { get; set; }

        [JsonPropertyName("job_no")]
        public string? JobNo { get; set; }

        [JsonPropertyName("company")]
        public string? Company { get; set; }

        [JsonPropertyName("work_type")]
        public string? WorkType { get; set; }

        [JsonPropertyName("due_date")]
        public string? DueDate { get; set; }

        [JsonPropertyName("due_weekday")]
        public string? DueWeekday { get; set; }

        [JsonPropertyName("due_time")]
        public string? DueTime { get; set; }

        [JsonPropertyName("manager")]
        public string? Manager { get; set; }

        [JsonPropertyName("materials")]
        public List<string> Materials { get; set; } = new();

        [JsonPropertyName("needs_review")]
        public bool NeedsReview { get; set; }

        [JsonPropertyName("review_reason")]
        public List<string> ReviewReason { get; set; } = new();
    }
}