using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace TaxAuto.Desktop.Models
{
    public class ValidationDto
    {
        [JsonPropertyName("items_total_amount")]
        public int ItemsTotalAmount { get; set; }

        [JsonPropertyName("document_total_raw")]
        public string? DocumentTotalRaw { get; set; }

        [JsonPropertyName("document_total_amount")]
        public int? DocumentTotalAmount { get; set; }

        [JsonPropertyName("total_amount_check")]
        public bool TotalAmountCheck { get; set; }
    }
}
