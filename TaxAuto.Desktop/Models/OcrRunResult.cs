using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaxAuto.Desktop.Models
{
    public class OcrRunResult
    {
        public string InputFile { get; set; } = "";
        public int ExitCode { get; set; }
        public string StandardOutput { get; set; } = "";
        public string StandardError { get; set; } = "";
        public string FullOutput { get; set; } = "";
        public string? ResultJsonPath { get; set; }

        public bool Success => ExitCode == 0;
    }
}
