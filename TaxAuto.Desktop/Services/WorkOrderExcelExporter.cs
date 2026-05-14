using TaxAuto.Desktop.Models;

namespace TaxAuto.Desktop.Services
{
    public class WorkOrderExcelExporter
    {
        public void Export(
            string excelPath,
            List<WorkOrderExcelResultDto> results,
            Action<string>? log = null)
        {
            throw new NotSupportedException("작업지시 엑셀 내보내기는 아직 구현되지 않았습니다.");
        }
    }
}