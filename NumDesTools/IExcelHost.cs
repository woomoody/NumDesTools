using System.Runtime.Versioning;

namespace NumDesTools;

[SupportedOSPlatform("windows")]
public interface IExcelHost
{
    object GetActiveWorkbook();
    object GetActiveSheet();
    object GetSelection();
    object GetActiveCell();
    object OpenWorkbook(string filename);
}
