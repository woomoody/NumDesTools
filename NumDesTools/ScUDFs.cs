//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using IExcel = Microsoft.Office.Interop.Excel;

using ExcelDna.Integration;

namespace NumDesTools
{
    public class ExcelUdf
    {
        [ExcelFunction(Category = "test", IsVolatile = true, IsMacroType = true, Description = "测试自定义函数")]
        public static double Sum2Num([ExcelArgument(Description = "选个格子")] double a, [ExcelArgument(Description = "选个格子")] double b)
        {
            return a + b;
        }
    }
}