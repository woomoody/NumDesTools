using System;
using Excel = Microsoft.Office.Interop.Excel;
// ReSharper disable All


namespace NumDesTools
{
    public class FocusLight
    {
        public static string Formula = "=(ROW()=CELL(\"row\"))+(COLUMN()=CELL(\"col\"))";
        private static void AddCondition(dynamic sheet)
        {
            var range = sheet.UsedRange;
            Excel.FormatConditions formatConditions = range.FormatConditions;
            int formatCount = 0;
            //存在就不新增了
            foreach (object obj in formatConditions)
            {
                Excel.FormatCondition formatObj = obj as Excel.FormatCondition;
                if (formatObj != null)
                {
                    if (formatObj.Formula1 == Formula)
                    {
                        formatCount++;
                    }
                }
            }
            if (formatCount == 0)
            {
                //设置新条件格式
                var formatCondition = formatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, Formula);
                formatCondition.Interior.Color = Excel.XlRgbColor.rgbOrange;
            }
        }
        public static void DeleteCondition(dynamic sheet)
        {
            var range = sheet.UsedRange;
            Excel.FormatConditions formatConditions = range.FormatConditions;
            // 循环遍历条件格式规则，找到指定条件格式并清除
            foreach (object obj in formatConditions)
            {
                Excel.FormatCondition formatCondition = obj as Excel.FormatCondition;
                if (formatCondition != null)
                {
                    if (formatCondition.Formula1 == Formula)
                    {
                        formatCondition.Delete(); // 清除指定条件格式
                        break; // 可以选择中断循环，因为找到了要清除的条件格式
                    }
                }
            }
        }
        public static void Calculate()
        {
            var sheet = NumDesAddIn.App.ActiveSheet;
            AddCondition(sheet);
            if (NumDesAddIn.FocusLabelText == "聚光灯：关闭")
            {
                DeleteCondition(sheet);
            }
            sheet.Calculate();
        }
    }
}
