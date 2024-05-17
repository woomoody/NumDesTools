namespace NumDesTools
{
    public class FocusLight
    {
        public static string Formula = "=(ROW()=CELL(\"row\"))+(COLUMN()=CELL(\"col\"))";

        private static void AddCondition(dynamic sheet)
        {
            var range = sheet.UsedRange;
            FormatConditions formatConditions = range.FormatConditions;
            int formatCount = 0;
            foreach (object obj in formatConditions)
            {
                if (obj is FormatCondition formatObj)
                {
                    if (formatObj.Formula1 == Formula)
                    {
                        formatCount++;
                    }
                }
            }

            if (formatCount == 0)
            {
                var formatCondition =
                    formatConditions.Add(XlFormatConditionType.xlExpression, Type.Missing, Formula);
                formatCondition.Interior.Color = XlRgbColor.rgbOrange;
            }
        }

        public static void DeleteCondition(dynamic sheet)
        {
            var range = sheet.UsedRange;
            FormatConditions formatConditions = range.FormatConditions;
            foreach (object obj in formatConditions)
            {
                if (obj is FormatCondition formatCondition)
                {
                    if (formatCondition.Formula1 == Formula)
                    {
                        formatCondition.Delete();
                        break;
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