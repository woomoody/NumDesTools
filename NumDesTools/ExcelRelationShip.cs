namespace NumDesTools;
/// <summary>
/// 废弃的Excel关系图
/// </summary>
internal class ExcelRelationShip
{
    //private static readonly dynamic App = ExcelDnaUtil.Application;
    //public static Dictionary<string, List<string>> ExcelLinkDictionary;
    //public static Dictionary<string, List<int>> ExcelFixKeyDictionary;
    //public static Dictionary<string, List<string>> ExcelFixKeyMethodDictionary;

    //public static void StartExcelData()
    //{
    //    var indexWk = App.ActiveWorkbook;
    //    var sheet = indexWk.ActiveSheet;
    //    var name = sheet.Name;
    //    var dataCount = Convert.ToInt32(sheet.Range["E3"].value);
    //    if (!name.Contains("【模板】"))
    //    {
    //        MessageBox.Show(@"当前表格不是正确【模板】，不能导出数据");
    //        return;
    //    }
    //    ExcelDic();
    //    var startModeId = sheet.Range["D3"].value.ToString();
    //    var startModeIdFix = sheet.range["F3"].value;
    //    var modeIdRow = new List<List<(int digitCount, string temp)>>();
    //    var tempList = new List<(int digitCount, string temp)> { (1, startModeId) };
    //    modeIdRow.Add(tempList);
    //    var excelIdGroupStart = new List<List<List<(int digitCount, string temp)>>>();
    //    var excelIdGroupStartTemp1 = new List<List<(int digitCount, string temp)>>();
    //    for (var i = 0; i < dataCount; i++)
    //    {
    //        //modeID原始位数
    //        var temp2 = KeyBitCount(startModeId.ToString());
    //        //字段值改写方法
    //        if (startModeIdFix == null) startModeIdFix = "";
    //        var temp1 = CellFixValueKeyList(startModeIdFix.ToString());
    //        //修改字符串
    //        var cellFixValue2 =
    //            RegNumReplaceNew(startModeId.ToString(), temp1, false, temp2, 1 + i);
    //        List<(int digitCount, string temp)> excelIdGroupStartTemp2 = KeyBitCount(cellFixValue2);

    //        excelIdGroupStartTemp1.Add(excelIdGroupStartTemp2);
    //    }
    //    excelIdGroupStart.Add(excelIdGroupStartTemp1);
    //    string writeMode = "修改";
    //    var fileName = new List<string> { sheet.Range["C3"].value.ToString() };
    //    var linksExcel = CreateRelationShip(fileName, modeIdRow, writeMode, excelIdGroupStart);
    //    //把模板连接数据备份到excel
    //    var sheetLink = indexWk.Sheets["索引关键词"];
    //    sheetLink.Range["C2:D200"].ClearContents();
    //    string[,] array = linksExcel.Select(t => new[] { t.Item1, "A"+t.Item2 }).ToArray().ToRectangularArray();
    //    sheetLink.Range["C2:D" + (linksExcel.Count + 1)].Value = array;
    //}
    //public static void ExcelDic()
    //{
    //    var indexWk = App.ActiveWorkbook;
    //    var sheet = indexWk.ActiveSheet;
    //    ExcelLinkDictionary = new Dictionary<string, List<string>>();
    //    ExcelFixKeyDictionary = new Dictionary<string, List<int>>();
    //    ExcelFixKeyMethodDictionary = new Dictionary<string, List<string>>();
    //    //读取模板表数据
    //    var rowsCount = (sheet.Cells[sheet.Rows.Count, "B"].End[XlDirection.xlUp].Row - 4) / 4 + 1;
    //    for (var i = 1; i <= rowsCount; i++)
    //    {
    //        var baseExcel = sheet.Cells[1, 1].Offset[4 + (i - 1) * 4, 1].Value.ToString();
    //        ExcelLinkDictionary[baseExcel] = new List<string>();
    //        ExcelFixKeyDictionary[baseExcel] = new List<int>();
    //        ExcelFixKeyMethodDictionary[baseExcel] = new List<string>();
    //        for (var j = 2; j <= 14; j++)
    //        {
    //            var linkExcel = sheet.Cells[1, 1].Offset[6 + (i - 1) * 4, j + 1].Value;
    //            var baseExcelFixKey = sheet.Cells[1, 1].Offset[7 + (i - 1) * 4, j + 1].Value;
    //            var baseExcelFixKeyMethod = sheet.Cells[1, 1].Offset[5 + (i - 1) * 4, j + 1].Value;
    //            ExcelLinkDictionary[baseExcel].Add(linkExcel);
    //            if (baseExcelFixKey == null)
    //                baseExcelFixKey = 0;
    //            else if (baseExcelFixKey.ToString() == "") baseExcelFixKey = 0;

    //            ExcelFixKeyDictionary[baseExcel].Add(Convert.ToInt32(baseExcelFixKey));
    //            if (baseExcelFixKeyMethod == null) baseExcelFixKeyMethod = "";

    //            ExcelFixKeyMethodDictionary[baseExcel].Add(Convert.ToString(baseExcelFixKeyMethod));
    //        }
    //    }
    //}

    //public static string ValueTypeToStringInNpoi(ICell cell, XSSFWorkbook workbook)
    //{
    //    var cellValueAsString = string.Empty;
    //    if (cell != null)
    //        switch (cell.CellType)
    //        {
    //            case CellType.Numeric:
    //                cellValueAsString = cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
    //                break;
    //            case CellType.String:
    //                cellValueAsString = cell.StringCellValue;
    //                break;
    //            case CellType.Boolean:
    //                cellValueAsString = cell.BooleanCellValue.ToString();
    //                break;
    //            case CellType.Error:
    //                cellValueAsString = cell.ErrorCellValue.ToString();
    //                break;
    //            case CellType.Formula:
    //                // Create a formula evaluator
    //                var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
    //                // Evaluate the formula and get the resulting value
    //                var value = evaluator.Evaluate(cell).NumberValue;
    //                cellValueAsString = value.ToString(CultureInfo.CurrentCulture);
    //                break;
    //            default:
    //                cellValueAsString = "";
    //                break;
    //        }

    //    return cellValueAsString;
    //}

    //public static int FindSourceRow(ISheet sheet, int col, string searchValue, XSSFWorkbook workbook)
    //{
    //    for (var i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
    //    {
    //        var row = sheet.GetRow(i);
    //        if (row != null)
    //        {
    //            var cell = row.GetCell(col);
    //            var cellValue = ValueTypeToStringInNpoi(cell, workbook);
    //            if (cellValue == searchValue) return i;
    //        }
    //    }

    //    return -1;
    //}

    //public static string RegNumReplaceNew(string text, List<(int, int)> digit, bool isCarry,
    //    List<(int digitCount, string temp)> keyBitCount, int addValue)
    //{
    //    var numCount = 1;
    //    var ditCount = 0;
    //    var pattern = "\\d+";
    //    // 使用正则表达式匹配数字
    //    var matches = Regex.Matches(text, pattern);
    //    foreach (Match match in matches)
    //    {
    //        var numStr = match.Value;
    //        var num = long.Parse(numStr);

    //        if (digit.Any(item => item.Item1 == numCount))
    //        {
    //            var newNum = num + (int)Math.Pow(10, digit[ditCount].Item2 - 1) * addValue;
    //            if (isCarry == false)
    //            {
    //                var digitCount =
    //                    (long)Math.Log10(newNum + 1) + 1 -
    //                    keyBitCount[numCount - 1].Item1; //数字位数-----需要更改i和原始数字相关，不能简单的只取一位了；；需要取得原来数字长度
    //                var digitValue = num / (int)Math.Pow(10, digit[ditCount].Item2 - 1) %
    //                                 (int)Math.Pow(10, digitCount + 1); // 获取要增加的数字位的值
    //                if (digitValue + addValue >= (int)Math.Pow(10, digitCount + 1))
    //                {
    //                    var number = num / (int)Math.Pow(10, digit[ditCount].Item2 + digitCount);
    //                    var numberMod = num % (int)Math.Pow(10, digit[ditCount].Item2 - 1);
    //                    var newDigitValue = (digitValue + addValue) * (int)Math.Pow(10, digit[ditCount].Item2 - 1);
    //                    number = number * (int)Math.Pow(10, digit[ditCount].Item2 + digitCount + 1) +
    //                                newDigitValue + numberMod;
    //                    text = text.Replace(numStr, number.ToString());
    //                }
    //                else
    //                {
    //                    text = text.Replace(numStr, newNum.ToString());
    //                }
    //            }
    //            else
    //            {
    //                text = text.Replace(numStr, newNum.ToString());
    //            }

    //            ditCount++;
    //        }
    //        else if (digit.Count == 1 && digit[0].Item1 == 0)
    //        {
    //            var newNum = num + (int)Math.Pow(10, digit[ditCount].Item2 - 1) * addValue;
    //            text = text.Replace(numStr, newNum.ToString());
    //        }

    //        numCount++;
    //    }

    //    return text;
    //}

    //public static List<(string, int)> CreateRelationShip(List<string> oldFileName, List<List<(int digitCount, string temp)>> oldModelId, string writeMode, List<List<List<(int digitCount, string temp)>>> oldExcelIdGroup)
    //{
    //    var indexWk = App.ActiveWorkbook;
    //    var excelPath = indexWk.Path;
    //    var sSheet = indexWk.ActiveSheet;
    //    var dataCount = Convert.ToInt32(sSheet.Range["E3"].value);
    //    var modeIFirstIdList = new List<(string, int)>();
    //    while (true)
    //    {
    //        var newModelId = new List<List<(int digitCount, string temp)>>();
    //        var newFileName = new List<string>();
    //        var newExcelId = new List<List<List<(int digitCount, string temp)>>>();
    //        var count = 0;
    //        foreach (var excelFile in oldFileName)
    //        {
    //            var path = excelPath + @"\" + excelFile;
    //            if (excelFile == "多语言")
    //            {
    //                var newPath = Path.GetDirectoryName(Path.GetDirectoryName(excelPath));
    //                path = newPath + @"\Excels\Localizations\Localizations.xlsx";
    //            }
    //            if (excelFile == "icon.xlsx")
    //            {
    //                continue;
    //            }
    //            var excel = new FileStream(path, FileMode.Open, FileAccess.Read);
    //            var workbook = new XSSFWorkbook(excel);
    //            var sheet = workbook.GetSheetAt(0);
    //            for (var k = 0; k < oldModelId[count].Count; k++)
    //            {
    //                var seachValue = oldModelId[count][k].Item2;
    //                var rowReSourceRow = FindSourceRow(sheet, 1, seachValue, workbook);
    //                if (rowReSourceRow == -1) continue;
    //                //模板ID记录，方便做Link
    //                if (k == 0)
    //                {
    //                    modeIFirstIdList.Add((excelFile, rowReSourceRow+1));
    //                }
    //                var rowSource = sheet.GetRow(rowReSourceRow) ?? sheet.CreateRow(rowReSourceRow);
    //                var colTotal = sheet.GetRow(1).LastCellNum + 1;
    //                if (writeMode == "新增")
    //                    if (sheet.LastRowNum != rowReSourceRow)
    //                        sheet.ShiftRows(rowReSourceRow + 1, sheet.LastRowNum, dataCount, true, false);
    //                //数据复制
    //                for (var i = 0; i < dataCount; i++)
    //                {
    //                    var rowTarget = sheet.GetRow(rowReSourceRow + i + 1) ?? sheet.CreateRow(rowReSourceRow + i + 1);
    //                    for (var j = 1; j < colTotal; j++)
    //                    {
    //                        var cellSource = rowSource.GetCell(j) ?? rowSource.GetCell(j);
    //                        if (cellSource != null)
    //                        {
    //                            var cellSourceValue = ValueTypeToStringInNpoi(cellSource, workbook);
    //                            var cellTarget = rowTarget.GetCell(j) ?? rowTarget.CreateCell(j);
    //                            //if(WriteMode=="修改") continue;
    //                            //表格的ID字段的修改--后续要添加其他字段的更改方式
    //                            if (j == 1)
    //                            {
    //                                var tempValue = oldExcelIdGroup[count][i][k].Item2;
    //                                cellTarget.SetCellValue(tempValue);
    //                                //Debug.Print(cellTarget.ToString());
    //                            }
    //                            else
    //                            {
    //                                cellTarget.SetCellValue(cellSourceValue);
    //                            }

    //                            cellTarget.CellStyle = cellSource.CellStyle;
    //                        }
    //                    }
    //                }

    //                if (excelFile == null) continue;
    //                if (ExcelLinkDictionary.ContainsKey(excelFile))
    //                {
    //                    var indexExcelCount = 0;
    //                    foreach (var indexExcel in ExcelLinkDictionary[excelFile])
    //                    {
    //                        var excelFileFixKey = ExcelFixKeyDictionary[excelFile][indexExcelCount];
    //                        //字典会把空值当0用
    //                        if (excelFileFixKey == 0)
    //                        {
    //                            indexExcelCount++;
    //                            continue;
    //                        }

    //                        //修改字段字典中的字段值，各自方法不一
    //                        var cellFixValueIdList = new List<List<(int digitCount, string temp)>>();
    //                        var newMode = new List<(int digitCount, string temp)>();
    //                        for (var i = 0; i < dataCount; i++)
    //                        {
    //                            var rowFix = sheet.GetRow(rowReSourceRow + i + 1) ?? sheet.CreateRow(rowReSourceRow + i + 1);
    //                            var cellFix = rowFix.GetCell(excelFileFixKey) ?? rowFix.CreateCell(excelFileFixKey);
    //                            var cellFixValue = ValueTypeToStringInNpoi(cellFix, workbook);
    //                            if(cellFixValue == null) continue;
    //                            //特殊表格例外处理
    //                            if (excelFile == "PictorialBookTagData.xlsx" && excelFileFixKey == 4)
    //                            {
    //                                var tempSc1 = sheet.GetRow(rowReSourceRow + i + 1).GetCell(excelFileFixKey-1).ToString();
    //                                if (tempSc1 == "1")
    //                                {
    //                                    continue;
    //                                }
    //                            }
    //                            else if (excelFile == "ShopMarketGood.xlsx" && excelFileFixKey == 4)
    //                            {
    //                                var tempSc1 = sheet.GetRow(rowReSourceRow + i + 1).GetCell(excelFileFixKey).ToString();
    //                                var tempSc2 = tempSc1.Contains("110");
    //                                if (tempSc2)
    //                                {
    //                                    continue;
    //                                }
    //                            }
    //                            //每个字段的Value修改方式不一，需要调用方法:检测string是否有[，如果有则需要正则把所有的数值提取出来并替换
    //                            //字段每个数字位数统计，原始modeID统计
    //                            var cellSourceValueList = KeyBitCount(cellFixValue);
    //                            //字段值改写方法
    //                            var temp1 = CellFixValueKeyList(ExcelFixKeyMethodDictionary[excelFile][indexExcelCount]);
    //                            //修改字符串
    //                            var cellFixValue2 = RegNumReplaceNew(cellFixValue, temp1, true, cellSourceValueList, 1 + i);
    //                            //统计新ID
    //                            var temp2 = KeyBitCount(cellFixValue2);
    //                            //标记重复项
    //                            newMode = cellSourceValueList.Except(temp2).ToList();
    //                            var newFix = temp2.Except(cellSourceValueList).ToList();

    //                            cellFixValueIdList.Add(newFix);
    //                            cellFix.SetCellValue(cellFixValue2);
    //                            //cellFix.CellStyle = cellFix.CellStyle;
    //                        }
    //                        //有关联表的字段的ID传递出去
    //                        //表格关联字典中寻找下一个递归文件，有关联表的字段ID要生成List递归
    //                        if (!string.IsNullOrEmpty(indexExcel))
    //                        {
    //                            newFileName.Add(indexExcel);
    //                            newModelId.Add(newMode);
    //                            newExcelId.Add(cellFixValueIdList);
    //                        }

    //                        indexExcelCount++;
    //                    }
    //                }
    //            }
    //            excel.Close();
    //            var excel2 = new FileStream(path, FileMode.Create, FileAccess.Write);
    //            workbook.Write(excel2);
    //            workbook.Close();
    //            excel2.Close();
    //            count++;
    //            App.StatusBar = "正在处理:" + excelFile + "文件";
    //        }

    //        if (newFileName.Count > 0)
    //        {
    //            oldFileName = newFileName;
    //            oldModelId = newModelId;
    //            oldExcelIdGroup = newExcelId;
    //            continue;
    //        }

    //        break;
    //    }
    //    return modeIFirstIdList;
    //}

    //public static void FixValueType()
    //{
    //    //string str = "1#2,3#2,2"; // 要处理的字符串
    //    var str = "1#2,3#2,4"; // 要处理的字符串
    //    var tempList = CellFixValueKeyList(str);

    //    var str1 = "[11001,11002,10003,10004]";

    //    var keyBitCount = KeyBitCount(str1);

    //    for (int i = 0; i < 20; i++)
    //    {
    //        str1 = RegNumReplaceNew(str1, tempList, false, keyBitCount, 2);
    //        Debug.Print(str1);

    //    }
    //}

    //private static List<(int digitCount, string temp)> KeyBitCount(string str)
    //{
    //    var regex = new Regex(@"\d+");
    //    var matches = regex.Matches(str);
    //    var keyBitCount = new List<(int digitCount, string temp)>();
    //    foreach (var match in matches)
    //    {
    //        var temp = match.ToString();
    //        var digitCount =temp.Length;
    //        keyBitCount.Add((digitCount, temp));
    //    }

    //    return keyBitCount;
    //}

    //private static List<(int, int)> CellFixValueKeyList(string str)
    //{
    //    var monkeyList = new List<(int, int)>();

    //    str ??= "";

    //    if (str.Contains(','))
    //    {
    //        var pairs = str.Split(',');
    //        foreach (var pair in pairs)
    //        {
    //            if (pair.Contains('#'))
    //            {
    //                var parts = pair.Split('#');
    //                if (!int.TryParse(parts[0], out var key)) // 尝试将值解析为整数，如果解析失败就将值设为 0
    //                {
    //                    MessageBox.Show($@"{str}#前必须有数值");
    //                    Environment.Exit(0);
    //                }

    //                if (!int.TryParse(parts[1], out var value)) // 尝试将值解析为整数，如果解析失败就将值设为 0
    //                    value = 1;

    //                monkeyList.Add((key, value));
    //            }
    //            else
    //            {
    //                monkeyList.Add((int.Parse(pair), 1));
    //            }
    //        }
    //    }
    //    else
    //    {
    //        if (str.Contains('#'))
    //        {
    //            var parts = str.Split('#');
    //            var key = Convert.ToInt32(parts[0]);
    //            var value = Convert.ToInt32(parts[1]);
    //            monkeyList.Add((key, value));
    //        }
    //        else
    //        {
    //            int strTemp;
    //            if (str == "")
    //            {
    //                strTemp = 0;
    //                monkeyList.Add((strTemp, 1));
    //            }
    //            else
    //            {
    //                strTemp = int.Parse(str);
    //                monkeyList.Add((0, strTemp));
    //            }
    //        }
    //    }

    //    return monkeyList;
    //}

    //public static void ExcelHyperLinks()
    //{
    //    var indexWk = App.ActiveWorkbook;
    //    var excelPath = indexWk.Path;
    //    var sheet = indexWk.ActiveSheet;
    //    //获取linkList
    //    var sheet2 = indexWk.Sheets["索引关键词"];
    //    var linksExcel = new List<(string, string)>();
    //    for (int i =2;i<=100;i++)
    //    {
    //        var temp = sheet2.Cells[i, 3].value;
    //        var temp2 = sheet2.Cells[i, 4].value;
    //        linksExcel.Add((temp,temp2));
    //    }
    //    for (var i = 3; i <= 101; i++)
    //    {
    //        for (var j = 2; j <= 20; j++)
    //        {
    //            var cell = sheet.Cells[i, j];
    //            if (cell.value != null && cell.value.ToString().Contains(".xlsx"))
    //            {
    //                int m = 0;
    //                string rows = "-1";
    //                foreach (var unused in linksExcel)
    //                {
    //                    if (cell.value.ToString() == linksExcel[m].Item1)
    //                    {
    //                        rows = linksExcel[m].Item2;
    //                    }

    //                    m++;
    //                }
    //                var path = excelPath + @"\" + cell.value.ToString();
    //                if (rows != "-1")
    //                {
    //                    var excel = new FileStream(excelPath + @"\" + cell.value.ToString(), FileMode.Open,
    //                    FileAccess.Read);
    //                    var workbook = new XSSFWorkbook(excel);
    //                    var sheetName = workbook.GetSheetAt(0).SheetName;
    //                    path = excelPath + @"\" + cell.value.ToString() + "#" + sheetName + "!" + rows;
    //                    workbook.Close();
    //                    excel.Close();
    //                }
    //                cell.Hyperlinks.Add(cell, path);
    //                cell.Font.Size = 9;
    //                cell.Font.Name = "微软雅黑";
    //                //cell.Copy();
    //                //cell.PasteSpecial(XlPasteType.xlPasteFormats);
    //                //cell.Style = cs1;
    //            }
    //        }
    //    }
    //}
}
/// <summary>
/// 废弃的Excel关系图2
/// </summary>
internal class ExcelRelationShipEpPlus
{
    //private static readonly dynamic App = ExcelDnaUtil.Application;
    //public static Dictionary<string, List<string>> ExcelLinkDictionary;
    //public static Dictionary<string, List<int>> ExcelFixKeyDictionary;
    //public static Dictionary<string, List<string>> ExcelFixKeyMethodDictionary;

    //public static void StartExcelData()
    //{
    //    var indexWk = App.ActiveWorkbook;
    //    var sheet = indexWk.ActiveSheet;
    //    var name = sheet.Name;
    //    var dataCount = Convert.ToInt32(sheet.Range["E3"].value);
    //    if (!name.Contains("【模板】"))
    //    {
    //        MessageBox.Show(@"当前表格不是正确【模板】，不能导出数据");
    //        return;
    //    }
    //    ExcelDic();
    //    var startModeId = sheet.Range["D3"].value.ToString();
    //    var startModeIdFix = sheet.range["F3"].value;
    //    var modeIdRow = new List<List<(int digitCount, string temp)>>();
    //    var tempList = new List<(int digitCount, string temp)> { (1, startModeId) };
    //    modeIdRow.Add(tempList);
    //    var excelIdGroupStart = new List<List<List<(int digitCount, string temp)>>>();
    //    var excelIdGroupStartTemp1 = new List<List<(int digitCount, string temp)>>();
    //    for (var i = 0; i < dataCount; i++)
    //    {
    //        //modeID原始位数
    //        var temp2 = KeyBitCount(startModeId.ToString());
    //        //字段值改写方法
    //        if (startModeIdFix == null) startModeIdFix = "";
    //        var temp1 = CellFixValueKeyList(startModeIdFix.ToString());
    //        //修改字符串
    //        var cellFixValue2 =
    //            RegNumReplaceNew(startModeId.ToString(), temp1, false, temp2, 1 + i);
    //        List<(int digitCount, string temp)> excelIdGroupStartTemp2 = KeyBitCount(cellFixValue2);

    //        excelIdGroupStartTemp1.Add(excelIdGroupStartTemp2);
    //    }
    //    excelIdGroupStart.Add(excelIdGroupStartTemp1);
    //    string writeMode = "新增";
    //    var fileName = new List<string> { sheet.Range["C3"].value.ToString() };

    //    var sw = new Stopwatch();
    //    sw.Start();
    //    var linksExcel = CreateRelationShip(fileName, modeIdRow, writeMode, excelIdGroupStart);

    //    sw.Stop();
    //    var ts2 = sw.Elapsed;
    //    Debug.Print("写入用时："+ts2.ToString());

    //    //把模板连接数据备份到excel
    //    var sheetLink = indexWk.Sheets["索引关键词"];
    //    sheetLink.Range["C2:D200"].ClearContents();
    //    string[,] array = linksExcel.Select(t => new[] { t.Item1, "A"+t.Item2 }).ToArray().ToRectangularArray();
    //    sheetLink.Range["C2:D" + (linksExcel.Count + 1)].Value = array;
    //    ExcelHyperLinks();
    //    //一般新增只进行1次
    //    sheet.range["B3"].value = "修复";
    //}

    //public static void ExcelDic()
    //{
    //    var indexWk = App.ActiveWorkbook;
    //    var sheet = indexWk.ActiveSheet;
    //    ExcelLinkDictionary = new Dictionary<string, List<string>>();
    //    ExcelFixKeyDictionary = new Dictionary<string, List<int>>();
    //    ExcelFixKeyMethodDictionary = new Dictionary<string, List<string>>();
    //    //读取模板表数据
    //    var rowsCount = (sheet.Cells[sheet.Rows.Count, "B"].End[XlDirection.xlUp].Row - 4) / 4 + 1;
    //    for (var i = 1; i <= rowsCount; i++)
    //    {
    //        var baseExcel = sheet.Cells[1, 1].Offset[4 + (i - 1) * 4, 1].Value.ToString();
    //        ExcelLinkDictionary[baseExcel] = new List<string>();
    //        ExcelFixKeyDictionary[baseExcel] = new List<int>();
    //        ExcelFixKeyMethodDictionary[baseExcel] = new List<string>();
    //        for (var j = 2; j <= 14; j++)
    //        {
    //            var linkExcel = sheet.Cells[1, 1].Offset[6 + (i - 1) * 4, j + 1].Value;
    //            var baseExcelFixKey = sheet.Cells[1, 1].Offset[7 + (i - 1) * 4, j + 1].Value;
    //            var baseExcelFixKeyMethod = sheet.Cells[1, 1].Offset[5 + (i - 1) * 4, j + 1].Value;
    //            ExcelLinkDictionary[baseExcel].Add(linkExcel);
    //            if (baseExcelFixKey == null)
    //                baseExcelFixKey = 0;
    //            else if (baseExcelFixKey.ToString() == "") baseExcelFixKey = 0;

    //            ExcelFixKeyDictionary[baseExcel].Add(Convert.ToInt32(baseExcelFixKey));
    //            if (baseExcelFixKeyMethod == null) baseExcelFixKeyMethod = "";

    //            ExcelFixKeyMethodDictionary[baseExcel].Add(Convert.ToString(baseExcelFixKeyMethod));
    //        }
    //    }
    //}
    //public static int FindSourceRow(ExcelWorksheet sheet, int col, string searchValue)
    //{
    //    for (int row = 2; row <= sheet.Dimension.End.Row; row++)
    //    {
    //        // 获取当前行的单元格数据
    //        var cellValue = sheet.Cells[row, col].Value;

    //        // 如果找到了匹配的值
    //        if (cellValue != null && cellValue.ToString() == searchValue)
    //        {
    //            // 返回该单元格的行地址
    //            var cellAddress = new ExcelCellAddress(row, col);
    //            var rowAddress = cellAddress.Row;
    //            return rowAddress;
    //        }
    //    }
    //    return -1;
    //}

    //public static string RegNumReplaceNew(string text, List<(int, int)> digit, bool isCarry,
    //    List<(int digitCount, string temp)> keyBitCount, int addValue)
    //{
    //    var numCount = 1;
    //    var ditCount = 0;
    //    var pattern = "\\d+";
    //    // 使用正则表达式匹配数字
    //    var matches = Regex.Matches(text, pattern);
    //    foreach (Match match in matches)
    //    {
    //        var numStr = match.Value;
    //        var num = long.Parse(numStr);

    //        if (digit.Any(item => item.Item1 == numCount))
    //        {
    //            var newNum = num + (int)Math.Pow(10, digit[ditCount].Item2 - 1) * addValue;
    //            if (isCarry == false)
    //            {
    //                var digitCount =
    //                    (long)Math.Log10(newNum + 1) + 1 -
    //                    keyBitCount[numCount - 1].Item1; //数字位数-----需要更改i和原始数字相关，不能简单的只取一位了；；需要取得原来数字长度
    //                var digitValue = num / (int)Math.Pow(10, digit[ditCount].Item2 - 1) %
    //                                 (int)Math.Pow(10, digitCount + 1); // 获取要增加的数字位的值
    //                if (digitValue + addValue >= (int)Math.Pow(10, digitCount + 1))
    //                {
    //                    var number = num / (int)Math.Pow(10, digit[ditCount].Item2 + digitCount);
    //                    var numberMod = num % (int)Math.Pow(10, digit[ditCount].Item2 - 1);
    //                    var newDigitValue = (digitValue + addValue) * (int)Math.Pow(10, digit[ditCount].Item2 - 1);
    //                    number = number * (int)Math.Pow(10, digit[ditCount].Item2 + digitCount + 1) +
    //                                newDigitValue + numberMod;
    //                    text = text.Replace(numStr, number.ToString());
    //                }
    //                else
    //                {
    //                    text = text.Replace(numStr, newNum.ToString());
    //                }
    //            }
    //            else
    //            {
    //                text = text.Replace(numStr, newNum.ToString());
    //            }

    //            ditCount++;
    //        }
    //        else if (digit.Count == 1 && digit[0].Item1 == 0)
    //        {
    //            var newNum = num + (int)Math.Pow(10, digit[ditCount].Item2 - 1) * addValue;
    //            text = text.Replace(numStr, newNum.ToString());
    //        }

    //        numCount++;
    //    }

    //    return text;
    //}

    //public static List<(string, int)> CreateRelationShip(List<string> oldFileName, List<List<(int digitCount, string temp)>> oldModelId, string writeMode, List<List<List<(int digitCount, string temp)>>> oldExcelIdGroup)
    //{
    //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    //    var indexWk = App.ActiveWorkbook;
    //    var excelPath = indexWk.Path;
    //    var sheet = indexWk.ActiveSheet;
    //    var dataCount = Convert.ToInt32(sheet.Range["E3"].value);
    //    var modeIFirstIdList = new List<(string, int)>();
    //    while (true)
    //    {
    //        var newModelId = new List<List<(int digitCount, string temp)>>();
    //        var newFileName = new List<string>();
    //        var newExcelId = new List<List<List<(int digitCount, string temp)>>>();
    //        var count = 0;
    //        foreach (var excelFile in oldFileName)
    //        {
    //            var path = excelPath + @"\" + excelFile;
    //            if (excelFile == "多语言")
    //            {
    //                var newPath = Path.GetDirectoryName(Path.GetDirectoryName(excelPath));
    //                path = newPath+ @"\Excels\Localizations\Localizations.xlsx";
    //            }
    //            //个别表需要单独批量处理
    //            if (excelFile is "icon.xlsx" or "item.xlsx" or "PictorialBookItemData.xlsx" or "ObjectLevel.xlsx" or null)
    //            {
    //                continue;
    //            }
    //            var excel = new ExcelPackage(new FileInfo(path));
    //            var workBook = excel.Workbook;
    //            sheet = workBook.Worksheets[0];

    //            for (var k = 0; k < oldModelId[count].Count; k++)
    //            {
    //                var seachValue = oldModelId[count][k].Item2;
    //                var rowReSourceRow = FindSourceRow(sheet, 2, seachValue);
    //                if (rowReSourceRow == -1) continue;
    //                //模板ID记录，方便做Link
    //                if (k == 0)
    //                {
    //                    modeIFirstIdList.Add((excelFile, rowReSourceRow+1));
    //                }
    //                var colCount = sheet.Dimension.Columns;
    //                if (writeMode == "新增")
    //                {
    //                    sheet.InsertRow(rowReSourceRow + 1, dataCount);
    //                }
    //                //数据复制
    //                for (var i =0; i < dataCount; i++)
    //                {
    //                    for (int j = 2; j < colCount+1; j++)
    //                    {
    //                        var cellSource = sheet.Cells[rowReSourceRow, j];
    //                        var cellTarget = sheet.Cells[rowReSourceRow+i+1,j];
    //                        if (j == 2)
    //                        {
    //                            //索引编号列数据单独更改
    //                            var tempValue = oldExcelIdGroup[count][i][k].Item2;
    //                            cellTarget.Value = tempValue;
    //                            //单元格样式更改
    //                            cellSource.CopyStyles(cellTarget);
    //                        }
    //                        else
    //                        {
    //                            cellSource.Copy(cellTarget);
    //                        }
    //                    }
    //                    //Debug.Print(cellTarget.ToString());
    //                }
    //                if (excelFile == null) continue;
    //                if (ExcelLinkDictionary.ContainsKey(excelFile))
    //                {
    //                    var indexExcelCount = 0;
    //                    var mutilExcelInOneKey = new List<string>();
    //                    foreach (var indexExcel in ExcelLinkDictionary[excelFile])
    //                    {
    //                        var excelFileFixKey = ExcelFixKeyDictionary[excelFile][indexExcelCount];
    //                        //字典会把空值当0用
    //                        if (excelFileFixKey == 0)
    //                        {
    //                            indexExcelCount++;
    //                            continue;
    //                        }

    //                        //修改字段字典中的字段值，各自方法不一
    //                        var cellFixValueIdList = new List<List<(int digitCount, string temp)>>();
    //                        var newMode = new List<(int digitCount, string temp)>();
    //                        for (var i = 0; i < dataCount; i++)
    //                        {
    //                            var cellFix = sheet.Cells[rowReSourceRow + i + 1, excelFileFixKey + 1];
    //                            var cellFixValue="";
    //                            if (cellFix.Value != null)
    //                            {
    //                                //Debug.Print(excelFile + "::" + cellFix.Value);
    //                                cellFixValue = cellFix.Value.ToString();
    //                            }
    //                            //特殊表格例外处理
    //                            if (excelFile == "PictorialBookTagData.xlsx" && excelFileFixKey ==4)
    //                            {
    //                                var tempSc1 = sheet.Cells[rowReSourceRow + i + 1, 2].Value.ToString();
    //                                var tempSc2 = tempSc1[tempSc1.Length - 1].ToString();
    //                                if (tempSc2=="2")
    //                                {
    //                                    continue;
    //                                }
    //                            }
    //                            else if (excelFile == "ShopMarketGood.xlsx" && excelFileFixKey == 4)
    //                            {
    //                                var tempSc1 = sheet.Cells[rowReSourceRow + i + 1, excelFileFixKey+1].Value.ToString();
    //                                var tempSc2 = tempSc1.Contains("110");
    //                                if (tempSc2)
    //                                {
    //                                    continue;
    //                                }
    //                            }
    //                            else if (excelFile == "Mission.xlsx" && excelFileFixKey == 13)
    //                            {
    //                                var tempSc1 = sheet.Cells[rowReSourceRow + i + 1, excelFileFixKey].Value.ToString();
    //                                if (tempSc1 == "30")
    //                                {
    //                                    mutilExcelInOneKey.Add("abc");
    //                                }
    //                                else if (tempSc1 == "PictorialBookItemData")
    //                                {
    //                                    mutilExcelInOneKey.Add("object.xlsx");
    //                                }
    //                            }
    //                            //每个字段的Value修改方式不一，需要调用方法:检测string是否有[，如果有则需要正则把所有的数值提取出来并替换
    //                            //字段每个数字位数统计，原始modeID统计
    //                            var cellSourceValueList = KeyBitCount(cellFixValue);
    //                            //字段值改写方法
    //                            var temp1 = CellFixValueKeyList(ExcelFixKeyMethodDictionary[excelFile][indexExcelCount]);
    //                            //修改字符串
    //                            var cellFixValue2 = RegNumReplaceNew(cellFixValue, temp1, true, cellSourceValueList, 1 + i);
    //                            //统计新ID
    //                            var temp2 = KeyBitCount(cellFixValue2);
    //                            //标记重复项
    //                            newMode = cellSourceValueList.Except(temp2).ToList();
    //                            var newFix = temp2.Except(cellSourceValueList).ToList();

    //                            cellFixValueIdList.Add(newFix);
    //                            cellFix.Value =cellFixValue2;
    //                        }
    //                        //有关联表的字段的ID传递出去
    //                        //表格关联字典中寻找下一个递归文件，有关联表的字段ID要生成List递归
    //                        if (!string.IsNullOrEmpty(indexExcel))
    //                        {
    //                            if (mutilExcelInOneKey.Count != 0)
    //                            {
    //                                newFileName = mutilExcelInOneKey;
    //                            }
    //                            newFileName.Add(indexExcel);
    //                            newModelId.Add(newMode);
    //                            newExcelId.Add(cellFixValueIdList);
    //                        }

    //                        indexExcelCount++;
    //                    }
    //                }
    //            }
    //            excel.Save();
    //            excel.Dispose();
    //            count++;
    //            App.StatusBar = "正在处理:" + excelFile + "文件";
    //        }
    //        if (newFileName.Count > 0)
    //        {
    //            oldFileName = newFileName;
    //            oldModelId = newModelId;
    //            oldExcelIdGroup = newExcelId;
    //            continue;
    //        }

    //        break;
    //    }
    //    return modeIFirstIdList;
    //}

    //public static List<(int digitCount, string temp)> KeyBitCount(string str)
    //{
    //    var regex = new Regex(@"\d+");
    //    var matches = regex.Matches(str);
    //    var keyBitCount = new List<(int digitCount, string temp)>();
    //    foreach (var match in matches)
    //    {
    //        var temp = match.ToString();
    //        var digitCount =temp.Length;
    //        keyBitCount.Add((digitCount, temp));
    //    }

    //    return keyBitCount;
    //}

    //public static List<(int, int)> CellFixValueKeyList(string str)
    //{
    //    var monkeyList = new List<(int, int)>();

    //    str ??= "";

    //    if (str.Contains(','))
    //    {
    //        var pairs = str.Split(',');
    //        foreach (var pair in pairs)
    //        {
    //            if (pair.Contains('#'))
    //            {
    //                var parts = pair.Split('#');
    //                if (!int.TryParse(parts[0], out var key)) // 尝试将值解析为整数，如果解析失败就将值设为 0
    //                {
    //                    MessageBox.Show($@"{str}#前必须有数值");
    //                    Environment.Exit(0);
    //                }

    //                if (!int.TryParse(parts[1], out var value)) // 尝试将值解析为整数，如果解析失败就将值设为 0
    //                    value = 1;

    //                monkeyList.Add((key, value));
    //            }
    //            else
    //            {
    //                monkeyList.Add((int.Parse(pair), 1));
    //            }
    //        }
    //    }
    //    else
    //    {
    //        if (str.Contains('#'))
    //        {
    //            var parts = str.Split('#');
    //            var key = Convert.ToInt32(parts[0]);
    //            var value = Convert.ToInt32(parts[1]);
    //            monkeyList.Add((key, value));
    //        }
    //        else
    //        {
    //            int strTemp;
    //            if (str == "")
    //            {
    //                strTemp = 0;
    //                monkeyList.Add((strTemp, 1));
    //            }
    //            else
    //            {
    //                strTemp = int.Parse(str);
    //                monkeyList.Add((0, strTemp));
    //            }
    //        }
    //    }

    //    return monkeyList;
    //}


    //public static void FixValueType()
    //{
    //    string str = "abc1001cde1001efg1001";
    //    string pattern = @"(\D*\d+)\D+(\d+)(\D*\d+)*";
    //    string result = Regex.Replace(str, pattern, m => {
    //        int num = int.Parse(m.Groups[2].Value);
    //        num++;
    //        return m.Groups[1].Value + num.ToString() + m.Groups[3].Value;
    //    });
    //    Debug.Print(result);
    //}
    //public static void ExcelHyperLinks()
    //{
    //    var indexWk = App.ActiveWorkbook;
    //    var excelPath = indexWk.Path;
    //    var sheet = indexWk.ActiveSheet;
    //    //获取linkList
    //    var sheet2 = indexWk.Sheets["索引关键词"];
    //    var linksExcel = new List<(string, string)>();
    //    for (int i =2;i<=100;i++)
    //    {
    //        var temp = sheet2.Cells[i, 3].value;
    //        var temp2 = sheet2.Cells[i, 4].value;
    //        linksExcel.Add((temp,temp2));
    //    }
    //    for (var i = 3; i <= 201; i++)
    //    {
    //        for (var j = 2; j <= 20; j++)
    //        {
    //            var cell = sheet.Cells[i, j];
    //            if (cell.value != null && cell.value.ToString().Contains(".xlsx"))
    //            {
    //                int m = 0;
    //                string rows = "-1";
    //                foreach (var unused in linksExcel)
    //                {
    //                    if (cell.value.ToString() == linksExcel[m].Item1)
    //                    {
    //                        rows = linksExcel[m].Item2;
    //                    }

    //                    m++;
    //                }
    //                var path = excelPath + @"\" + cell.value.ToString();
    //                if (rows != "-1")
    //                {
    //                    var excel = new FileStream(excelPath + @"\" + cell.value.ToString(), FileMode.Open,
    //                    FileAccess.Read);
    //                    var workbook = new XSSFWorkbook(excel);
    //                    var sheetName = workbook.GetSheetAt(0).SheetName;
    //                    path = excelPath + @"\" + cell.value.ToString() + "#" + sheetName + "!" + rows;
    //                    workbook.Close();
    //                    excel.Close();
    //                }
    //                cell.Hyperlinks.Add(cell, path);
    //                cell.Font.Size = 9;
    //                cell.Font.Name = "微软雅黑";
    //                //cell.Copy();
    //                //cell.PasteSpecial(XlPasteType.xlPasteFormats);
    //                //cell.Style = cs1;
    //            }
    //        }
    //    }
    //}
}