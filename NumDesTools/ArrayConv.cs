using System.Collections.Concurrent;
using System.Data;
using System.Data.OleDb;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MiniExcelLibs;
using NumDesTools.Config;
using OfficeOpenXml;
using DataTable = System.Data.DataTable;
using ExcelReference = ExcelDna.Integration.ExcelReference;

// ReSharper disable All

#pragma warning disable CA1416

namespace NumDesTools;

public static partial class PubMetToExcel
{
    public static Color GetCellBackgroundColor(Range cell)
    {
        var color = Color.Empty;

        if (cell.Interior.Color != null)
        {
            object excelColor = cell.Interior.Color;
            if (excelColor is double)
            {
                var colorValue = (double)excelColor;
                var intValue = (int)colorValue;
                var red = intValue & 0xFF;
                var green = (intValue & 0xFF00) >> 8;
                var blue = (intValue & 0xFF0000) >> 16;
                color = Color.FromArgb(red, green, blue);
            }
        }

        return color;
    }

    public static List<string> ReadWriteTxt(string filePath)
    {
        var textLineList = new List<string>();
        if (!File.Exists(filePath))
        {
            if (filePath != null)
                using (var writer = File.CreateText(filePath))
                {
                    writer.WriteLine("Alice路径");
                    writer.WriteLine("Cove路径");
                    writer.Close();
                }
        }
        else
        {
            using var reader = new StreamReader(filePath);
            while (reader.ReadLine() is { } line)
                textLineList.Add(line);
        }

        return textLineList;
    }

    public static string ErrorLogAnalysis(
        List<List<(string, string, string)>> errorList,
        Worksheet sheet
    )
    {
        var errorLog = "";
        for (var i = 0; i < errorList.Count; i++)
        for (var j = 0; j < errorList[i].Count; j++)
        {
            var errorCell = errorList[i][j].Item1;
            var errorExcelLog = errorList[i][j].Item2;
            var errorExcelName = errorList[i][j].Item3;
            if (errorCell == "-1")
                continue;
            errorLog =
                errorLog + "【" + errorCell + "】" + errorExcelName + "#" + errorExcelLog + "\r\n";
        }

        return errorLog;
    }

    public static string ConvertToExcelColumn(int columnNumber)
    {
        var columnName = "";

        while (columnNumber > 0)
        {
            var remainder = (columnNumber - 1) % 26;
            columnName = (char)('A' + remainder) + columnName;
            columnNumber = (columnNumber - 1) / 26;
        }

        return columnName;
    }

    public static void OpenExcelAndSelectCell(string filePath, string sheetName, string cellAddress)
    {
        try
        {
            if (!File.Exists(filePath))
            {
                // ReSharper disable LocalizableElement
                MessageBox.Show(@"文件不存在，请检查！");
                // ReSharper restore LocalizableElement
                return;
            }

            AppServices.App.ScreenUpdating = false;
            var workbook = AppServices.App.Workbooks.Open(filePath);

            Worksheet worksheet = null;
            try
            {
                // 尝试获取工作表
                worksheet = (Worksheet)workbook.Sheets[sheetName];
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // 如果工作表不存在，则选择第一个工作表
                worksheet = (Worksheet)workbook.Sheets[1];
            }

            var cellAddressDefault = "1";
            if (cellAddress != null)
            {
                MatchCollection matches = Regex.Matches(cellAddress, @"\d+");
                cellAddressDefault = matches[0].ToString();
                var realCellAddress = $"B{cellAddressDefault}:Z{cellAddressDefault}";
                var cellRange = worksheet.Range[realCellAddress];

                AppServices.App.ScreenUpdating = true;
                worksheet.Activate();
                cellRange.Select();
            }

            AppServices.App.ScreenUpdating = true;
        }
        catch (Exception e)
        {
            MessageBox.Show($"打开文件失败: {e.Message}");
        }

        GC.Collect();
    }

    public static void ListToArrayToRange(
        List<List<object>> targetList,
        dynamic workSheet,
        int startRow,
        int startCol
    )
    {
        var rowCount = targetList.Count;
        var columnCount = 0;
        foreach (var innerList in targetList)
        {
            var currentColumnCount = innerList.Count;
            columnCount = Math.Max(columnCount, currentColumnCount);
        }

        var targetDataArr = new object[rowCount, columnCount];
        for (var i = 0; i < rowCount; i++)
        for (var j = 0; j < targetList[i].Count; j++)
            targetDataArr[i, j] = targetList[i][j];
        var targetRange = workSheet.Range[
            workSheet.Cells[startRow, startCol],
            workSheet.Cells[startRow + rowCount - 1, startCol + columnCount - 1]
        ];
        targetRange.Value = targetDataArr;
    }

    //Alice文件路径修正
    public static (string filePath, string sheetName) AliceFilePathFix(
        string workbookPath,
        string selectSheetName
    )
    {
        workbookPath = Path.GetDirectoryName(workbookPath);

        var isMatch = selectSheetName.Contains(".xls");

        string filePath = String.Empty;
        string sheetName = "Sheet1";
        if (isMatch)
        {
            if (!workbookPath.Contains(@"\Tables"))
            {
                workbookPath = workbookPath + @"\Tables\";
            }
            else
            {
                workbookPath = workbookPath + @"\";
            }

            if (selectSheetName.Contains("#") && !selectSheetName.Contains("##"))
            {
                var excelSplit = selectSheetName.Split("#");
                filePath = workbookPath + excelSplit[0];
                sheetName = excelSplit[1];
            }
            else if (selectSheetName.Contains("##"))
            {
                var excelSplit = selectSheetName.Split("##");
                var sharpCount = excelSplit.Length;
                if (selectSheetName.Contains("克朗代克"))
                {
                    filePath = workbookPath + excelSplit[0] + @"\" + excelSplit[1];
                    sheetName = sharpCount == 3 ? excelSplit[2] : "Sheet1";
                }
                else
                {
                    selectSheetName = workbookPath + excelSplit[0];
                    sheetName = excelSplit[1];
                }
            }
            else
            {
                while (workbookPath.Contains(@"\Tables"))
                {
                    workbookPath = Path.GetDirectoryName(workbookPath);
                }

                switch (selectSheetName)
                {
                    case "Localizations.xlsx":
                        filePath = workbookPath + @"\Localizations\Localizations.xlsx";
                        break;
                    case "UIConfigs.xlsx":
                        filePath = workbookPath + @"\UIs\UIConfigs.xlsx";
                        break;
                    case "UIItemConfigs.xlsx":
                        filePath = workbookPath + @"\UIs\UIItemConfigs.xlsx";
                        break;
                    default:
                        filePath = workbookPath + @"\Tables\" + selectSheetName;
                        break;
                }

                sheetName = "Sheet1";
            }
        }

        return (filePath, sheetName);
    }

    //二维数组搜索指定行的数据，返回指定行对应列数据
    public static string FindValueInFirstRow(
        object[,] array,
        string value,
        int findIndex = 0,
        int returnIndex = 1
    )
    {
        // 获取数组的列数
        int columns = array.GetLength(1);
        for (int col = 0; col < columns; col++)
        {
            if (array[findIndex, col]?.ToString() == value)
            {
                return array[returnIndex, col]?.ToString();
            }
        }

        // 如果未找到匹配的值，返回 null
        return string.Empty;
    }

    //Range二维数组List化
    public static List<List<object>> RangeDataToList(object[,] rangeValue)
    {
        var rows = rangeValue.GetLength(0);
        var columns = rangeValue.GetLength(1);
        var sheetData = new List<List<object>>();
        for (var row = 1; row <= rows; row++)
        {
            var rowList = new List<object>();
            for (var column = 1; column <= columns; column++)
            {
                var value = rangeValue[row, column];
                rowList.Add(value);
            }

            sheetData.Add(rowList);
        }

        return sheetData;
    }

    //二维数组List化
    public static List<List<object>> Array2DDataToList(object[,] rangeValue)
    {
        var rows = rangeValue.GetLength(0);
        var columns = rangeValue.GetLength(1);
        var sheetData = new List<List<object>>();
        for (var row = 0; row < rows; row++)
        {
            var rowList = new List<object>();
            for (var column = 0; column < columns; column++)
            {
                var value = rangeValue[row, column];
                rowList.Add(value);
            }

            sheetData.Add(rowList);
        }

        return sheetData;
    }

    //二维List一维化
    public static List<object> List2DToListRowOrCol(
        List<List<object>> twoDimensionalList,
        bool byRow
    )
    {
        List<object> flattenedList = new List<object>();

        if (byRow)
        {
            foreach (var row in twoDimensionalList)
            {
                flattenedList.AddRange(row);
            }
        }
        else
        {
            if (twoDimensionalList.Count == 0)
                return flattenedList;
            int columnCount = twoDimensionalList[0].Count;
            for (int col = 0; col < columnCount; col++)
            {
                foreach (var row in twoDimensionalList)
                {
                    flattenedList.Add(row[col]);
                }
            }
        }

        return flattenedList;
    }

    public static List<int> GenerateUniqueRandomList(int minValue, int maxValue, int baseValue)
    {
        var list = new List<int>();

        for (var i = minValue; i <= maxValue; i++)
            list.Add(i + baseValue);

        var random = new Random();
        var n = list.Count;
        for (var i = n - 1; i > 0; i--)
        {
            var j = random.Next(0, i + 1);
            var temp = list[i];
            list[i] = list[j];
            list[j] = temp;
        }

        return list;
    }

    //二维List转二维数组
    public static object[,] ConvertListToArray(List<List<object>> listOfLists)
    {
        var rowCount = listOfLists.Count;
        if (rowCount == 0)
            return new object[0, 0];
        var colCount = listOfLists.Max(innerList => innerList.Count);

        // 初始化二维数组
        var twoDArray = new object[rowCount, colCount];

        // 遍历每个子列表
        for (var i = 0; i < rowCount; i++)
        {
            var innerList = listOfLists[i];

            for (var j = 0; j < colCount; j++)
            {
                // 如果当前列索引超出子列表长度，补充空值（null 或 ""）
                twoDArray[i, j] = j < innerList.Count ? innerList[j] : null;
            }
        }

        return twoDArray;
    }

    public static object[,] ConvertListToArray(List<List<string>> listOfLists)
    {
        var rowCount = listOfLists.Count;
        if (rowCount == 0)
            return new object[0, 0];
        var colCount = listOfLists.Max(innerList => innerList.Count);

        // 初始化二维数组
        var twoDArray = new object[rowCount, colCount];

        // 遍历每个子列表
        for (var i = 0; i < rowCount; i++)
        {
            var innerList = listOfLists[i];

            for (var j = 0; j < colCount; j++)
            {
                // 如果当前列索引超出子列表长度，补充空值（null 或 ""）
                twoDArray[i, j] = j < innerList.Count ? innerList[j] : null;
            }
        }

        return twoDArray;
    }

    //一维List转二维数组
    public static object[,] ConvertList1ToArrayRow(List<object> listOfLists)
    {
        // 获取列数
        var colCount = listOfLists.Count;

        // 获取最大行数（找出最长的子列表）
        int rowCount = 1;

        // 初始化二维数组
        var twoDArray = new object[rowCount, colCount];

        // 遍历每个子列表
        for (var i = 0; i < colCount; i++)
        {
            var innerList = listOfLists[i];

            twoDArray[rowCount - 1, i] = innerList;
        }

        return twoDArray;
    }

    public static object[,] ConvertList1ToArray(List<string> listOfLists)
    {
        // 获取行数
        var rowCount = listOfLists.Count;

        // 获取最大列数（找出最长的子列表）
        int colCount = 1;

        // 初始化二维数组
        var twoDArray = new string[rowCount, colCount];

        // 遍历每个子列表
        for (var i = 0; i < rowCount; i++)
        {
            var innerList = listOfLists[i];

            twoDArray[i, colCount - 1] = innerList;
        }

        return twoDArray;
    }

    public static string[,] ConvertListArrayToTwoArray(List<string[]> listArray)
    {
        if (listArray.Count == 0)
            return new string[0, 0];
        var rowmax = listArray.Count;
        var colmax = listArray[0].GetLength(0);

        // 初始化二维数组
        var arrObjects = new string[rowmax, colmax];

        for (int i = 0; i < rowmax; i++)
        {
            var list = listArray[i];
            for (int j = 0; j < colmax; j++)
            {
                arrObjects[i, j] = list[j];
            }
        }
        return arrObjects;
    }

    //一维List转一维数组
    public static object[] ConvertListToArray(List<object> listOfLists)
    {
        var rowCount = listOfLists.Count;
        var twoDArray = new object[rowCount];

        for (var i = 0; i < rowCount; i++)
        {
            twoDArray[i] = listOfLists[i];
        }

        return twoDArray;
    }

    public static object[] ConvertListToArray(List<string> listOfLists)
    {
        var rowCount = listOfLists.Count;
        var twoDArray = new object[rowCount];

        for (var i = 0; i < rowCount; i++)
        {
            twoDArray[i] = listOfLists[i];
        }

        return twoDArray;
    }

    public static (int row, int column) FindValueInRangeByVsto(
        Range searchRange,
        object valueToFind
    )
    {
        Range foundRange = searchRange.Find(valueToFind);
        if (foundRange != null)
        {
            return (foundRange.Row, foundRange.Column);
        }
        else
        {
            return (-1, -1);
        }
    }

    public static (List<object> sheetHeaderCol, List<List<object>> sheetData) RangeToListByVsto(
        Range rangeData,
        Range rangeHeader,
        int headRow
    )
    {
        object[,] rangeValue = rangeData.Value;
        object[,] headRangeValue = rangeHeader.Value;
        var sheetData = new List<List<object>>();
        var sheetHeaderCol = new List<object>();
        for (var row = 1; row <= rangeValue.GetLength(0); row++)
        {
            var rowList = new List<object>();
            for (var column = 1; column <= rangeValue.GetLength(1); column++)
            {
                var valueData = rangeValue[row, column];
                rowList.Add(valueData);
            }

            sheetData.Add(rowList);
        }

        for (var column = 1; column <= rangeValue.GetLength(1); column++)
        {
            var value = headRangeValue[headRow, column];
            sheetHeaderCol.Add(value);
        }

        var excelData = (sheetHeaderCol, sheetData);
        return excelData;
    }

    //随机数列表唯一方案
    public static List<List<int>> UniqueRandomMethod(
        int numberOfRolls,
        int numberOfSchemes,
        int maxRand
    )
    {
        var result = new List<List<int>>();
        var seenSchemes = new HashSet<string>();
        var random = new Random();

        for (var i = 0; i < numberOfSchemes; i++)
        {
            var scheme = new List<int>();

            for (var j = 0; j < numberOfRolls; j++)
            {
                var randomNumber = random.Next(1, maxRand + 1);
                scheme.Add(randomNumber);
            }

            var schemeString = string.Join(",", scheme);
            if (seenSchemes.Add(schemeString))
            {
                result.Add(scheme);
            }
        }

        return result;
    }

    //二维数组字典化
    public static Dictionary<int, List<object>> TwoDArrayToDictionary(object[,] array)
    {
        Dictionary<int, List<object>> dictionary = new Dictionary<int, List<object>>();

        int rows = array.GetLength(0);
        int cols = array.GetLength(1);

        for (int i = 0; i < rows; i++)
        {
            List<object> rowArray = new List<object>();
            for (int j = 0; j < cols; j++)
            {
                rowArray.Add(array[i, j]);
            }

            dictionary[i + 1] = rowArray;
        }

        return dictionary;
    }

    //二维数组字典化-首列为Key，0开始
    public static Dictionary<string, List<string>> TwoDArrayToDictionaryFirstKey(object[,] array)
    {
        var dict = new Dictionary<string, List<string>>();
        for (int i = 0; i < array.GetLength(0); i++)
        {
            string key = array[i, 0]?.ToString();
            if (string.IsNullOrEmpty(key))
                continue;

            var row = new List<string>();
            for (int j = 0; j < array.GetLength(1); j++)
                row.Add(array[i, j]?.ToString());
            dict[key] = row;
        }
        return dict;
    }

    //二维数组字典化-首列为Key，0开始
    public static Dictionary<string, string> TwoDArrayToDicFirstKeyStr(object[,] array)
    {
        var dict = new Dictionary<string, string>();
        for (int i = 0; i < array.GetLength(0); i++)
        {
            string key = array[i, 0]?.ToString();
            if (string.IsNullOrEmpty(key))
                continue;

            var cols = new List<string>();
            for (int j = 1; j < array.GetLength(1); j++) // j=0 是 key 列，跳过
                cols.Add(array[i, j]?.ToString() ?? string.Empty);
            dict[key] = string.Join("#", cols);
        }
        return dict;
    }

    //二维数组字典化-首行为Key，0开始
    public static Dictionary<string, List<string>> TwoDArrayToDictionaryFirstRowKey(object[,] array)
    {
        var dict = new Dictionary<string, List<string>>();
        for (int j = 0; j < array.GetLength(1); j++)
        {
            string key = array[0, j]?.ToString();

            var col = new List<string>();

            // 不保存表头数据
            for (int i = 1; i < array.GetLength(0); i++)
                col.Add(array[i, j]?.ToString());
            dict[key] = col;
        }
        return dict;
    }

    //二维数组字典化-首列为Key,ExcelRange对象，1开始
    public static Dictionary<string, List<string>> TwoDArrayToDictionaryFirstKey1(object[,] array)
    {
        var dict = new Dictionary<string, List<string>>();
        for (int i = 1; i <= array.GetLength(0); i++)
        {
            string key = array[i, 1]?.ToString();
            if (string.IsNullOrEmpty(key))
                continue;
            var row = new List<string>();
            for (int j = 1; j <= array.GetLength(1); j++)
                row.Add(array[i, j]?.ToString());
            dict[key] = row;
        }
        return dict;
    }

    //二维数组字典化-首列为Key,ExcelRange对象，1开始
    public static Dictionary<string, string> TwoDArrayToDicFirstKeyStr1(object[,] array)
    {
        var dict = new Dictionary<string, string>();
        for (int i = 1; i <= array.GetLength(0); i++)
        {
            string key = array[i, 1]?.ToString();
            if (string.IsNullOrEmpty(key))
                continue;

            var cols = new List<string>();
            for (int j = 1; j <= array.GetLength(1); j++)
                cols.Add(array[i, j]?.ToString() ?? string.Empty);
            dict[key] = string.Join("#", cols);
        }
        return dict;
    }

    //二维数组字典化-首行为Key,ExcelRange对象，1开始
    public static Dictionary<string, List<string>> TwoDArrayToDictionaryFirstRowKey1(
        object[,] array
    )
    {
        var dict = new Dictionary<string, List<string>>();
        for (int j = 1; j <= array.GetLength(1); j++)
        {
            string key = array[1, j]?.ToString();
            if (string.IsNullOrEmpty(key))
                continue;

            var col = new List<string>();
            for (int i = 1; i <= array.GetLength(0); i++)
                col.Add(array[i, j]?.ToString());
            dict[key] = col;
        }
        return dict;
    }

    //二维数组转二维字典
    public static Dictionary<(object, object), string> Array2DToDic2D(
        int rowCount,
        int colCount,
        object[,] modelRangeValue
    )
    {
        var modelValue = new Dictionary<(object, object), string>();
        for (int row = 2; row <= rowCount; row++)
        {
            for (int col = 2; col <= colCount; col++)
            {
                var rowIndex = modelRangeValue[row, 1];
                var colIndex = modelRangeValue[1, col];
                if (rowIndex == null || colIndex == null)
                {
                    MessageBox.Show(@"模版表中表头有空值，请检查模版数据是否正确！");
                    return new Dictionary<(object, object), string>();
                }

                string value = modelRangeValue[row, col]?.ToString() ?? "";
                modelValue[(rowIndex, colIndex)] = value;
            }
        }

        return modelValue;
    }

    public static Dictionary<(object, object), string> Array2DToDic2D0(
        int rowCount,
        int colCount,
        object[,] modelRangeValue
    )
    {
        object[,] modelRangeValues = (object[,])modelRangeValue;
        var modelValue = new Dictionary<(object, object), string>();
        for (int row = 1; row < rowCount; row++)
        {
            for (int col = 1; col < colCount; col++)
            {
                var rowIndex = modelRangeValues[row, 0];
                var colIndex = modelRangeValues[0, col];
                if (rowIndex == null || colIndex == null)
                {
                    MessageBox.Show(@"模版表中表头有空值，请检查模版数据是否正确！");
                    return null;
                }

                string value = modelRangeValues[row, col]?.ToString() ?? "";
                modelValue[(rowIndex, colIndex)] = value;
            }
        }

        return modelValue;
    }

    //字典二维数组化
    public static object[,] DictionaryTo2DArray<TKey, TValue>(
        Dictionary<TKey, List<TValue>> dictionary,
        int? maxRows = null,
        int? maxCols = null
    )
    {
        int rows = maxRows ?? dictionary.Count;
        int cols = maxCols ?? (dictionary.Values.Max(list => list.Count) + 1);

        object[,] array2D = new object[rows, cols];

        int row = 0;
        foreach (var kvp in dictionary)
        {
            if (row >= rows)
                break;

            for (int col = 0; col < Math.Min(kvp.Value.Count, cols); col++)
            {
                array2D[row, col] = kvp.Value[col];
            }

            row++;
        }

        return array2D;
    }

    //字典二维数组化-带Key(数据模版专用）
    public static object[,] DictionaryTo2DArrayKey<TKey, TValue>(
        Dictionary<TKey, List<TValue>> dictionary,
        int maxRows,
        int maxCols
    )
    {
        object[,] array2D = new object[maxRows, maxCols];

        int row = 0;
        foreach (var kvp in dictionary)
        {
            bool isFirstValue = true;
            foreach (var value in kvp.Value)
            {
                array2D[row, 0] = value;
                array2D[row, 1] = null;
                array2D[row, 2] = isFirstValue ? kvp.Key : null;
                isFirstValue = false;
                row++;
            }
        }

        return array2D;
    }

    //二维数据字符串连接化缩短列数
    public static object[,] ConvertToCommaSeparatedArray(object[,] array2D)
    {
        int rows = array2D.GetLength(0);
        int cols = array2D.GetLength(1);

        string[,] newArray2D = new string[rows, 1];

        for (int i = 0; i < rows; i++)
        {
            List<string> rowElements = new List<string>();
            for (int j = 0; j < cols; j++)
            {
                rowElements.Add(array2D[i, j]?.ToString() ?? "null");
            }

            newArray2D[i, 0] = string.Join(",", rowElements);
        }

        return newArray2D;
    }

    //字典里随机选择若干条数据
    public static Dictionary<int, List<int>> RandChooseDataFormDictionary(
        Dictionary<int, List<int>> sourceDic,
        int chooseCount
    )
    {
        // 将字典的键转换为列表
        List<int> keys = sourceDic.Keys.ToList();
        //不够则有多少取多少
        chooseCount = Math.Min(chooseCount, keys.Count);
        // 使用随机数生成器随机选择 N个键
        Random random = new Random();
        List<int> selectedKeys = keys.OrderBy(x => random.Next()).Take(chooseCount).ToList();
        // 使用选中的键从字典中获取对应的值
        Dictionary<int, List<int>> selectedData = new Dictionary<int, List<int>>();
        foreach (int key in selectedKeys)
        {
            selectedData[key] = sourceDic[key];
        }

        return selectedData;
    }

    //二维数组去重
    public static object[,] CleanRepeatValue(
        object[,] array,
        int index,
        bool isRow,
        int baseIndex,
        bool emptyFilter = true
    )
    {
        var seen = new HashSet<object>(); // 用于存储已出现的基准值
        var tempResult = new List<object[]>(); // 临时存储去重后的结果

        int rows = array.GetLength(0); // 获取行数
        int cols = array.GetLength(1); // 获取列数

        // 检查 baseIndex 是否为 0 或 1
        if (baseIndex != 0 && baseIndex != 1)
        {
            throw new ArgumentException("Base index must be 0 or 1.", nameof(baseIndex));
        }

        // 检查 index 是否超出数组的范围 (根据 baseIndex 调整)
        if (
            index < baseIndex
            || (isRow && index >= cols + baseIndex)
            || (!isRow && index >= rows + baseIndex)
        )
        {
            throw new ArgumentOutOfRangeException(
                nameof(index),
                "Index is outside the bounds of the array."
            );
        }

        // 遍历方向控制
        int outerLoop = isRow ? cols : rows;
        int innerLoop = isRow ? rows : cols;

        for (int i = baseIndex; i < outerLoop + baseIndex; i++) // 根据 baseIndex 调整循环起点
        {
            // 如果 baseIndex 是 1，直接使用 index 和 i；如果是 0，减去 baseIndex
            var key = isRow
                ? array[
                    baseIndex == 1 ? index : index - baseIndex,
                    baseIndex == 1 ? i : i - baseIndex
                ]
                : array[
                    baseIndex == 1 ? i : i - baseIndex,
                    baseIndex == 1 ? index : index - baseIndex
                ];

            if (emptyFilter)
            {
                // 过滤掉 null 和空字符串
                if (key == null || (key is string str && string.IsNullOrWhiteSpace(str)))
                {
                    continue; // 跳过空值
                }
            }

            if (!seen.Contains(key))
            {
                seen.Add(key);
                var row = new object[innerLoop];

                for (int j = baseIndex; j < innerLoop + baseIndex; j++) // 根据 baseIndex 调整循环起点
                {
                    // 检查是否超出数组边界
                    if (
                        isRow && (j - baseIndex >= rows || i - baseIndex >= cols)
                        || !isRow && (i - baseIndex >= rows || j - baseIndex >= cols)
                    )
                    {
                        throw new IndexOutOfRangeException(
                            $"Index out of bounds: i={i}, j={j}, rows={rows}, cols={cols}"
                        );
                    }

                    // 如果按行去重，保留列的值；否则保留行的值
                    row[j - baseIndex] = isRow
                        ? array[
                            baseIndex == 1 ? j : j - baseIndex,
                            baseIndex == 1 ? i : i - baseIndex
                        ]
                        : array[
                            baseIndex == 1 ? i : i - baseIndex,
                            baseIndex == 1 ? j : j - baseIndex
                        ];
                }

                tempResult.Add(row);
            }
        }

        // 将临时结果转换为二维数组
        var result = new object[tempResult.Count, innerLoop];
        for (int i = 0; i < tempResult.Count; i++)
        {
            for (int j = 0; j < innerLoop; j++)
            {
                result[i, j] = tempResult[i][j];
            }
        }

        return result;
    }

    //二维数组复制到剪切板
    public static void CopyArrayToClipboard(object[,] array)
    {
        // 获取数组的行数和列数
        int rows = array.GetLength(0);
        int cols = array.GetLength(1);

        // 构建制表符分隔的字符串
        StringBuilder sb = new StringBuilder();

        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < cols; j++)
            {
                if (array[i, j] != null)
                {
                    sb.Append(array[i, j].ToString());
                }

                // 如果不是最后一列，添加制表符
                if (j < cols - 1)
                {
                    sb.Append("\t");
                }
            }

            // 如果不是最后一行，添加换行符
            if (i < rows - 1)
            {
                sb.AppendLine();
            }
        }

        // 将字符串复制到剪贴板
        Clipboard.SetText(sb.ToString());
    }

    //数组变为二维化字符串
    public static string ArrayToArrayStr(object selectValue)
    {
        var resultStr = string.Empty;

        if (selectValue is object[,])
        {
            // 如果是二维数组
            var values = (object[,])selectValue;
            int rows = values.GetLength(0); // 获取行数
            int cols = values.GetLength(1); // 获取列数

            // 用 StringBuilder 拼接字符串
            var result = new System.Text.StringBuilder();

            for (int i = 1; i <= rows; i++) // 遍历每一行
            {
                for (int j = 1; j <= cols; j++) // 遍历每一列
                {
                    var cellValue = values[i, j] ?? ""; // 获取单元格值，处理空值
                    result.Append(cellValue.ToString()); // 拼接单元格值

                    if (j < cols)
                    {
                        result.Append(","); // 列之间用逗号分隔
                    }
                }

                if (i < rows)
                {
                    result.AppendLine(); // 行之间换行
                }
            }

            resultStr = result.ToString();
        }
        else
        {
            // 如果是单个值
            resultStr = selectValue.ToString();
        }

        return resultStr;
    }

    //Excel多选Range合并为二维数组

    public static object[,] MergeRanges(object[] areas, bool mergeByRow)
    {
        int totalRows = 0;
        int totalCols = 0;

        // 计算合并后的数组大小
        foreach (var area in areas)
        {
            object[,] areaValues = (object[,])area;
            if (mergeByRow)
            {
                totalRows += areaValues.GetLength(0); // 累加行数
                totalCols = Math.Max(totalCols, areaValues.GetLength(1)); // 取最大列数
            }
            else
            {
                totalCols += areaValues.GetLength(1); // 累加列数
                totalRows = Math.Max(totalRows, areaValues.GetLength(0)); // 取最大行数
            }
        }

        // 创建合并后的二维数组
        object[,] mergedArray = new object[totalRows, totalCols];

        // 按行或按列合并数据
        if (mergeByRow)
        {
            int currentRow = 0;
            foreach (var area in areas)
            {
                object[,] areaValues = (object[,])area;
                int areaRows = areaValues.GetLength(0);
                int areaCols = areaValues.GetLength(1);

                for (int i = 0; i < areaRows; i++)
                {
                    for (int j = 0; j < areaCols; j++)
                    {
                        mergedArray[currentRow + i, j] = areaValues[i + 1, j + 1];
                    }
                }

                currentRow += areaRows; // 更新当前行位置
            }
        }
        else
        {
            int currentCol = 0;
            foreach (var area in areas)
            {
                object[,] areaValues = (object[,])area;
                int areaRows = areaValues.GetLength(0);
                int areaCols = areaValues.GetLength(1);

                for (int i = 0; i < areaRows; i++)
                {
                    for (int j = 0; j < areaCols; j++)
                    {
                        mergedArray[i, currentCol + j] = areaValues[i + 1, j + 1];
                    }
                }

                currentCol += areaCols; // 更新当前列位置
            }
        }

        return mergedArray;
    }

    // 查找二维数组中的值，返回行和列的元组
    public static (int, int) FindValueIn2DArray(object[,] array, object value)
    {
        // 获取数组的行和列的起始索引
        int rowStart = array.GetLowerBound(0);
        int colStart = array.GetLowerBound(1);

        // 获取数组的行和列的结束索引
        int rowEnd = array.GetUpperBound(0);
        int colEnd = array.GetUpperBound(1);

        // 遍历数组
        for (int row = rowStart; row <= rowEnd; row++) // 遍历行
        {
            for (int col = colStart; col <= colEnd; col++) // 遍历列
            {
                // 检查是否为空值，并进行比较
                if (array[row, col] != null && array[row, col].ToString() == value.ToString())
                {
                    return (row, col); // 找到值，返回行和列
                }
            }
        }

        return (-1, -1); // 未找到值，返回 (-1, -1)
    }

    #region 自定义数组类型判断

    //检查并解析一维数组
    public static bool IsValidArray(string input, out object[] array)
    {
        array = null;
        if (input.StartsWith("[") && input.EndsWith("]"))
        {
            string content = input.Substring(1, input.Length - 2);
            array = content.Split(',').Select(s => (object)s.Trim()).ToArray();
            return true;
        }

        return false;
    }

    //检查并解析二维数组
    public static bool IsValidArray(string input, out object[][] array)
    {
        array = null;
        // 使用正则表达式验证二维数组的格式
        if (
            Regex.IsMatch(
                input,
                @"^\[\[(?:[^\[\]]+,\s*)*[^\[\]]+\](?:,\s*\[(?:[^\[\]]+,\s*)*[^\[\]]+\])*\]$"
            )
        )
        {
            // 去掉最外层的方括号
            input = input.Trim('[', ']');

            // 分割每一行
            var rows = input.Split(new[] { "],[" }, StringSplitOptions.None);

            // 去掉每一行的方括号
            rows = rows.Select(row => row.Trim('[', ']')).ToArray();

            // 转换为二维数组
            array = rows.Select(row =>
                    row.Split(',').Select(value => (object)value.Trim()).ToArray()
                )
                .ToArray();
            return true;
        }

        return false;
    }

    //检查一维数组中的元素是否为指定类型
    public static bool IsArrayOfType(object[] array, Type type)
    {
        if (array == null || type == null)
        {
            return false;
        }

        foreach (var element in array)
        {
            if (element == null)
            {
                return false;
            }

            try
            {
                // 尝试将元素转换为目标类型
                Convert.ChangeType(element, type);
            }
            catch (InvalidCastException)
            {
                return false;
            }
        }

        return true;
    }

    //检查二维数组中的元素是否为指定类型
    public static bool IsArrayOfType(object[][] array, Type type)
    {
        if (array == null || type == null)
        {
            return false;
        }

        foreach (var row in array)
        {
            foreach (var element in row)
            {
                if (element == null)
                {
                    return false;
                }

                try
                {
                    // 尝试将元素转换为目标类型
                    Convert.ChangeType(element, type);
                }
                catch (InvalidCastException)
                {
                    return false;
                }
            }
        }

        return true;
    }

    // 合并二维数组
    public static object[,] Merge2DArrays0(object[,] array1, object[,] array2)
    {
        // 合并数组
        int rowCount = array1.GetLength(0);
        int baseColCount = array1.GetLength(1);
        int tagColCount = array2.GetLength(1);

        if (rowCount != array2.GetLength(0))
        {
            throw new InvalidOperationException("两个数组的行数不一致，无法合并！");
        }

        object[,] mergedArray = new object[rowCount, baseColCount + tagColCount];

        for (int row = 0; row < rowCount; row++)
        {
            for (int col = 0; col < baseColCount; col++)
            {
                mergedArray[row, col] = array1[row, col];
            }

            for (int col = 0; col < tagColCount; col++)
            {
                mergedArray[row, baseColCount + col] = array2[row, col];
            }
        }

        // 输出合并后的数组（示例）
        PluginLog.Write("合并成功！");
        return mergedArray;
    }

    public static object[,] Merge2DArrays1(object[,] array1, object[,] array2)
    {
        // 合并数组
        int rowCount = array1.GetLength(0);
        int baseColCount = array1.GetLength(1);
        int tagColCount = array2.GetLength(1);

        if (rowCount != array2.GetLength(0))
        {
            throw new InvalidOperationException("两个数组的行数不一致，无法合并！");
        }

        object[,] mergedArray = new object[rowCount, baseColCount + tagColCount];

        for (int row = 0; row < rowCount; row++)
        {
            for (int col = 0; col < baseColCount; col++)
            {
                mergedArray[row, col] = array1[row + 1, col + 1];
            }

            for (int col = 0; col < tagColCount; col++)
            {
                mergedArray[row, baseColCount + col] = array2[row + 1, col + 1];
            }
        }

        // 输出合并后的数组（示例）
        PluginLog.Write("合并成功！");
        return mergedArray;
    }
    #endregion
}
