using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NumDesTools.ExcelToLua
{
    public static class FieldTypeDefine
{
    //value type
    public const int INT = 0;            //int
    public const int LONG = 1;           //long
    public const int FLOAT = 2;     //float
    public const int DOUBLE = 3;         //double
    public const int NUMBER = 4;     //lua number
    public const int BOOLEAN = 5;        //bool
    public const int STRING = 10;        //string

    //reference type
    public const int INT_ARRAY = 100;    //List<int>
    public const int LONG_ARRAY = 101;   //List<long>
    public const int FLOAT_ARRAY = 102;  //List<float>
    public const int DOUBLE_ARRAY = 103; //List<double>
    public const int NUMBER_ARRAY = 104; //lua number array
    public const int BOOL_ARRAY = 105;   //List<bool>
    public const int STRING_ARRAY = 110; //List<string>
    public const int INT_ARRAY2 = 200;   //List<List<int>>
    public const int LUA_TABLE = 300;    //lua table
    public const int OBJECT_ARRAY = 400; //List<object>
    public const int OBJECT_ARRAY2 = 401;//List<List<object>>
    public const int REWARD = 501;       //奖励: List<int>
    public const int REWARD_ARRAY = 502;       //奖励: List<List<int>>
}

public class FieldData
{
    public int col;
    public int type;
    public string name;
    public string desc;
    public int index;
    /// <summary>
    /// 该field激活默认值
    /// </summary>
    public bool activeDefaultValue;
    /// <summary>
    /// 默认值，不是所有的表都需要
    /// </summary>
    public string defaultValue;
}

public class CellData
{
    public int row;
    public int col;
    public string value;

    public CellData(int row, int col, string v)
    {
        this.row = row;
        this.col = col;
        value = v;
    }
}

public class RowData
{
    public List<CellData> cells = new List<CellData>();

    public void AddCell(CellData data)
    {
        cells.Add(data);
    }

    public void AddCell(int row, int col, string value)
    {
        cells.Add(new CellData(row, col, value));
    }
}

public class SheetData
{
    public int startRow;
    public int startCol;
    public List<FieldData> fields = new List<FieldData>();
    public List<RowData> rows = new List<RowData>();
    public string name;
    public string desc;
    /// <summary>
    /// 此表支持默认值
    /// </summary>
    public bool hadDefaultValue = false;

    public SheetData(int startRow, int startCol)
    {
        this.startRow = startRow;
        this.startCol = startCol;
    }

    /// <summary>
    /// 字表构造用
    /// </summary>
    public SheetData(SheetData mainData,int subId)
    {
        startRow = mainData.startRow;
        startCol = mainData.startCol;
        fields = mainData.fields;
        name = $"{mainData.name}_{subId}";
        desc = $"{mainData.desc}_{subId}";
        hadDefaultValue = mainData.hadDefaultValue;
    }

    public void AddField(int col, int fieldType, string fieldName, string fieldDesc,bool isActiveDefaultValue,string defaultValue)
    {
        fields.Add(new FieldData()
        {
            col = col,
            type = fieldType,
            name = fieldName,
            desc = fieldDesc,
            index = fields.Count,
            activeDefaultValue = isActiveDefaultValue,
            defaultValue = defaultValue,
        });
    }

    public void AddField(FieldData field)
    {
        fields.Add(field);
    }

    public List<int> GetFieldCols()
    {
        List<int> cols = new List<int>();
        for (int i = 0; i < fields.Count; i++)
        {
            cols.Add(fields[i].col);
        }
        return cols;
    }

    public void AddRowData(RowData row)
    {
        rows.Add(row);
    }
}

public class ExcelReader
{
    public static List<ISheet> GetSheets(string filepath,bool isLog = true)
    {
        if (isLog)
        {
            Debug.Print($"file Name:{filepath}");
        }

        List<ISheet> list = new List<ISheet>();
        XSSFWorkbook xssfWorkbook;
        using (FileStream file = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            xssfWorkbook = new XSSFWorkbook(file);
        }
		
        for (int i = 0; i < xssfWorkbook.NumberOfSheets; ++i)
        {
            ISheet sheet = xssfWorkbook.GetSheetAt(i);
            if (IsValidSheet(sheet, true))
            {
                list.Add(sheet);
            }
        }

        return list;
    }
	
    public static  string GetCellValue(ISheet sheet, int i, int j)
    {
        return sheet.GetRow(i)?.GetCell(j)?.ToString() ?? "";
    }
    
    public static  bool IsValidSheet(ISheet sheet, bool isClient)
    {
        if (CheckStringChinese(sheet.SheetName))return false;
        if (sheet.GetRow(0) == null) return false;
        if (sheet.GetRow(0).Cells.Count == 0) return false;
        if (isClient && sheet.SheetName.StartsWith("s_")) return false;
        if (!isClient && sheet.SheetName.StartsWith("c_")) return false;
        if (sheet.SheetName.StartsWith("#")) return false;
        return true;
    }
    
    public static  string GetCellValue(IRow row, int i)
    {
        try
        {
            if (row?.GetCell(i)?.CellType == CellType.Formula)
            {
                ICell cell = row.GetCell(i);
                if (cell.CachedFormulaResultType == CellType.Numeric)
                {
                    return cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
                }
                return cell.StringCellValue ?? "";
            }
            return row?.GetCell(i)?.ToString() ?? "";
        }
        catch (Exception e)
        {
            Debug.Print($"{e} ~!: {row.Sheet.SheetName} {row.GetCell(i)}");
            throw;
        }
    }
    
    /// <summary>
    /// 检测字符串是否包含汉字
    /// </summary>
    /// <param name="text"></param>
    /// <returns></returns>
    public static bool CheckStringChinese(string text)
    {
        for (int i = 0; i < text.Length; i++)
        {
            if (text[i] > 127)
                return true;
        }

        return false;
    }
    
    /// <summary>
    /// 配置表中的类型转int
    /// </summary>
    /// <param name="type"></param>
    /// <returns></returns>
    /// <exception cref="Exception"></exception>
    public static int ParseType(string type)
    {
        type = type.ToLower();

        switch (type)
        {
            case "bool[]":
                return FieldTypeDefine.BOOL_ARRAY;
            case "int[]":
                return FieldTypeDefine.INT_ARRAY;
            case "long[]":
                return FieldTypeDefine.LONG_ARRAY;
            case "float[]":
                return FieldTypeDefine.FLOAT_ARRAY;
            case "double[]":
                return FieldTypeDefine.DOUBLE_ARRAY;
            case "string[]":
                return FieldTypeDefine.STRING_ARRAY;
            case "number[]":
                return FieldTypeDefine.NUMBER_ARRAY;
            case "int[][]":
                return FieldTypeDefine.INT_ARRAY2;
            case "int":
                return FieldTypeDefine.INT;
            case "int64":
            case "long":
                return FieldTypeDefine.LONG;
            case "float":
                return FieldTypeDefine.FLOAT;
            case "double":
                return FieldTypeDefine.DOUBLE;
            case "number":
            case "luanumber":
                return FieldTypeDefine.NUMBER;
            case "string":
                return FieldTypeDefine.STRING;
            case "table":
                return FieldTypeDefine.LUA_TABLE;
            case "bool":
            case "boolean":
                return FieldTypeDefine.BOOLEAN;
            case "object[]":
                return FieldTypeDefine.OBJECT_ARRAY;
            case "object[][]":
                return FieldTypeDefine.OBJECT_ARRAY2;
            case "reward[]":
                return FieldTypeDefine.REWARD_ARRAY;
            case "reward":
                return FieldTypeDefine.REWARD;
            default:
                throw new Exception($"不支持此类型: {type}");
        }
    }
    
    public static void ReadSheetFields(ISheet sheet, ref SheetData data)
    {
        // 记录前两行非数值信息 字段名称 字段类型
        int colCount = sheet.GetRow(1).LastCellNum;

        for (int j = data.startCol; j < colCount; j++)
        {
            string fieldName = GetCellValue(sheet, data.startRow + 0, j);
            string fieldType = GetCellValue(sheet, data.startRow + 1, j).Trim();
            string fieldDesc = GetCellValue(sheet, data.startRow + 2, j);
            bool isActiveDefaultValue = false;
            string defaultValue = String.Empty;
            if (fieldType.Contains("="))
            {
                var s = fieldType.Split(new[] {'='}, 2);
                fieldType = s[0].Trim();
                defaultValue = s[1].Trim();
                isActiveDefaultValue = true;
                data.hadDefaultValue = true;
            }

            if (fieldName.Trim().Length == 0 || fieldName.Trim().Length == 0) continue;
            if (fieldName.StartsWith("#")) continue;
            if (fieldName.StartsWith("s_")) continue;

            data.AddField(j, ParseType(fieldType), fieldName, fieldDesc, isActiveDefaultValue, defaultValue);
        }
    }
    
    public static void ReadSheetRows(ISheet sheet, ref SheetData data)
    {
        if (!IsValidSheet(sheet, true) || data.fields.Count == 0)
        {
            return;
        }

        HashSet<string> idset = new HashSet<string>();
        List<int> fieldCols = data.GetFieldCols();
        int maxRow = sheet.LastRowNum;
        for (int i = data.startRow + 3; i <= maxRow; i++)
        {
            IRow row = sheet.GetRow(i);
            RowData rowData = new RowData();
            for (int j = 0; j < fieldCols.Count; j++)
            {
                int col = fieldCols[j];
                string value = GetCellValue(row, col);
                if (j == 0)
                {
                    string id = value.Trim();
                    if (string.IsNullOrEmpty(id))
                    {
                        throw new Exception($"{sheet.SheetName} 第{i+1}行, ID为空!");                        
                    }

                    if (!idset.Add(id))
                    {
                        throw new Exception($"{sheet.SheetName} 第{i+1}行, ID重复:{id}!");
                    }
                }
		        
                CellData cellData = new CellData(i, col, value);
                rowData.AddCell(cellData);
            }

            data.AddRowData(rowData);
        }
    }

    /// <summary>
    /// 读取Excel
    /// </summary>
    /// <param name="excelfile">excel文件路径</param>
    /// <param name="startRow">读表的起始行</param>
    /// <param name="startCol">读表的起始列</param>
    /// <param name="singleSheetMode">是否为单表模式(true:只导出excel的第一个sheet, false:读取所有sheet)</param>
    /// <param name="ignoreRowData">是否忽略行数据(true:只导出字段信息,用于生成cs的类结构代码,false:导出所有行数据)</param>
    /// <returns></returns>
    public static List<SheetData> Read(string excelfile, int startRow = 0, int startCol = 0, bool singleSheetMode = false, bool ignoreRowData = false,bool isLog = true)
    {
        List<ISheet> sheets = GetSheets(excelfile,isLog);
        List<SheetData> list = new List<SheetData>();
        for (int i = 0; i < sheets.Count; i++)
        {
            ISheet sheet = sheets[i];
            SheetData data = new SheetData(startRow, startCol);
            data.name = sheet.SheetName;
            data.desc = GetCellValue(sheet, 0, 1);
            ReadSheetFields(sheet, ref data);
            if(!ignoreRowData) ReadSheetRows(sheet, ref data);
            string filename = singleSheetMode ? Path.GetFileNameWithoutExtension(excelfile) : sheet.SheetName;
            data.name = filename;
            list.Add(data);
            if (singleSheetMode)
                break;
        }

        return list;
    }
}
}