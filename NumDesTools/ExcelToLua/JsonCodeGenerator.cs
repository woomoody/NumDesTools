using System.Text;
namespace NumDesTools.ExcelToLua
{
    public static class JsonCodeGenerator
{
    public static string ToJsonCode(SheetData data)
    {
        StringBuilder text = new StringBuilder();
        
        text.AppendLine("{");
        
        //行数据转成Table数组
        for (int i = 0; i < data.rows.Count; i++)
        {
            text.Append("\t");
            RowData2Json(data, data.rows[i], ref text);
            if (i < data.rows.Count - 1) text.Append(",");
            text.AppendLine();
        }
        text.AppendLine("}");
        return text.ToString();
    }

    public static string ConfigToJsonCode(SheetData data)
    {
        int keyIndex = data.fields[0].index;
        int devIndex = data.fields[1].index;
        int pubIndex = data.fields[2].index;
        
        StringBuilder text = new StringBuilder();
        text.AppendLine("{");
        for (int i = 0; i < data.rows.Count; i++)
        {
            var cells = data.rows[i].cells;
            string key = cells[keyIndex].value;
            string devValue = cells[devIndex].value;
            string pubValue = cells[pubIndex].value;
            text.Append($"\t\"{key}\":[\"{devValue}\",\"{pubValue}\"]");
            if (i < data.rows.Count - 1) text.Append(",");
            text.AppendLine();
        }
        text.AppendLine("}");
        return text.ToString();
    }
    
    public static string LocalizationToJson(SheetData data)
    {
        StringBuilder text = new StringBuilder();
        var keyField = data.fields[0];
        var valueField = data.fields[1];
        text.AppendLine("{");
        for (int i = 0; i < data.rows.Count; i++)
        {
            var row = data.rows[i];
            string key = row.cells[keyField.index].value;
            string value = row.cells[valueField.index].value;
            if (value.Contains("\"")) value = value.Replace("\"", "\\\"");
            text.Append($"\t\"{key}\" : \"{value}\"");
            if (i < data.rows.Count - 1) text.Append(',');
            text.AppendLine();
        }
        text.AppendLine("}");
        return text.ToString();
    }
    
    private static void RowData2Json(SheetData sheet, RowData row, ref StringBuilder text)
    {
        var cells = row.cells;
        string key = Cell2JsonValue(cells[0], sheet.fields[0]);
        text.Append($"{key}:{{");
        for(int i = 1; i < sheet.fields.Count; i++)
        {
            var field = sheet.fields[i];
            var cell = cells[field.index];
            text.Append($"\"{field.name}\"");
            text.Append(":");
            text.Append(Cell2JsonValue(cell, field));
            if (i < cells.Count - 1) text.Append(",");
        }
        text.Append("}");
    }
    
    private static string Cell2JsonValue(CellData cell, FieldData field)
    {
        switch (field.type)
        {
            case FieldTypeDefine.INT:
            case FieldTypeDefine.LONG:
            case FieldTypeDefine.FLOAT:
            case FieldTypeDefine.DOUBLE:
            case FieldTypeDefine.NUMBER:
                //json number
                return cell.value;
            case FieldTypeDefine.BOOLEAN:
                return cell.value == "true" ? "true" : "false";
            case FieldTypeDefine.STRING:
                if (cell.value.Contains("\"")) return cell.value.Replace("\"", "\\\"");
                return $"\"{cell.value}\"";
            case FieldTypeDefine.INT_ARRAY:
            case FieldTypeDefine.LONG_ARRAY:
            case FieldTypeDefine.FLOAT_ARRAY:
            case FieldTypeDefine.DOUBLE_ARRAY:
            case FieldTypeDefine.BOOL_ARRAY:
            case FieldTypeDefine.STRING_ARRAY:
            case FieldTypeDefine.INT_ARRAY2:
            case FieldTypeDefine.NUMBER_ARRAY:
            case FieldTypeDefine.OBJECT_ARRAY:
            case FieldTypeDefine.OBJECT_ARRAY2:
                if (string.IsNullOrEmpty(cell.value)) return "{}";
                if (cell.value.IndexOf('[') < 0)return $"{{{cell.value}}}";
                return cell.value.Replace('[', '{').Replace(']', '}');
        }
        return string.Empty;
    }
    
    public static string RechargeToJson(SheetData data)
    {
        StringBuilder text = new StringBuilder();
        
        text.AppendLine("[");
        
        //行数据转成Table数组
        for (int i = 0; i < data.rows.Count; i++)
        {
            text.Append("\t");
            RowData2Json2(data, data.rows[i], ref text);
            if (i < data.rows.Count - 1) text.Append(",");
            text.AppendLine();
        }
        text.AppendLine("]");
        return text.ToString();
    }
    
    private static void RowData2Json2(SheetData sheet, RowData row, ref StringBuilder text)
    {
        var cells = row.cells;
        //string key = Cell2JsonValue(cells[0], sheet.fields[0]);
        text.Append("{");
        for(int i = 0; i < sheet.fields.Count; i++)
        {
            var field = sheet.fields[i];
            var cell = cells[field.index];
            text.Append($"\"{field.name}\"");
            text.Append(":");
            text.Append(Cell2JsonValue2(cell, field));
            if (i < cells.Count - 1) text.Append(",");
        }
        text.Append("}");
    }
    
    private static string Cell2JsonValue2(CellData cell, FieldData field)
    {
        switch (field.type)
        {
            case FieldTypeDefine.INT:
            case FieldTypeDefine.LONG:
            case FieldTypeDefine.FLOAT:
            case FieldTypeDefine.DOUBLE:
            case FieldTypeDefine.NUMBER:
                if (field.activeDefaultValue)
                {
                    return field.defaultValue;
                }
                else
                {
                    //json number
                    return cell.value;
                }
            case FieldTypeDefine.BOOLEAN:
                return cell.value == "true" ? "true" : "false";
            case FieldTypeDefine.STRING:
                if (cell.value.Contains("\"")) return cell.value.Replace("\"", "\\\"");
                return $"\"{cell.value}\"";
            case FieldTypeDefine.INT_ARRAY:
            case FieldTypeDefine.LONG_ARRAY:
            case FieldTypeDefine.FLOAT_ARRAY:
            case FieldTypeDefine.DOUBLE_ARRAY:
            case FieldTypeDefine.BOOL_ARRAY:
            case FieldTypeDefine.STRING_ARRAY:
            case FieldTypeDefine.INT_ARRAY2:
            case FieldTypeDefine.NUMBER_ARRAY:
            case FieldTypeDefine.OBJECT_ARRAY:
            case FieldTypeDefine.OBJECT_ARRAY2:
                // if (string.IsNullOrEmpty(cell.value)) return "{}";
                // if (cell.value.IndexOf('[') < 0)return $"{{{cell.value}}}";
                // return cell.value.Replace('[', '{').Replace(']', '}');
                return cell.value;
        }
        return string.Empty;
    }
} }