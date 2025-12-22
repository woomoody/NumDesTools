using System.Text;
using System.Text.RegularExpressions;

namespace NumDesTools.ExcelToLua
{
    public static class LuaCodeGenerator
    {
        /// <summary>
        /// 无需做空判断
        /// </summary>
        private static readonly HashSet<string> NotAddCheckNullValue = new HashSet<string>()
        {
            "Tables.ChangeRewardGroupData",
            "Tables.RedPoint",
            "Tables.BlockedTile",
            "Tables.ItemAdsorb",
            "Tables.PropsGiftBagTrigger",
            "Tables.BpBuildData",
            "Tables.BpChestData",
            "Tables.BpCreateData",
            "Tables.BpMergeData",
            "Tables.BpOrderData",
            "Tables.BpCollectData",
            "Tables.EventTriggerData",
            "Tables.EventTriggerLteData",
            "Tables.EventTriggerMustData",
            "Tables.EventTriggerZooData",
            "Tables.ChangeChainData",
            "Tables.UserHierarchyCountry",
            "Tables.UserHierarchyAf",
            "Tables.UserHierarchy",
            "Tables.ItemIdMapping",
            "Tables.ActivityMineBuffDropIdMapping",
            "Tables.RedPointIdMapping",
            "Tables.ScoreIdMapping",
            "Tables.IconIdMapping",
            "Tables.UserGroupActivity",
            "Tables.PopsGiftLteEnergyExtraParam",
            "Tables.AdsCoolingData",
            "Tables.GuideGroupHierarchy",
            "Tables.ActivityClientHierarchyGroupData",
            "Tables.DropIdMapping",
            "Tables.PopsGiftActRewardExtraParam",
            "Tables.ActivityBalloonEffect",
            "Tables.QAConfig",
            "Tables.ActivityBpRewardChangeData",
            "Tables.GuideGroupData",
            "Tables.FlotageSkinSceneData",
            "Tables.EventTriggerLtePortalData",
            "Tables.KlondikeBag",
            "Tables.ActivityPhotoIdMapping",
            "Tables.IconIdChangeByCondition",
            "Tables.SkinRewardConvertData",
            "Tables.HelpTargetIdMapping",
            "Tables.ShopMarketIds",
            "Tables.HelpTargetIdMapping",
            "Tables.Field01Data",
            "Tables.Field03Data",
            "Tables.US_WHAT_W",
            "Tables.US_WHAT_H",
            "Tables.US_WHAT_A",
            "Tables.US_WHAT_T",
            "Tables.US_WHAT_Label",
            "Tables.HelpMustItemIdUI",
            "Tables.EntryGatherInfos",
            "Tables.LteIconIdMapping_",
            "Tables.LteElementPrefabMapping_",
            "Tables.LteStringKeyMapping_",
            "Tables.EventTriggerMainStoryLevelData"
        };

        private static bool IsNotAddCheckNullValue(string tableName)
        {
            return NotAddCheckNullValue.Any(tableName.StartsWith);
        }

        public static string ToLuaTable(
            SheetData data,
            string tableName,
            FieldData commentField = null,
            bool isIgnoreCheckNullValue = false
        )
        {
            StringBuilder text = new StringBuilder();

            #region 文件描述

            if (!string.IsNullOrEmpty(data.desc))
                text.AppendLine($"-- {data.desc}");

            #endregion 文件描述

            #region 参数定义，添加lua注释

            AddLuaAnnotation(text, data, tableName);

            #endregion 参数定义，添加lua注释

            #region 数据内容

            if (data.hadDefaultValue) // 存在默认值的生成模式
                CreateDataSegmentMode2(data, tableName, null, null, ref text, isIgnoreCheckNullValue);
            else // 无默认值的传统生成模式
                CreateDataSegmentMode1(data, tableName, commentField, ref text);

            #endregion 数据内容

            #region 编辑器下id空值检测

            if (
                !isIgnoreCheckNullValue
                && !data.hadDefaultValue
                && !IsNotAddCheckNullValue(tableName)
            ) //该表不包含默认值处理（默认值内部会处理id空）&&不在过滤配置中
            {
                text.AppendLine(@"if IsEditor then");
                var outName = tableName;
                if (outName.Equals("UIAppendItems"))
                {
                    outName = "UIItemConfigs";
                }
                text.AppendLine($"    Tables.CheckNullValue({tableName} , \"{outName}\");");
                text.AppendLine(@"end");
            }

            #endregion 编辑器下id空值检测

            return text.ToString();
        }

        public static string ToLuaTableByKeyValue(
            SheetData data,
            string tableName,
            FieldData keyField,
            FieldData valueField,
            FieldData commentField = null,
            string tableDesc = "",
            bool isAddCheckNullCode = true
        )
        {
            StringBuilder text = new StringBuilder();
            text.AppendLine($"-- {tableDesc}");
            text.AppendLine($"{tableName} = {{");

            for (int i = 0; i < data.rows.Count; i++)
            {
                var row = data.rows[i];
                string key = row.cells[keyField.index].value;
                string comment = row.cells[commentField.index].value;
                string value = Cell2LuaValue(row.cells[valueField.index], valueField); //row.cells[valueField.index].value;

                text.AppendLine($"\t-- {comment}");
                text.AppendLine($"\t[\"{key}\"] = {value},");
            }

            text.AppendLine("}");

            if (isAddCheckNullCode)
            {
                text.AppendLine(@"if IsEditor then");
                text.AppendLine($"    Tables.CheckNullValue({tableName} , \"{tableName}\");");
                text.AppendLine(@"end");
            }

            return text.ToString();
        }

        public static string FieldsToLuaTable(
            List<FieldData> fields,
            string name,
            int startIndex = 1,
            string annotation = ""
        )
        {
            StringBuilder text = new StringBuilder();
            if (!string.IsNullOrEmpty(annotation))
            {
                text.AppendLine($"--{annotation}");
            }

            text.AppendLine($"{name} = {{");
            for (int i = startIndex; i < fields.Count; i++)
            {
                text.AppendLine($"\t-- {fields[i].desc}");
                text.AppendLine($"\t{fields[i].name} = nil,");
            }

            text.AppendLine("}");
            return text.ToString();
        }

        private static void RowData2LuaTable(
            SheetData sheet,
            RowData row,
            FieldData commentField,
            List<FieldData> ignoreFields,
            ref StringBuilder text
        )
        {
            var cells = row.cells;
            //解析key
            string key = Cell2LuaValue(cells[0], sheet.fields[0]);
            if (commentField != null && sheet.fields.Contains(commentField))
            {
                text.AppendLine($"\t-- {cells[commentField.index].value}");
            }

            text.Append($"\t[{key}] = {{");
            //循环解析具体字段
            for (int i = 0; i < sheet.fields.Count; i++)
            {
                var field = sheet.fields[i];
                if (field == commentField)
                    continue;
                if (ignoreFields != null) // 如果有忽略字段，尝试跳过忽略字段
                {
                    bool needContinue = false;
                    foreach (var t in ignoreFields)
                    {
                        if (field == t || field.name == t.name) // 子表过滤的话对象不同，需要判断name
                        {
                            needContinue = true;
                            break;
                        }
                    }
                    if (needContinue)
                        continue;
                }

                #region 解析具体数据列

                // if (sheet.name == "Type" && field.name == "belongMapType") // 断点用测试代码
                // {
                //     Debug.Log("");
                // }

                if (
                    field.activeDefaultValue
                    && (
                        string.IsNullOrEmpty(cells[field.index].value)
                        || cells[field.index].value.Equals(field.defaultValue)
                    )
                )
                    continue; // 支持激活默认值 && 未付当前值的数据 || 当前值等于默认值 直接跳过

                var cell = cells[field.index];
                text.Append(" ");
                text.Append(field.name);
                text.Append(" = ");
                text.Append(Cell2LuaValue(cell, field));

                #endregion 解析具体数据列
                if (i < cells.Count - 1)
                    text.Append(",");
            }

            text.Append("}");
        }

        private static void AddLuaAnnotation(StringBuilder text, SheetData data, string tableName)
        {
            string className = "Excel." + tableName.Replace("Tables.", "");
            text.AppendLine($"---@class {className} @{data.desc}");
            for (int i = 0; i < data.fields.Count; i++)
            {
                var field = data.fields[i];
                if (field.name == SplitFieldName)
                    continue; // 跳过分表字段
                text.AppendLine(
                    $"---@field {field.name} {GetLuaAnnotationType(field)} @{field.desc}"
                );
            }

            text.AppendLine();
        }

        private static string GetLuaAnnotationType(FieldData field)
        {
            switch (field.type)
            {
                case FieldTypeDefine.INT:
                case FieldTypeDefine.LONG:
                case FieldTypeDefine.FLOAT:
                case FieldTypeDefine.DOUBLE:
                case FieldTypeDefine.NUMBER:
                    return "number";
                case FieldTypeDefine.BOOLEAN:
                    return "boolean";
                case FieldTypeDefine.STRING:
                    return "string";
                case FieldTypeDefine.INT_ARRAY:
                case FieldTypeDefine.LONG_ARRAY:
                case FieldTypeDefine.FLOAT_ARRAY:
                case FieldTypeDefine.DOUBLE_ARRAY:
                    return "number[]";
                case FieldTypeDefine.BOOL_ARRAY:
                    return "boolean[]";
                case FieldTypeDefine.STRING_ARRAY:
                    return "string[]";
                case FieldTypeDefine.INT_ARRAY2:
                    return "table";
                case FieldTypeDefine.NUMBER_ARRAY:
                    return "number[]";
                case FieldTypeDefine.LUA_TABLE:
                    return "table";
                case FieldTypeDefine.OBJECT_ARRAY:
                case FieldTypeDefine.OBJECT_ARRAY2:
                    return "table";
                case FieldTypeDefine.REWARD_ARRAY:
                    return "table";
                case FieldTypeDefine.REWARD:
                    return "table";
            }

            return string.Empty;
        }

        private static string Cell2LuaValue(CellData cell, FieldData field)
        {
            //支持字段激活默认值则当前未赋值直接返回
            if (field.activeDefaultValue && string.IsNullOrEmpty(cell.value))
                return String.Empty;

            if (
                field.type != FieldTypeDefine.STRING
                && field.type != FieldTypeDefine.STRING_ARRAY
                && field.type != FieldTypeDefine.LUA_TABLE
                && string.IsNullOrEmpty(cell.value)
            )
            {
                throw new Exception($"{field.desc}字段,第{(cell.row + 1)}行无效!");
            }

            return GetLuaValueByType(cell.value, field.type);
        }

        private static string Field2LuaValue(FieldData field)
        {
            //todo 筛选逻辑

            return GetLuaValueByType(field.defaultValue, field.type);
        }

        private static string GetLuaValueByType(string v, int t)
        {
            switch (t)
            {
                case FieldTypeDefine.INT:
                case FieldTypeDefine.LONG:
                case FieldTypeDefine.FLOAT:
                case FieldTypeDefine.DOUBLE:
                case FieldTypeDefine.NUMBER:
                    //lua number
                    return v;
                case FieldTypeDefine.BOOLEAN:
                    return v == "true" ? "true" : "false";
                case FieldTypeDefine.STRING:
                    v = v.Replace("\\\"", "\"");
                    v = v.Replace("\\n", "\n");
                    if (v.Contains("\"") || v.Contains("\n"))
                        return $"[[{v}]]";
                    return $"\"{v}\"";
                case FieldTypeDefine.STRING_ARRAY:
                    if (string.IsNullOrEmpty(v) || v.Equals("[]"))
                        return "{}";
                    if (v.IndexOf('"') < 0)
                    {
                        string str = v;
                        if (str.IndexOf('[') < 0)
                            str = $"[{v}]";
                        str = $"{str}"
                            .Replace("[", "{\"")
                            .Replace("]", "\"}")
                            .Replace(",", "\",\"");

                        return str;
                    }
                    if (v.IndexOf('[') < 0)
                        return $"{{{v}}}";
                    return v.Replace('[', '{').Replace(']', '}');
                case FieldTypeDefine.INT_ARRAY:
                case FieldTypeDefine.LONG_ARRAY:
                case FieldTypeDefine.FLOAT_ARRAY:
                case FieldTypeDefine.DOUBLE_ARRAY:
                case FieldTypeDefine.BOOL_ARRAY:
                case FieldTypeDefine.INT_ARRAY2:
                case FieldTypeDefine.NUMBER_ARRAY:
                case FieldTypeDefine.OBJECT_ARRAY:
                case FieldTypeDefine.OBJECT_ARRAY2:
                case FieldTypeDefine.REWARD_ARRAY:
                case FieldTypeDefine.REWARD:
                    if (string.IsNullOrEmpty(v))
                        return "{}";
                    if (v.IndexOf('[') < 0)
                        return $"{{{v}}}";
                    string text = v.Replace('[', '{').Replace(']', '}');
                    if (t == FieldTypeDefine.REWARD)
                    {
                        if (!ValidReward(text))
                            throw new Exception($"奖励格式不正确! {v}");
                        return $"__reward({text})";
                    }
                    if (t == FieldTypeDefine.REWARD_ARRAY)
                    {
                        if (!ValidRewardArray(text))
                            throw new Exception($"奖励格式不正确! {v}");
                        return $"__reward2({text})";
                    }
                    return text;
                case FieldTypeDefine.LUA_TABLE:
                    if (string.IsNullOrEmpty(v))
                        return "nil";
                    return v;
            }
            return string.Empty;
        }

        private static Regex _rewardRegex = new Regex(@"^\{\d+(,\s*\d+)*\}$");

        private static bool ValidReward(string value)
        {
            if (string.IsNullOrEmpty(value) || value == "{}")
                return true;

            return _rewardRegex.IsMatch(value);
        }

        private static Regex _rewardArrayRegex = new Regex(
            @"^\{(\{\d+(,\s*\d+)*\})(\s*,\{\d+(,\s*\d+)*\})*\}$"
        );

        private static bool ValidRewardArray(string value)
        {
            if (string.IsNullOrEmpty(value) || value == "{{}}" || value == "{}")
                return true;

            return _rewardArrayRegex.IsMatch(value);
        }

        /// <summary>
        /// 创建数据段 - 模式1 - 传统模式，不携带默认值
        /// </summary>
        private static void CreateDataSegmentMode1(
            SheetData data,
            string tableName,
            FieldData commentField,
            ref StringBuilder text
        )
        {
            string className = "Excel." + tableName.Replace("Tables.", "");
            if (data.fields[0].type != FieldTypeDefine.STRING)
            {
                text.AppendLine(
                    $"---@type table<{GetLuaAnnotationType(data.fields[0])},{className}>"
                );
            }
            text.AppendLine($"{tableName} = {{");
            //行数据转成Table数组
            for (int i = 0; i < data.rows.Count; i++)
            {
                RowData2LuaTable(data, data.rows[i], commentField, null, ref text);
                if (i < data.rows.Count - 1)
                    text.Append(",");
                text.AppendLine();
            }

            text.AppendLine("}");
        }

        /// <summary>
        /// 创建数据段 - 模式2 - 生成table携带默认值相关字段
        /// </summary>
        private static void CreateDataSegmentMode2(
            SheetData data,
            string tableName,
            List<FieldData> ignoreFields,
            string mainTableName,
            ref StringBuilder text,
            bool isIgnoreCheckNullValue
        )
        {
            if (string.IsNullOrEmpty(mainTableName)) // 子表不生成默认值，用主表
            {
                text.AppendLine("\n---数据单元元表");
                text.AppendLine("local dataCellMetaTable = {");
                text.AppendLine("    __index = {");
                for (int i = 0; i < data.fields.Count; i++)
                {
                    if (data.fields[i].activeDefaultValue) // 激活默认值
                    {
                        text.AppendLine(
                            $"        {data.fields[i].name} = {Field2LuaValue(data.fields[i])},"
                        ); // 逐行定义默认值
                    }
                }
                text.AppendLine("    },");
                text.AppendLine("}");
            }
            text.AppendLine("\n---数据表");
            text.AppendLine("local data = {");
            for (int i = 0; i < data.rows.Count; i++) //行数据转成Table数组
            {
                RowData2LuaTable(data, data.rows[i], null, ignoreFields, ref text);
                if (i < data.rows.Count - 1)
                    text.Append(",");
                text.AppendLine();
            }
            text.AppendLine("}");
            text.AppendLine("\n---配置访问表");
            if (string.IsNullOrEmpty(mainTableName)) // 子表不生成代码提示，用主表
            {
                string className = "Excel." + tableName.Replace("Tables.", "");
                if (data.fields[0].type != FieldTypeDefine.STRING)
                {
                    text.AppendLine(
                        $"---@type table<{GetLuaAnnotationType(data.fields[0])},{className}>"
                    );
                }
            }
            text.AppendLine($"{tableName} = {{}}");
            text.AppendLine(
                string.IsNullOrEmpty(mainTableName)
                    ? $"Tables.SetDataTableMetatable({tableName},data,dataCellMetaTable,\"{data.name}\",{(!IsNotAddCheckNullValue(tableName) && !isIgnoreCheckNullValue).ToString().ToLower()})"
                    : $"Tables.SetDataTableMetatable({tableName},data,{mainTableName}._dataCellMetaTable,\"{data.name}\",{(!IsNotAddCheckNullValue(tableName) && !isIgnoreCheckNullValue).ToString().ToLower()})"
            );
        }

        #region 子表模式相关处理

        /// <summary>
        /// 切分关键词
        /// </summary>
        public static readonly string SplitFieldName = "_sub_table_id";

        /// <summary>
        /// 转lua配表文本 - 分子表模式
        /// </summary>
        /// <param name="data"></param>
        /// <param name="tableName"></param>
        /// <returns></returns> 返回值的定义 key 为 data.name ，value为 配表字符串
        public static Dictionary<string, string> ToLuaTableSplitMode(
            SheetData data,
            string tableName,
            bool isIgnoreCheckNullValu
        )
        {
            #region 查找分组字段

            FieldData splitField = null;
            //找到分组列
            for (int i = 0; i < data.fields.Count; i++)
            {
                if (data.fields[i].name == SplitFieldName)
                {
                    splitField = data.fields[i];
                    break;
                }
            }

            if (splitField == null)
            {
                LogDisplay.RecordLine(
                    "[{0}] ,{1}",
                    DateTime.Now.ToString(CultureInfo.InvariantCulture),
                    "需要导出分表的配表未找到约定的分表字段"
                );
                Debug.Print("需要导出分表的配表未找到约定的分表字段");
                return null;
            }

            #endregion

            #region 依据分组序号拆分多个子数据

            Dictionary<int, SheetData> subDataCache = new Dictionary<int, SheetData>();
            for (int i = 0; i < data.rows.Count; i++) // 循环行数据进行数据分组
            {
                if (
                    int.TryParse(
                        Cell2LuaValue(data.rows[i].cells[splitField.index], splitField),
                        out int subId
                    )
                )
                {
                    if (!subDataCache.ContainsKey(subId)) // 无缓存则创建缓存
                        subDataCache.Add(subId, new SheetData(data, subId));
                    //填入数据
                    subDataCache[subId].AddRowData(data.rows[i]); // 写入数据
                }
                else
                {
                    LogDisplay.RecordLine(
                        "[{0}] ,{1}",
                        DateTime.Now.ToString(CultureInfo.InvariantCulture),
                        "分组id值填写错误，string无法正确转换为int"
                    );
                    Debug.Print("分组id值填写错误，string无法正确转换为int");
                }
            }

            #endregion

            Dictionary<string, string> result = new Dictionary<string, string>();
            StringBuilder text = new StringBuilder();

            #region 生成总表数据

            //文件描述
            if (!string.IsNullOrEmpty(data.desc))
                text.AppendLine($"-- {data.desc}");

            //参数定义，添加lua注释，仅总表处理
            AddLuaAnnotation(text, data, tableName);

            //数据内容
            CreateDataSegment_SplitMode_MainData(
                data,
                tableName,
                splitField,
                ref text,
                isIgnoreCheckNullValu
            );
            result.Add(data.name, text.ToString());

            #endregion

            #region 遍历生成字表数据

            List<FieldData> subTableIgnoreFields = new List<FieldData>() { splitField };
            foreach (KeyValuePair<int, SheetData> valuePair in subDataCache)
            {
                text.Clear(); // 清空缓存
                string subTabName = $"Tables.{valuePair.Value.name}"; // 子表名称
                //子表不创建表头，参数定义lua注释
                //数据内容
                CreateDataSegmentMode2(
                    valuePair.Value,
                    subTabName,
                    subTableIgnoreFields,
                    tableName,
                    ref text,
                    isIgnoreCheckNullValu
                );
                result.Add(valuePair.Value.name, text.ToString());
            }

            #endregion

            return result;
        }

        /// <summary>
        /// 创建数据段 - 分表模式 - 仅生成总表数据
        /// </summary>
        private static void CreateDataSegment_SplitMode_MainData(
            SheetData data,
            string tableName,
            FieldData subfields,
            ref StringBuilder text,
            bool isIgnoreCheckNullValu
        )
        {
            text.AppendLine("\n---数据单元元表");
            text.AppendLine("local dataCellMetaTable = {");
            text.AppendLine("    __index = {");
            for (int i = 0; i < data.fields.Count; i++)
            {
                if (data.fields[i].name == SplitFieldName)
                    continue; // 子表列跳过默认赋值
                if (data.fields[i].activeDefaultValue) // 激活默认值
                {
                    text.AppendLine(
                        $"        {data.fields[i].name} = {Field2LuaValue(data.fields[i])},"
                    ); // 逐行定义默认值
                }
            }
            text.AppendLine("    },");
            text.AppendLine("}");
            text.AppendLine("\n---数据表");
            text.AppendLine("local data = {");
            for (int i = 0; i < data.rows.Count; i++) //行数据转成Table数组
            {
                RowData2LuaTable_SplitMode_MainData(data, data.rows[i], subfields, ref text);
                if (i < data.rows.Count - 1)
                    text.Append(",");
                text.AppendLine();
            }
            text.AppendLine("}");
            text.AppendLine("\n---配置访问表");
            string className = "Excel." + tableName.Replace("Tables.", ""); // class引用名
            if (data.fields[0].type != FieldTypeDefine.STRING)
                text.AppendLine(
                    $"---@type table<{GetLuaAnnotationType(data.fields[0])},{className}>"
                );
            text.AppendLine(
                $"{tableName} = {{_subTableCache = {{}}, _dataCellMetaTable = dataCellMetaTable}}"
            );
            text.AppendLine(
                $"Tables.SetSubTableMetatable({tableName},data,\"{data.name}\",{(!IsNotAddCheckNullValue(tableName) && !isIgnoreCheckNullValu).ToString().ToLower()})"
            );
        }

        /// <summary>
        /// 解析行，拆表模式，主表数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <param name="subfields"></param>
        /// <param name="text"></param>
        private static void RowData2LuaTable_SplitMode_MainData(
            SheetData sheet,
            RowData row,
            FieldData subfields,
            ref StringBuilder text
        )
        {
            var cells = row.cells;
            //解析key
            string key = Cell2LuaValue(cells[0], sheet.fields[0]);
            //解析子表id
            string subTableId = Cell2LuaValue(cells[subfields.index], subfields);
            text.Append($"\t[{key}] = {subTableId}");
        }

        #endregion
    }
}
