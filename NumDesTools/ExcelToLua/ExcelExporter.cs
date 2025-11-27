using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Lua = NLua.Lua;
using Match = System.Text.RegularExpressions.Match;

namespace NumDesTools.ExcelToLua
{
    public class ExcelExporter
    {
        //excel文件夹

        static string JsonBaseFolder
        {
            get
            {
                var basePath = NumDesAddIn.BasePath;
                string jsonBaseFolder;

                if (
                    basePath.Contains("Lte资源映射")
                    || basePath.Contains("二合")
                    || basePath.Contains("工会")
                    || basePath.Contains("克朗代克")
                )
                {
                    jsonBaseFolder = Path.GetFullPath(Path.Combine(basePath, "./../../../../"));
                }
                else if (
                    basePath.Contains("Configs")
                    || basePath.Contains("UIs")
                    || basePath.Contains("Localizations")
                )
                {
                    jsonBaseFolder = Path.GetFullPath(Path.Combine(basePath, "./../../"));
                }
                else
                {
                    jsonBaseFolder = Path.GetFullPath(Path.Combine(basePath, "./../../../"));
                }
                return jsonBaseFolder.Replace("\\", "/");
            }
        }

        static string LocalizationOutputTempFolder => $"{JsonBaseFolder}Code/Localizations/Lua";

        //lua文件夹
        static string LuaOutputFolder => $"{JsonBaseFolder}Code/Assets/LuaScripts/Tables";

        static string LocalizationOutputFolder =>
            $"{JsonBaseFolder}Code/Asests/LuaScripts/Localizations";

        //json文件夹
        static string JsonOutputFolder => $"{JsonBaseFolder}Code/Assets/Game/Jsons";

        //导出json的excel列表
        static List<string> _toJsonExcels = new List<string>() { "Configs", "LocalizationFonts" };

        //c#多语言表
        static string CLocalizationExcelFile => "LocalizationDefault";

        //lua多语言表
        static string LuaLocalizationExcelFile => "Localizations";

        //UI配置表
        static string UiconfigExcelFile => "UIConfigs";

        //UI配置表
        static string UiItemconfigExcelFile => "UIItemConfigs";

        //三方支付配表
        private static string RechargeGlobalOfficial => "RechargeGlobalOfficial";

        //Config配置表
        static string ConfigExcelFile => "Configs";

        private static string ExcelWriteMd5Path =>
            Path.Combine(NumDesAddIn.BasePath, "./../../public/Excels/");
        private static string ExcelMd5Path => ExcelWriteMd5Path.Replace("\\", "/");

        private static string[] _localizations =
        {
            "ChineseSimplified",
            "English",
            "German",
            "French",
            "Russian",
            "Spanish",
            "PortuguesePortugal",
            "Japanese",
            "Korean",
            "ChineseTraditional",
            "Italian",
        };

        private static string[] _localizationsExcludeFileName =
        {
            "LocalizationDefault",
            "LocalizationFonts",
        };

        public static bool NeedMergeLocalization = false;

        class Md5Info
        {
            public string md5;
            public string path;
        }

        class ExcelMd5Info
        {
            public string md5;
            public List<string> infos;
        }

        private static Dictionary<string, Md5Info> _md5Dir;
        private static Dictionary<string, ExcelMd5Info> _excelMd5Dir;

        public static void ExportAllExcel()
        {
            //bool confirm = EditorUtility.DisplayDialog("导出全部Excel","是否导出全部，耗时长","确定","取消");
            //if (confirm)
            //{
            //	Debug.Log("执行操作");
            //	ExportAll();
            //}
            //else
            //{
            //	Debug.Log("取消操作");
            //}
        }

        public static void ExportAll(string[] files)
        {
            List<FieldData> luaTableFields = new List<FieldData>();
            InitExcelMd5();
            for (int i = 0; i < files.Length; i++)
            {
                string file = files[i].Replace('\\', '/');
                string fileName = Path.GetFileNameWithoutExtension(file);
                if (fileName.Contains("#") || fileName.Contains("~"))
                    continue;

                if (ComparisonMd5(file, fileName, true))
                {
                    continue;
                }
                var isAll = fileName.Contains("$");
                var list = Export(
                    file,
                    fileName,
                    luaTableFields,
                    isAll,
                    fileName.Contains("$$"),
                    false
                );
                if (list != null)
                {
                    SaveExcelMd5(fileName, file, list);
                }
            }

            if (NeedMergeLocalization)
            {
                MergeLocalizationLuaFile();
            }

            Debug.Print("导表完成!");
            SaveAllMd5();
        }

        //[MenuItem("Tools/导出Excel(只导出Git变更)", false, 2001)]
        static void ExportGitChangedExcelFiles()
        {
            var files = GetGitChangedExcelFiles();
            List<FieldData> luaTableFields = new List<FieldData>();
            for (int i = 0; i < files.Count; i++)
            {
                string file = files[i].Replace('\\', '/');
                string fileName = Path.GetFileNameWithoutExtension(file);
                if (fileName.Contains("#") || fileName.Contains("~"))
                    continue;

                var isAll = fileName.Contains("$");
                Export(file, fileName, luaTableFields, isAll, fileName.Contains("$$"), false);
            }

            if (NeedMergeLocalization)
            {
                MergeLocalizationLuaFile();
            }

            Debug.Print("导表完成!");
        }

        static List<string> GetGitChangedExcelFiles()
        {
            string command =
                "git config --global core.quotepath false&cd .&cd ../public/Excels&git status .";
            command = "/c chcp 437&&" + command.Trim().TrimEnd('&') + "&exit";

            Process process = new Process();
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.Arguments = command;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.CreateNoWindow = true;

            List<string> files = new List<string>();
            Regex fileRegex = new Regex(@"(\w+/[^~#].+?\.xlsx?)");

            string basePath = Path.Combine(NumDesAddIn.BasePath, "./../../public/Excels");

            process.OutputDataReceived += (sender, e) =>
            {
                var line = e.Data;
                if (fileRegex.IsMatch(line))
                {
                    var match = fileRegex.Match(line);
                    string file = Path.Combine(basePath, match.Groups[1].Value);
                    files.Add(file);
                }
            };

            //process.ErrorDataReceived += (sender, e) => {};

            process.Start();
            process.BeginOutputReadLine();
            //process.BeginErrorReadLine();
            process.WaitForExit();
            process.Close();

            var results = new List<string>();
            foreach (var file in files)
                if (File.Exists(file))
                    results.Add(file);

            return results;
        }

        #region MD5处理

        static void SaveAllMd5()
        {
            var path = GetExcelMd5Path(true, false);
            if (File.Exists(path))
            {
                File.Delete(path);
            }
            var fileStream = File.Create(path);
            fileStream.Close();
            string[] md5List = new string[_md5Dir.Count];
            int index = 0;
            foreach (var md5 in _md5Dir)
            {
                md5List[index++] = string.Format(
                    "{0}|{1}|{2}",
                    md5.Key,
                    md5.Value.md5,
                    md5.Value.path
                );
            }
            File.WriteAllLines(path, md5List);

            path = GetExcelMd5Path(true, true);
            if (File.Exists(path))
            {
                File.Delete(path);
            }
            fileStream = File.Create(path);
            fileStream.Close();
            md5List = new string[_excelMd5Dir.Count];
            index = 0;
            StringBuilder sb = new StringBuilder();
            foreach (var md5 in _excelMd5Dir)
            {
                foreach (var str in md5.Value.infos)
                {
                    sb.Append($"{str},");
                }
                md5List[index++] = string.Format(
                    "{0}|{1}|{2}",
                    md5.Key,
                    md5.Value.md5,
                    sb.ToString()
                );
                sb.Clear();
            }
            File.WriteAllLines(path, md5List);
        }

        static bool ComparisonMd5(string path, string key, bool isExcel = false)
        {
            if (!File.Exists(path))
            {
                return false;
            }
            string newMd5 = null;
            if (!isExcel)
            {
                if (_md5Dir != null)
                {
                    newMd5 = Md5Helper.FileMd5(path);
                    Md5Info md5Value;
                    if (_md5Dir.TryGetValue(key, out md5Value) && md5Value.md5.Equals(newMd5))
                        return true;
                }
            }
            else
            {
                if (_excelMd5Dir != null)
                {
                    newMd5 = Md5Helper.FileMd5(path);
                    ExcelMd5Info md5Value;
                    if (_excelMd5Dir.TryGetValue(key, out md5Value) && md5Value.md5.Equals(newMd5))
                    {
                        foreach (var info in md5Value.infos)
                        {
                            if (
                                !_md5Dir.ContainsKey(info)
                                || !ComparisonMd5(_md5Dir[info].path, info)
                            )
                            {
                                return false;
                            }
                        }
                        return true;
                    }
                }
            }
            return false;
        }

        static void SaveExcelMd5(string key, string path, List<string> infos)
        {
            if (!_excelMd5Dir.ContainsKey(key))
            {
                _excelMd5Dir[key] = new ExcelMd5Info();
            }

            ExcelMd5Info info = _excelMd5Dir[key];
            info.md5 = Md5Helper.FileMd5(path);
            info.infos = infos;
        }

        static void SaveMd5(string path, string key)
        {
            if (_md5Dir == null)
                return;

            if (!_md5Dir.ContainsKey(key))
            {
                _md5Dir[key] = new Md5Info();
            }

            Md5Info info = _md5Dir[key];
            info.md5 = Md5Helper.FileMd5(path);
            info.path = path;
        }

        static string GetExcelMd5Path(bool isWrite, bool isExcel)
        {
            string fileName = isExcel ? "ExcelRelationPath" : "ExcelMD5Path";
            return $"{(isWrite ? ExcelWriteMd5Path : ExcelMd5Path)}{fileName}.txt";
        }

        static void InitExcelMd5()
        {
            if (_md5Dir != null)
                _md5Dir.Clear();
            else
                _md5Dir = new Dictionary<string, Md5Info>();
            var path = GetExcelMd5Path(false, false);
            var splitStrList = new string[] { "|" };
            var splitStrList1 = new string[] { "," };
            if (File.Exists(path))
            {
                var md5Paths = File.ReadAllLines(path);
                foreach (var md5 in md5Paths)
                {
                    var value = md5.Split(splitStrList, StringSplitOptions.RemoveEmptyEntries);
                    if (value.Length == 3)
                    {
                        _md5Dir[value[0]] = new Md5Info() { md5 = value[1], path = value[2], };
                    }
                }
            }

            if (_excelMd5Dir != null)
                _excelMd5Dir.Clear();
            else
                _excelMd5Dir = new Dictionary<string, ExcelMd5Info>();
            path = GetExcelMd5Path(false, true);
            if (File.Exists(path))
            {
                var md5Paths = File.ReadAllLines(path);
                foreach (var md5 in md5Paths)
                {
                    var value = md5.Split(splitStrList, StringSplitOptions.RemoveEmptyEntries);
                    if (value.Length == 3)
                    {
                        _excelMd5Dir[value[0]] = new ExcelMd5Info()
                        {
                            md5 = value[1],
                            infos = new List<string>(
                                value[2].Split(splitStrList1, StringSplitOptions.RemoveEmptyEntries)
                            ),
                        };
                    }
                }
            }
        }

        #endregion MD5处理

        public static List<string> Export(
            string file,
            string fileName,
            List<FieldData> luaTableFields,
            bool isAll,
            bool isIgnoreCheckNullValue,
            bool isExcelMd5Change
        )
        {
            List<SheetData> list = ExcelReader.Read(file, 1, 1, !isAll, false);
            if (list.Count == 0 || list[0].fields.Count == 0)
                return null;
            int count = 1;
            if (isAll)
            {
                count = list.Count;
            }
            SheetData data = null;
            List<string> infoMd5Key = new List<string>();
            string md5Key = null;
            for (int i = 0; i < count; i++)
            {
                md5Key = null;
                data = list[i];
                if (isAll)
                {
                    fileName = data.name;
                }
                if (
                    Path.GetFileName(Path.GetDirectoryName(file)) == LuaLocalizationExcelFile
                    && !_localizationsExcludeFileName.Contains(fileName)
                )
                {
                    //localization to lua table
                    string output = $"{LocalizationOutputTempFolder}";
                    ExportLuaLocationTables(data, output, isExcelMd5Change, infoMd5Key);
                    //clear Editor Localization Data
                    //LocalizationManager.Instance.ClearLocalizationData();
                }
                else if (fileName == CLocalizationExcelFile)
                {
                    //localization to json
                    string output = $"{JsonOutputFolder}";
                    ExportCLocationTables(data, output, isExcelMd5Change, infoMd5Key);
                }
                else if (fileName == UiconfigExcelFile)
                {
                    //ui config to lua table
                    string output = $"{LuaOutputFolder}/UIs.lua.txt";
                    md5Key = "UIs";
                    if (
                        !ExportLuaTable(
                            output,
                            data,
                            md5Key,
                            data.fields[1],
                            false,
                            isExcelMd5Change,
                            infoMd5Key
                        )
                    )
                    {
                        continue;
                    }
                }
                else if (fileName == UiItemconfigExcelFile)
                {
                    //ui config to lua table
                    string output = $"{LuaOutputFolder}/UIAppendItems.lua.txt";
                    md5Key = "UIAppendItems";
                    if (
                        !ExportLuaTable(
                            output,
                            data,
                            md5Key,
                            data.fields[1],
                            false,
                            isExcelMd5Change,
                            infoMd5Key
                        )
                    )
                    {
                        continue;
                    }
                    // ExportLuaTable(output, data, "UIAppendItems", data.fields[1]);
                }
                else if (fileName == ConfigExcelFile)
                {
                    //config to json
                    string output = $"{JsonOutputFolder}";
                    string outputWritePath = $"{output}/{fileName}.json";
                    if (isExcelMd5Change)
                    {
                        if (ComparisonMd5(outputWritePath, fileName))
                            continue;
                    }

                    string jsonValue = JsonCodeGenerator.ConfigToJsonCode(data);
                    if (!Directory.Exists(output))
                        Directory.CreateDirectory(output);
                    File.WriteAllText(outputWritePath, jsonValue);

                    // MD5Dir[fileName] = MD5Helper.FileMD5(outputWritePath);
                    SaveMd5(outputWritePath, fileName);
                    md5Key = fileName;
                }
                else if (_toJsonExcels.Contains(fileName))
                {
                    //excel to json
                    string output = $"{JsonOutputFolder}";
                    string outputWritePath = $"{output}/{fileName}.json";

                    if (isExcelMd5Change)
                    {
                        if (ComparisonMd5(fileName, outputWritePath))
                        {
                            continue;
                        }
                    }

                    string jsonValue = JsonCodeGenerator.ToJsonCode(data);
                    if (!Directory.Exists(output))
                        Directory.CreateDirectory(output);
                    File.WriteAllText(outputWritePath, jsonValue);
                    SaveMd5(outputWritePath, fileName);
                    md5Key = fileName;
                    // MD5Dir[fileName] = MD5Helper.FileMD5(outputWritePath);
                }
                else
                {
                    if (data.name.Contains("#"))
                    {
                        continue;
                    }

                    if (!Regex.IsMatch(fileName, "^[a-zA-Z_0-9]+$"))
                    {
                        Debug.Print($"配表名称非法 ：<<{fileName}>> 已跳过该表，相关策划需确认");
                        continue;
                    }

                    //excel to lua table
                    string output = $"{LuaOutputFolder}/{fileName}.lua.txt";
                    if (fileName == "Constant")
                    {
                        if (!ExportConstantLuaTable(data, output, isExcelMd5Change))
                        {
                            continue;
                        }
                        md5Key = fileName;
                        // ExportConstantLuaTable(data, output);
                    } /*
				else if (data.fields[0].type == FieldTypeDefine.STRING)
				{
					ExportLuaTable(output, data, $"{data.name}s", data.fields[1]);
				}*/
                    else
                    {
                        md5Key = $"Tables.{fileName}";
                        if (
                            !ExportLuaTable(
                                output,
                                data,
                                md5Key,
                                null,
                                isIgnoreCheckNullValue,
                                isExcelMd5Change,
                                infoMd5Key
                            )
                        )
                        {
                            continue;
                        }
                    }

                    luaTableFields.Add(new FieldData() { name = fileName, desc = data.desc });
                }

                //三方支付表单独超导一份Json
                if (fileName == RechargeGlobalOfficial)
                {
                    //config to json
                    string output = $"{JsonOutputFolder}";
                    string outputWritePath = $"{output}/{fileName}.json";
                    if (isExcelMd5Change)
                    {
                        if (ComparisonMd5(outputWritePath, fileName))
                            continue;
                    }

                    string jsonValue = JsonCodeGenerator.RechargeToJson(data);
                    if (!Directory.Exists(output))
                        Directory.CreateDirectory(output);
                    File.WriteAllText(outputWritePath, jsonValue);

                    // MD5Dir[fileName] = MD5Helper.FileMD5(outputWritePath);
                    SaveMd5(outputWritePath, fileName);
                    md5Key = fileName;
                }

                if (!string.IsNullOrEmpty(md5Key))
                {
                    infoMd5Key.Add(md5Key);
                }

                Debug.Print($"{fileName} done.");
            }

            return infoMd5Key;
        }

        /// <summary>
        /// 导出常量表
        /// </summary>
        /// <param name="data"></param>
        /// <param name="output"></param>
        static bool ExportConstantLuaTable(SheetData data, string output, bool isExcelMd5Change)
        {
            string tableName = $"Tables.{data.name}";

            if (!isExcelMd5Change && ComparisonMd5(tableName, output))
            {
                return false;
            }

            var keyField = data.fields[0];
            var valueField = data.fields[2];
            var commentField = data.fields[1];
            string luaTableValue = LuaCodeGenerator.ToLuaTableByKeyValue(
                data,
                tableName,
                keyField,
                valueField,
                commentField,
                "常量表"
            );
            FileInfo fileInfo = new FileInfo(output);
            if (!fileInfo.Directory.Exists)
                fileInfo.Directory.Create();
            FileWriteLuaText(data.name, output, luaTableValue);
            SaveMd5(output, data.name);
            return true;
        }

        static bool ExportLuaTable(
            string outputFile,
            SheetData data,
            string tableName,
            FieldData commentField,
            bool isIgnoreCheckNullValue,
            bool isExcelMd5Change,
            List<string> infoMd5Key
        )
        {
            FileInfo file = new FileInfo(outputFile);

            if (!isExcelMd5Change && ComparisonMd5(tableName, outputFile))
            {
                return false;
            }
            if (!file.Directory.Exists)
                file.Directory.Create();

            bool isSplitMode = false;
            //判断是否拆表,有无对应字段
            for (int i = 0; i < data.fields.Count; i++)
            {
                if (data.fields[i].name == LuaCodeGenerator.SplitFieldName)
                {
                    isSplitMode = true;
                    break;
                }
            }

            var luaCheck = new Lua();

            if (isSplitMode) // 需要拆分表
            {
                //需要重新整理 output 路径
                string mainTxtPath = $"{data.name}.lua.txt"; // 原始输出文件
                var luaTableValues = LuaCodeGenerator.ToLuaTableSplitMode(
                    data,
                    tableName,
                    isIgnoreCheckNullValue
                );
                foreach (var valuePair in luaTableValues)
                {
                    string curTxtPath = $"{valuePair.Key}.lua.txt"; // 当前输出文件
                    string curOutputFile = outputFile.Replace(mainTxtPath, curTxtPath); //当前输出路径
                    string subTabName = $"Tables.{valuePair.Key}"; // 子表名称
                    FileWriteLuaText(subTabName, curOutputFile, valuePair.Value, luaCheck);
                    SaveMd5(curOutputFile, subTabName);
                    infoMd5Key.Add(subTabName);
                }
            }
            else // 导出单表
            {
                string luaTableValue = LuaCodeGenerator.ToLuaTable(
                    data,
                    tableName,
                    commentField,
                    isIgnoreCheckNullValue
                );
                FileWriteLuaText(tableName, outputFile, luaTableValue, luaCheck);
                SaveMd5(outputFile, tableName);
            }
            luaCheck.Dispose();

            return true;
        }

        /// <summary>
        /// 导出lua语言表table
        /// </summary>
        /// <param name="data"></param>
        /// <param name="output"></param>
        static void ExportLuaLocationTables(
            SheetData data,
            string output,
            bool isExcelMd5Change,
            List<string> keyList
        )
        {
            for (int i = 2; i < data.fields.Count; i++)
            {
                string locationName = $"{data.fields[i].name}";
                string outputFile = $"{output}/{data.name}{locationName}.lua.txt";
                if (!isExcelMd5Change && ComparisonMd5(locationName, outputFile))
                {
                    continue;
                }

                FieldData keyField = data.fields[0]; //key field
                FieldData commentField = data.fields[1]; //注释
                FieldData valueField = data.fields[i]; //内容
                string tableName = $"Localizations.{locationName}";
                string tableDesc = $"本地化配置: {valueField.desc}";
                string luaTableValue = LuaCodeGenerator.ToLuaTableByKeyValue(
                    data,
                    tableName,
                    keyField,
                    valueField,
                    commentField,
                    tableDesc,
                    false
                );
                if (!Directory.Exists(output))
                    Directory.CreateDirectory(output);
                File.WriteAllText($"{output}/{data.name}{locationName}.lua.txt", luaTableValue);
                SaveMd5(outputFile, locationName);
                if (keyList != null)
                {
                    keyList.Add(locationName);
                }
            }

            NeedMergeLocalization = true;
            if (File.Exists($"{LocalizationOutputFolder}/{LuaLocalizationExcelFile}.lua.txt"))
            {
                return;
            }

            //Localizations.lua.txt
            string value = LuaCodeGenerator.FieldsToLuaTable(
                data.fields,
                LuaLocalizationExcelFile,
                2,
                "本地化配置"
            );
            //添加两个自定义方法
            StringBuilder text = new StringBuilder(value);
            text.AppendLine(
                @"
--- 初始化语言表Table, 设置语言不存在的错误打印, 并返回语言key
function Localizations.Init(t)
	setmetatable(t, {__index=function(_t, key)
		local errMsg = ""Localization Key '"" .. key .. ""' is not exist!""
		Debug.LogError(errMsg)
		if IsEditor then Solar.Log.MessageBox(errMsg); end
		return key
	end})
end

-- Lua层本地化语言表数据关联回调全局事件派发(由C#回调回来)(此处自动生成，请不要手动修改！)
function __RELATE_LOCALIZATION_TABLE_DATA()
	if Lang then setmetatable(Lang, nil) end
	local languageName = tostring(SolarRoot.Localization.LanguageName)
	local result, msg = pcall(require, ""Localizations"" .. languageName)
	if not result then Debug.LogError(msg) end
	Lang = Localizations[languageName]
	if not Lang then
		PrintError(string.format(""the language [%s] is not support!"", languageName))
		require(""LocalizationsEnglish"")
		languageName = ""English""
	end
	if SolarRoot.Lua:HasPatchFile(""LocalizationsPatch"") then
		local _, dat = pcall(require, ""LocalizationsPatch"")
		if dat and type(dat) == ""table"" then
			local changes = dat[languageName]
			if changes and type(changes) == ""table"" then
				for k, v in pairs(changes) do
					Lang[k] = v
				end
			end
		end
	end
	Localizations.Init(Lang)
end

__RELATE_LOCALIZATION_TABLE_DATA()"
            );

            File.WriteAllText(
                $"{LocalizationOutputFolder}/{LuaLocalizationExcelFile}.lua.txt",
                text.ToString()
            );
        }

        /// <summary>
        /// 导出c#端语言表
        /// </summary>
        /// <param name="data"></param>
        /// <param name="output"></param>
        static void ExportCLocationTables(
            SheetData data,
            string output,
            bool isExcelMd5Change,
            List<string> keyList
        )
        {
            for (int i = 1; i < data.fields.Count; i++)
            {
                SheetData sub = new SheetData(data.startRow, data.startCol);
                sub.name = data.name + data.fields[i].name;
                string outputFile = $"{output}/{sub.name}.json";
                // if (!isExcelMd5Change && ComparisonMD5(outputFile,outputFile))
                // {
                // 	continue;
                // }
                sub.AddField(data.fields[0]); //key field
                sub.AddField(data.fields[i]); //内容
                sub.rows = data.rows;
                string jsonValue = JsonCodeGenerator.LocalizationToJson(sub);
                if (!Directory.Exists(output))
                    Directory.CreateDirectory(output);
                File.WriteAllText(outputFile, jsonValue);
                SaveMd5(outputFile, outputFile);
                if (keyList != null)
                {
                    keyList.Add(outputFile);
                }
            }
        }

        /// <summary>
        /// 写入lua文件
        /// </summary>
        /// <param name="name"> 表名 </param>
        /// <param name="path"> 保存路径 </param>
        /// <param name="contents"> 存储信息 </param>
        /// <param name="lua"> 传入lua环境，可空若外部循环调用建议自己在循环开始时创建lua循环结束时释放 </param>
        private static void FileWriteLuaText(
            string name,
            string path,
            string contents,
            Lua lua = null
        )
        {
            #region 尝试模拟编译检测

            if (name != "Localizations") // todo 检测黑名单(如果之后需要过滤多个表，再做hash处理)
            {
                var createLua = false; // 是否内部创建lua
                if (lua == null)
                {
                    lua = new Lua();
                    createLua = true;
                }

                try
                {
                    //创建必要的table信息，饱和罗列
                    string s =
                        @$"
			Tables = {{}}	-- Tables 实例化，共用
			{name.Split('_')[0]} = {{}} -- 子表用，父表实例化
			Tables.SetDataTableMetatable = function() end -- 父表用，设置元表
			Tables.SetSubTableMetatable = function() end -- 子表用，设置元表
";
                    string c = s + contents;
                    // Debug.LogError(c);
                    lua.DoString(c);
                }
                catch (Exception e)
                {
                    Debug.Print($"配表导出文件无法正确编译，请检查配置。   : {name}\n{e.Message}"); //
                }
                if (createLua)
                    lua.Dispose();
            }

            #endregion 尝试模拟编译检测

            File.WriteAllText(path, contents);
        }

        public static void MergeLocalizationLuaFile()
        {
            foreach (var language in _localizations)
            {
                string originPath = Path.Combine(
                    LocalizationOutputFolder,
                    $"{LuaLocalizationExcelFile}{language}.lua.txt"
                );
                string[] content = new[] { $"{LuaLocalizationExcelFile}.{language} = {{" };
                string[] files = Directory.GetFiles(
                    LocalizationOutputTempFolder,
                    $"*{language}.lua.txt",
                    SearchOption.AllDirectories
                );
                foreach (var file in files)
                {
                    var tempFile = File.ReadAllLines(file);
                    tempFile = tempFile.Skip(2).ToArray();
                    tempFile = tempFile.Take(tempFile.Length - 1).ToArray();
                    content = content.Concat(tempFile).ToArray();
                }

                content = content.Append("}").ToArray();
                File.WriteAllLines(originPath, content);
            }

            CheckLocalizationLuaDuplicateKeys();
            NeedMergeLocalization = false;
        }

        public static void CheckLocalizationLuaDuplicateKeys()
        {
            string originPath = Path.Combine(
                LocalizationOutputFolder,
                $"{LuaLocalizationExcelFile}English.lua.txt"
            );
            string[] content = File.ReadAllLines(originPath);
            HashSet<string> map = new HashSet<string>(content.Length);
            HashSet<string> duplicatekeys = new HashSet<string>();
            string pattern = @"\[""([^""]+)""\]";
            int count = 0;
            foreach (var line in content)
            {
                Match match = Regex.Match(line, pattern);

                if (match.Success)
                {
                    string key = match.Groups[1].Value;
                    if (map.Contains(key))
                    {
                        duplicatekeys.Add(key);
                    }
                    else
                    {
                        map.Add(key);
                    }

                    count++;
                }
            }

            if (duplicatekeys.Count > 0)
            {
                string[] files = Directory.GetFiles(
                    LocalizationOutputTempFolder,
                    $"*English.lua.txt",
                    SearchOption.AllDirectories
                );
                List<string> contents = new List<string>();
                List<string> fileNames = new List<string>();
                foreach (var file in files)
                {
                    var tempFile = File.ReadAllText(file);
                    contents.Add(tempFile);
                    fileNames.Add(Path.GetFileName(file).Replace("English.lua.txt", ".xlsx"));
                }
                foreach (var key in duplicatekeys)
                {
                    Debug.Print("===key duplicate :" + key);
                    for (int i = 0; i < contents.Count; i++)
                    {
                        if (contents[i].Contains($"\"{key}\""))
                        {
                            Debug.Print("===key duplicate in :" + fileNames[i]);
                        }
                    }
                }
            }
            Debug.Print("===lacalization count:" + count);
        }

        public static void CheckLocalizationLuaDuplicateValues()
        {
            string originPath = Path.Combine(
                LocalizationOutputFolder,
                $"{LuaLocalizationExcelFile}ChineseSimplified.lua.txt"
            );
            string[] content = File.ReadAllLines(originPath);
            Dictionary<string, int> map = new Dictionary<string, int>(content.Length);
            string pattern = @"=\s*""([^""]+)"""; // 匹配 "任意非引号内容"
            int count = 0;
            int countAll = 0;
            foreach (var line in content)
            {
                Match match = Regex.Match(line, pattern);

                if (match.Success)
                {
                    string v = match.Groups[1].Value;
                    if (map.ContainsKey(v))
                    {
                        map[v]++;
                        countAll++;
                    }
                    else
                    {
                        map.Add(v, 1);
                    }
                }
            }

            var sortedDict = map.OrderByDescending(x => x.Value)
                .ToDictionary(x => x.Key, x => x.Value);

            foreach (var item in sortedDict)
            {
                if (item.Value > 1)
                {
                    Debug.Print($"===v:{item.Key} duplicate count:{item.Value}");
                    count++;
                }
                else
                {
                    break;
                }
            }

            Debug.Print("===lacalization  duplicate value count:" + count);
            Debug.Print("===lacalization  duplicate value total count:" + countAll);
        }
    }

    public static class Md5Helper
    {
        /// <summary>
        /// 计算文件的MD5哈希值
        /// </summary>
        /// <param name="filePath">文件完整路径</param>
        /// <returns>32位小写MD5字符串</returns>
        public static string FileMd5(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentException("文件路径不能为空");

            if (!File.Exists(filePath))
                throw new FileNotFoundException($"文件不存在: {filePath}");

            try
            {
                using (var md5 = MD5.Create())
                    using (
                        var stream = new FileStream(
                            filePath,
                            FileMode.Open,
                            FileAccess.Read,
                            FileShare.Read,
                            4096,
                            FileOptions.SequentialScan
                        )
                    )
                    {
                        byte[] hashBytes = md5.ComputeHash(stream);
                        return ByteArrayToHexString(hashBytes);
                    }
            }
            catch (Exception ex)
            {
                throw new Exception($"计算文件MD5时出错: {ex.Message}", ex);
            }
        }

        private static string ByteArrayToHexString(byte[] bytes)
        {
            StringBuilder sb = new StringBuilder(bytes.Length * 2);
            foreach (byte b in bytes)
            {
                sb.Append(b.ToString("x2")); // 小写十六进制
            }
            return sb.ToString();
        }
    }
}
