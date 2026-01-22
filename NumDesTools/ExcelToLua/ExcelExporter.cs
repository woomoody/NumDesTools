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

        //三方支付V2配表 - google
        // ReSharper disable once InconsistentNaming
        private static string RechargeThirdPayV2_GP => "RechargeGP";

        //三方支付V2配表 - ios
        // ReSharper disable once InconsistentNaming
        private static string RechargeThirdPayV2_ios => "RechargeIOS";

        //Config配置表
        static string ConfigExcelFile => "Configs";

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

        // ReSharper disable once RedundantDefaultMemberInitializer
        public static bool NeedMergeLocalization = false;

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

            for (int i = 0; i < files.Length; i++)
            {
                string file = files[i].Replace('\\', '/');
                string fileName = Path.GetFileNameWithoutExtension(file);
                if (fileName.Contains("#") || fileName.Contains("~"))
                    continue;

                var isAll = fileName.Contains("$");
                Export(file, fileName, luaTableFields, isAll, fileName.Contains("$$"));
            }

            if (NeedMergeLocalization)
            {
                MergeLocalizationLuaFile();
            }

            LogDisplay.RecordLine(
                "[{0}] ,{1}",
                DateTime.Now.ToString(CultureInfo.InvariantCulture),
                "导表完成"
            );
            Debug.Print("导表完成!");
        }

        //[MenuItem("Tools/导出Excel(只导出Git变更)", false, 2001)]
        // ReSharper disable once UnusedMember.Local
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
                Export(file, fileName, luaTableFields, isAll, fileName.Contains("$$"));
            }

            if (NeedMergeLocalization)
            {
                MergeLocalizationLuaFile();
            }

            LogDisplay.RecordLine(
                "[{0}] ,{1}",
                DateTime.Now.ToString(CultureInfo.InvariantCulture),
                "导表完成"
            );
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

            // ReSharper disable once UnusedParameter.Local
            process.OutputDataReceived += (sender, e) =>
            {
                var line = e.Data;
                // ReSharper disable once AssignNullToNotNullAttribute
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

        public static void Export(
            string file,
            string fileName,
            List<FieldData> luaTableFields,
            bool isAll,
            bool isIgnoreCheckNullValue
        )
        {
            List<SheetData> list = ExcelReader.Read(file, 1, 1, !isAll);
            if (list.Count == 0 || list[0].fields.Count == 0)
                return;
            int count = 1;
            if (isAll)
            {
                count = list.Count;
            }

            for (int i = 0; i < count; i++)
            {
                var data = list[i];
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
                    ExportLuaLocationTables(data, output);
                    //clear Editor Localization Data
                    //LocalizationManager.Instance.ClearLocalizationData();
                }
                else if (fileName == CLocalizationExcelFile)
                {
                    //localization to json
                    string output = $"{JsonOutputFolder}";
                    ExportCLocationTables(data, output);
                }
                else if (fileName == UiconfigExcelFile)
                {
                    //ui config to lua table
                    string output = $"{LuaOutputFolder}/UIs.lua.txt";
                    if (!ExportLuaTable(output, data, "UIs", data.fields[1], false))
                    {
                        continue;
                    }
                }
                else if (fileName == UiItemconfigExcelFile)
                {
                    //ui config to lua table
                    string output = $"{LuaOutputFolder}/UIAppendItems.lua.txt";
                    if (!ExportLuaTable(output, data, "UIAppendItems", data.fields[1], false))
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

                    string jsonValue = JsonCodeGenerator.ConfigToJsonCode(data);
                    if (!Directory.Exists(output))
                        Directory.CreateDirectory(output);
                    File.WriteAllText(outputWritePath, jsonValue);
                }
                else if (_toJsonExcels.Contains(fileName))
                {
                    //excel to json
                    string output = $"{JsonOutputFolder}";
                    string outputWritePath = $"{output}/{fileName}.json";

                    string jsonValue = JsonCodeGenerator.ToJsonCode(data);
                    if (!Directory.Exists(output))
                        Directory.CreateDirectory(output);
                    File.WriteAllText(outputWritePath, jsonValue);
                }
                else
                {
                    if (data.name.Contains("#"))
                    {
                        continue;
                    }

                    if (!Regex.IsMatch(fileName, "^[a-zA-Z_0-9]+$"))
                    {
                        LogDisplay.RecordLine(
                            "[{0}] , {1}表名非法",
                            DateTime.Now.ToString(CultureInfo.InvariantCulture),
                            fileName
                        );

                        Debug.Print($"配表名称非法 ：<<{fileName}>> 已跳过该表，相关策划需确认");
                        continue;
                    }

                    //excel to lua table
                    string output = $"{LuaOutputFolder}/{fileName}.lua.txt";
                    if (fileName == "Constant")
                    {
                        if (!ExportConstantLuaTable(data, output))
                        {
                            continue;
                        }
                        // ExportConstantLuaTable(data, output);
                    } /*
				else if (data.fields[0].type == FieldTypeDefine.STRING)
				{
					ExportLuaTable(output, data, $"{data.name}s", data.fields[1]);
				}*/
                    else
                    {
                        if (
                            !ExportLuaTable(
                                output,
                                data,
                                $"Tables.{fileName}",
                                null,
                                isIgnoreCheckNullValue
                            )
                        )
                        {
                            continue;
                        }
                    }

                    luaTableFields.Add(new FieldData() { name = fileName, desc = data.desc });
                }

                //三方支付表单独超导一份Json
                if (
                    fileName == RechargeGlobalOfficial
                    || fileName == RechargeThirdPayV2_GP
                    || fileName == RechargeThirdPayV2_ios
                )
                {
                    //config to json
                    string output = $"{JsonOutputFolder}";
                    string outputWritePath = $"{output}/{fileName}.json";

                    string jsonValue = JsonCodeGenerator.RechargeToJson(data);
                    if (!Directory.Exists(output))
                        Directory.CreateDirectory(output);
                    File.WriteAllText(outputWritePath, jsonValue);
                }

                LogDisplay.RecordLine(
                    "[{0}] , {1}完成导表",
                    DateTime.Now.ToString(CultureInfo.InvariantCulture),
                    fileName
                );
                Debug.Print($"{fileName} done.");
            }
        }

        /// <summary>
        /// 导出常量表
        /// </summary>
        /// <param name="data"></param>
        /// <param name="output"></param>
        static bool ExportConstantLuaTable(SheetData data, string output)
        {
            string tableName = $"Tables.{data.name}";

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
            if (fileInfo.Directory is not null && !fileInfo.Directory.Exists)
                fileInfo.Directory.Create();
            FileWriteLuaText(data.name, output, luaTableValue);

            return true;
        }

        static bool ExportLuaTable(
            string outputFile,
            SheetData data,
            string tableName,
            FieldData commentField,
            bool isIgnoreCheckNullValue
        )
        {
            FileInfo file = new FileInfo(outputFile);

            if (file.Directory is not null && !file.Directory.Exists)
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
            }
            luaCheck.Dispose();

            return true;
        }

        /// <summary>
        /// 导出lua语言表table
        /// </summary>
        /// <param name="data"></param>
        /// <param name="output"></param>
        static void ExportLuaLocationTables(SheetData data, string output)
        {
            for (int i = 2; i < data.fields.Count; i++)
            {
                string locationName = $"{data.fields[i].name}";
                string unused = $"{output}/{data.name}{locationName}.lua.txt";

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
                    if (output is not null)
                    {
                        Directory.CreateDirectory(output);
                    }

                File.WriteAllText($"{output}/{data.name}{locationName}.lua.txt", luaTableValue);
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
        static void ExportCLocationTables(SheetData data, string output)
        {
            for (int i = 1; i < data.fields.Count; i++)
            {
                SheetData sub = new SheetData(data.startRow, data.startCol);
                sub.name = data.name + data.fields[i].name;
                string outputFile = $"{output}/{sub.name}.json";

                sub.AddField(data.fields[0]); //key field
                sub.AddField(data.fields[i]); //内容
                sub.rows = data.rows;
                string jsonValue = JsonCodeGenerator.LocalizationToJson(sub);
                if (!Directory.Exists(output))
                    if (output is not null)
                    {
                        Directory.CreateDirectory(output);
                    }

                File.WriteAllText(outputFile, jsonValue);
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
                    LogDisplay.RecordLine(
                        "[{0}] ,配表导出文件无法正确编译，请检查配置。   :{1}",
                        DateTime.Now.ToString(CultureInfo.InvariantCulture),
                        name
                    );

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
                    LogDisplay.RecordLine(
                        "[{0}] , ===多语言存在重复key:{1}",
                        DateTime.Now.ToString(CultureInfo.InvariantCulture),
                        key
                    );

                    Debug.Print("===多语言存在重复key:" + key);
                    for (int i = 0; i < contents.Count; i++)
                    {
                        if (contents[i].Contains($"\"{key}\""))
                        {
                            LogDisplay.RecordLine(
                                "[{0}] , ===重复key字段在表:{1}",
                                DateTime.Now.ToString(CultureInfo.InvariantCulture),
                                fileNames[i]
                            );
                            Debug.Print("===重复key字段在表:" + fileNames[i]);
                        }
                    }
                }
            }

            LogDisplay.RecordLine(
                "[{0}] , ===多语言表总数:",
                DateTime.Now.ToString(CultureInfo.InvariantCulture),
                count
            );

            Debug.Print("===多语言表总数:" + count);
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
                    LogDisplay.RecordLine(
                        "[{0}] , === 多语言值{1}:重复数量:{2}}",
                        DateTime.Now.ToString(CultureInfo.InvariantCulture),
                        item.Key,
                        item.Value
                    );

                    Debug.Print($"=== 多语言值:{item.Key} 重复数量:{item.Value}");
                    count++;
                }
                else
                {
                    break;
                }
            }

            LogDisplay.RecordLine(
                "[{0}] , ===多语言重复数量:{1}",
                DateTime.Now.ToString(CultureInfo.InvariantCulture),
                count
            );

            LogDisplay.RecordLine(
                "[{0}] , ===多语言重复行总数:{1}",
                DateTime.Now.ToString(CultureInfo.InvariantCulture),
                countAll
            );
            Debug.Print("===多语言重复数量:" + count);
            Debug.Print("===多语言重复行总数:" + countAll);
        }
    }
}
