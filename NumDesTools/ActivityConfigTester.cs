using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NumDesTools.ExcelToLua;
using Lua = NLua.Lua;
using MessageBox = System.Windows.MessageBox;

#pragma warning disable CA1416

namespace NumDesTools;

/// <summary>
/// 活动配置程序级验证器。
///
/// 规则来源：Config/ActivityTableRules.json（AI 分析代码后生成的表-字段-引用映射）。
/// 不需要修改本文件即可扩展校验规则——只需编辑 JSON 配置文件。
///
/// 流程：
///   1. 读取 ActivityTableRules.json，解析表/字段/跨表引用规则
///   2. 从当前工作簿路径推断 Code/Assets/LuaScripts/Tables/ 目录
///   3. Git 有改动的活动 xlsx → ExcelExporter.Export() 重新导表
///   4. 将所有相关 lua.txt 加载进同一个 NLua 实例
///   5. 根据规则动态生成 Lua 校验脚本并执行
///   6. 把 Lua 层错误行号映射回 Excel 行，输出可读报告
/// </summary>
public static class ActivityConfigTester
{
    // ─── 规则配置文件路径（我的文档\NumDesTools\Config\ActivityTableRules.json）──────
    private static string RulesFilePath =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "NumDesTools", "Config", "ActivityTableRules.json"
        );

    // Lua 错误中行号的正则：[string "..."]:42:
    private static readonly Regex LuaErrorLineRegex = new(
        @"\[string[^\]]*\]:(\d+):",
        RegexOptions.Compiled
    );

    // Lua txt 每条数据行的 key 正则：\t[123] = {
    private static readonly Regex LuaRowKeyRegex = new(
        @"^\t\[(\d+)\]",
        RegexOptions.Compiled
    );

    // ─── 规则数据模型（对应 JSON 结构）──────────────────────────────────────────

    private class RulesRoot
    {
        [JsonProperty("tables")]
        public List<TableRule> Tables { get; set; } = new();

        /// <summary>业务 Lua 文件列表，相对于 Code/Assets/LuaScripts/ 目录。</summary>
        [JsonProperty("codeFiles")]
        public List<string> CodeFiles { get; set; } = new();

        /// <summary>业务代码加载后要执行的 Lua 入口代码片段（模拟游戏初始化）。</summary>
        [JsonProperty("entryPoints")]
        public List<string> EntryPoints { get; set; } = new();

        /// <summary>
        /// 针对单个活动 ID 调用业务逻辑的 Lua 模板，{id} 会被替换为实际 ID。
        /// 留空则跳过 ID 级业务代码验证，只做配置规则静态检查。
        /// 示例："ActivityManager:AddGMClientActivity({id}, os.time()+86400)"
        /// </summary>
        [JsonProperty("activityTestEntry")]
        public string ActivityTestEntry { get; set; } = "";

        /// <summary>
        /// 活动类型值 → 对应子表名映射（只做存在性检查，向后兼容）。
        /// typeSubTableRules 存在时优先使用 typeSubTableRules，否则回退到此字段。
        /// </summary>
        [JsonProperty("typeTableMap")]
        public Dictionary<string, string> TypeTableMap { get; set; } = new();

        /// <summary>
        /// 每种 type 的子表深度字段校验规则。
        /// key = ActivityType 数字值字符串，value = 子表规则。
        /// 存在时覆盖 typeTableMap 中同 type 的配置。
        /// </summary>
        [JsonProperty("typeSubTableRules")]
        public Dictionary<string, SubTableRule> TypeSubTableRules { get; set; } = new();

        /// <summary>
        /// 一个 type 需要同时验证多张子表时使用（如 type=106 IslandDecoration）。
        /// </summary>
        [JsonProperty("typeMultiSubTableRules")]
        [JsonConverter(typeof(MultiSubTableRulesConverter))]
        public Dictionary<string, List<MultiSubTableEntry>> TypeMultiSubTableRules { get; set; } = new();

        /// <summary>追加到全局桩的额外 Lua 代码（模拟引擎 API）。</summary>
        [JsonProperty("globalStubExtras")]
        public string GlobalStubExtras { get; set; } = "";
    }

    private class SubTableRule
    {
        [JsonProperty("table")]           public string Table         { get; set; }
        [JsonProperty("lookupField")]     public string LookupField   { get; set; } = "activityID";
        [JsonProperty("sceneValidation")] public bool   SceneValidation { get; set; }
        [JsonProperty("fields")]          public List<FieldRule> Fields { get; set; } = new();
    }

    private class MultiSubTableEntry
    {
        [JsonProperty("table")]       public string Table       { get; set; }
        [JsonProperty("lookupField")] public string LookupField { get; set; } = "activityID";
        [JsonProperty("desc")]        public string Desc        { get; set; }
    }

    // 兼容两种格式：字符串数组（"ActivityBpRewardData.xlsx"）或对象数组（{table, lookupField}）
    private class MultiSubTableRulesConverter : JsonConverter<Dictionary<string, List<MultiSubTableEntry>>>
    {
        public override Dictionary<string, List<MultiSubTableEntry>> ReadJson(
            JsonReader reader, Type objectType, Dictionary<string, List<MultiSubTableEntry>>? existingValue,
            bool hasExistingValue, JsonSerializer serializer)
        {
            var result = new Dictionary<string, List<MultiSubTableEntry>>();
            var obj = JObject.Load(reader);
            foreach (var prop in obj.Properties())
            {
                var entries = new List<MultiSubTableEntry>();
                foreach (var token in prop.Value)
                {
                    if (token.Type == JTokenType.String)
                    {
                        // 字符串：从 xlsx 文件名推断 LuaKey
                        var xlsx = token.Value<string>() ?? "";
                        var name = Path.GetFileNameWithoutExtension(xlsx);
                        entries.Add(new MultiSubTableEntry
                        {
                            Table       = "Tables." + name,
                            LookupField = "activityID",
                            Desc        = name
                        });
                    }
                    else
                    {
                        var e = token.ToObject<MultiSubTableEntry>(serializer);
                        if (e != null) entries.Add(e);
                    }
                }
                result[prop.Name] = entries;
            }
            return result;
        }

        public override void WriteJson(JsonWriter writer, Dictionary<string, List<MultiSubTableEntry>>? value, JsonSerializer serializer)
            => serializer.Serialize(writer, value);
    }

    private class TableRule
    {
        [JsonProperty("name")]      public string Name      { get; set; }
        [JsonProperty("luaKey")]    public string LuaKey    { get; set; }
        [JsonProperty("excelFile")] public string ExcelFile { get; set; }
        [JsonProperty("desc")]      public string Desc      { get; set; }
        [JsonProperty("keyField")]  public string KeyField  { get; set; } = "id";
        [JsonProperty("fields")]    public List<FieldRule> Fields { get; set; } = new();
    }

    private class FieldRule
    {
        [JsonProperty("name")]        public string Name        { get; set; }
        [JsonProperty("desc")]        public string Desc        { get; set; }
        [JsonProperty("required")]    public bool   Required    { get; set; }
        [JsonProperty("type")]        public string Type        { get; set; } = "any";
        [JsonProperty("refTable")]    public string RefTable    { get; set; }
        [JsonProperty("refField")]    public string RefField    { get; set; } = "id";
        [JsonProperty("refIsArray")]  public bool   RefIsArray  { get; set; }
        [JsonProperty("customCheck")] public string CustomCheck { get; set; }
    }

    private record LineMapEntry(string Id, int ExcelDisplayRow);

    // tableName → (byLine, byId) — built once per table, reused across all error lookups
    private record TableLineMaps(
        Dictionary<int, LineMapEntry> ByLine,
        Dictionary<string, LineMapEntry> ById);

    // ─── 公共入口 ─────────────────────────────────────────────────────────────────

    /// <summary>全量测试所有活动表。</summary>
    public static void TestAll() => Run(null, null);

    /// <summary>测试指定活动 ID（逗号分隔）。</summary>
    public static void TestByIds(string idsCsv)
    {
        var ids = idsCsv
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToHashSet();
        Run(ids, null);
    }

    /// <summary>只测试 Git 变更涉及的活动表。</summary>
    public static void TestGitChanged()
    {
        var rules = LoadRules();
        if (rules == null) return;

        var excelPath = NumDesAddIn.App.ActiveWorkbook.FullName;
        var changedFiles = SvnGitTools.GitDiffFileCount(Path.GetDirectoryName(excelPath) ?? excelPath);

        var activityTableNames = rules.Tables.Select(t => t.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (!changedFiles.Any(f => activityTableNames.Contains(Path.GetFileNameWithoutExtension(f))))
        {
            MessageBox.Show("当前 Git 工作区没有活动相关表格的改动。");
            return;
        }

        Run(null, changedFiles, rules);
    }

    // ─── 核心流程 ─────────────────────────────────────────────────────────────────

    private static void Run(HashSet<string> filterIds, List<string> gitChangedFiles, RulesRoot rules = null)
    {
        NumDesAddIn.App.StatusBar = "活动配置验证：读取规则...";

        rules ??= LoadRules();
        if (rules == null) return;

        var excelPath = NumDesAddIn.App.ActiveWorkbook.FullName;

        // 2. 推断 Lua 输出目录
        var luaDir = FindLuaOutputDir(excelPath);
        if (luaDir == null)
        {
            MessageBox.Show(
                "无法定位 Code/Assets/LuaScripts/Tables 目录。\n"
                + "请确认当前工作簿在 public/Excels/Tables/ 下。"
            );
            return;
        }

        // 3. 只导 git 有改动的文件，lua.txt 已存在的表直接复用
        var needExport = gitChangedFiles
            ?? SvnGitTools.GitDiffFileCount(Path.GetDirectoryName(excelPath) ?? "");
        ExportChangedActivityExcels(needExport, rules);

        var report = new StringBuilder();
        report.AppendLine("═══════════════ 活动配置验证报告 ═══════════════");
        report.AppendLine($"时间    ：{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        report.AppendLine($"规则文件：{RulesFilePath}");
        report.AppendLine($"Lua目录 ：{luaDir}");
        report.AppendLine();

        var errorCount = RunLuaValidation(rules, luaDir, filterIds, report);

        report.AppendLine();
        report.AppendLine($"═════ 合计：{errorCount} 个配置问题 ══════");

        NumDesAddIn.App.StatusBar = false;
        ErrorLogCtp.DisposeCtp();
        PluginLog.Write(report.ToString());
        ErrorLogCtp.CreateCtpNormal(report.ToString());

        MessageBox.Show(errorCount > 0
            ? $"发现 {errorCount} 个配置问题，查看右侧报告面板。"
            : "所有活动配置验证通过！");
    }

    // ─── 规则文件加载 ─────────────────────────────────────────────────────────────

    private static RulesRoot LoadRules()
    {
        if (!File.Exists(RulesFilePath))
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(RulesFilePath)!);
                File.WriteAllText(RulesFilePath, DefaultRulesJson, Encoding.UTF8);
                MessageBox.Show(
                    $"已在以下路径生成默认规则配置文件，请按实际项目填写后重新执行验证：\n{RulesFilePath}"
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show($"生成默认规则配置文件失败：{ex.Message}");
            }
            return null;
        }

        try
        {
            var json = File.ReadAllText(RulesFilePath, Encoding.UTF8);
            return JsonConvert.DeserializeObject<RulesRoot>(json);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"读取规则配置失败：{ex.Message}");
            return null;
        }
    }

    private const string DefaultRulesJson = """
        {
          "_comment": [
            "活动配置验证规则映射。由 AI 分析 Code 业务代码后生成，供 ActivityConfigTester.cs 使用。",
            "人工可直接编辑扩展，无需修改 C# 代码。",
            "",
            "── 顶层字段说明 ──",
            "  tables[]           配置表加载规则（按顺序加载，被引用的表要先加载）",
            "  codeFiles[]        需要加载的游戏业务 Lua 文件（相对于 Code/Assets/LuaScripts/ 目录）",
            "  activityTestEntry  针对单个活动ID调用业务逻辑的Lua模板，{id}会被替换为实际ID",
            "                     填写后【验证活动(指定ID)】会用真实业务代码逐ID测试",
            "                     示例：\"ActivityManager:AddGMClientActivity({id}, os.time()+86400)\"",
            "  entryPoints[]      全量初始化时调用的Lua代码片段（无ID过滤时使用）",
            "  globalStubExtras   追加到全局桩的额外 Lua 代码（模拟游戏引擎提供的全局 API）",
            "",
            "── tables[].fields 字段说明 ──",
            "  required     true = 不能为空",
            "  type         期望 Lua 值类型: number / string / table / boolean / any",
            "  refTable     跨表引用检查：该字段的值必须在此 luaKey 表中存在",
            "  refIsArray   true = 字段是数组，对每个元素做引用检查",
            "  customCheck  内联 Lua 片段，变量 row/id/err(field,msg) 可用"
          ],

          "tables": [
            {
              "name": "ActivityClientData",
              "luaKey": "Tables.ActivityClientData",
              "excelFile": "ActivityClientData.xlsx",
              "desc": "客户端活动主表，每条记录对应一个具体活动实例",
              "keyField": "id",
              "fields": [
                { "name": "id",              "desc": "活动唯一ID",   "required": true,  "type": "number" },
                { "name": "type",            "desc": "活动类型枚举", "required": true,  "type": "number" },
                { "name": "isActivityGroup", "desc": "是否活动组",   "required": false, "type": "number" }
              ]
            },
            {
              "name": "ActivityClientHierarchyData",
              "luaKey": "Tables.ActivityClientHierarchyData",
              "excelFile": "ActivityClientHierarchyData.xlsx",
              "desc": "活动等级子表，activityIds 引用主表活动 ID",
              "keyField": "id",
              "fields": [
                { "name": "id",          "desc": "子级唯一ID", "required": true, "type": "number" },
                { "name": "activityIds", "desc": "引用 ActivityClientData 的活动 ID", "required": true, "type": "number",
                  "refTable": "Tables.ActivityClientData", "refIsArray": false }
              ]
            },
            {
              "name": "ActivityClientHierarchyGroupData",
              "luaKey": "Tables.ActivityClientHierarchyGroupData",
              "excelFile": "ActivityClientHierarchyGroupData.xlsx",
              "desc": "活动组顶层表，hierarchyActivityIDs 列出组内各等级 ID",
              "keyField": "id",
              "fields": [
                { "name": "id", "desc": "活动组唯一ID", "required": true, "type": "number" },
                { "name": "hierarchyActivityIDs", "desc": "数组，引用 ActivityClientHierarchyData 的 id 列表",
                  "required": true, "type": "table", "refTable": "Tables.ActivityClientHierarchyData", "refIsArray": true }
              ]
            },
            {
              "name": "ActivityServerData",
              "luaKey": "Tables.ActivityServerData",
              "excelFile": "ActivityServerData.xlsx",
              "desc": "服务端活动配置，id 必须在 ActivityClientData 中存在",
              "keyField": "id",
              "fields": [
                { "name": "id", "desc": "活动ID，必须与 ActivityClientData.id 对应", "required": true, "type": "number",
                  "refTable": "Tables.ActivityClientData", "refIsArray": false }
              ]
            }
          ],

          "_comment_codeFiles": [
            "填写游戏 Code/Assets/LuaScripts/ 下的业务逻辑 Lua 文件（相对路径）。",
            "这些文件会在配置表加载后依次注入同一个 NLua 运行时。",
            "填写后可启用【业务逻辑验证】，让真实游戏代码跑一遍配置数据。",
            "根据 Code 搜索 Tables.ActivityClientData 等关键词可找到应填的文件。",
            "示例（请根据实际项目修改）：",
            "  \"Logics/Controller/Activity/ActivityManager.lua.txt\"",
            "  \"Logics/Controller/Activity/ActivityLogicBase.lua.txt\""
          ],
          "codeFiles": [],

          "_comment_activityTestEntry": [
            "单个活动ID验证入口模板，{id} 会被替换为实际活动ID。",
            "需要配合 codeFiles 中加载的业务代码使用。",
            "填写后【验证活动(指定ID)】按钮会调用真实业务逻辑逐ID测试，报告配置错误。",
            "示例：",
            "  \"ActivityManager:AddGMClientActivity({id}, os.time()+86400)\"",
            "  \"ActivityManager:JudgeHierarchyGroupOpenActivity({id})\""
          ],
          "activityTestEntry": "",

          "_comment_entryPoints": [
            "codeFiles 加载后要执行的 Lua 代码片段（无ID过滤时的全量初始化入口）。",
            "示例：\"ActivityManager:Init()\""
          ],
          "entryPoints": [],

          "_comment_globalStubExtras": [
            "追加到全局桩的额外 Lua 代码，用于模拟游戏引擎提供的全局 API。",
            "内置桩已包含：Debug/print/class/handler/require/ArchiveManager/EventManager/CS/UnityEngine/SolarRoot 等。",
            "若业务代码还用到其他全局对象，在此补充。",
            "示例：\"Localization = { GetText = function(k) return k end }\""
          ],
          "globalStubExtras": ""
        }
        """;

    // ─── 导表 ─────────────────────────────────────────────────────────────────────

    private static void ExportChangedActivityExcels(List<string> changedFiles, RulesRoot rules)
    {
        var activityNames = rules.Tables.Select(t => t.Name)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var file in changedFiles)
        {
            if (file.Contains("#") || file.Contains("~")) continue;
            if (!file.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) continue;

            var name = Path.GetFileNameWithoutExtension(file);
            if (!activityNames.Contains(name)) continue;

            try
            {
                NumDesAddIn.App.StatusBar = $"活动配置验证：重新导表 {name}...";
                ExcelExporter.Export(file, name, new List<FieldData>(), false, false);
            }
            catch (Exception ex)
            {
                LogDisplay.RecordLine($"[{DateTime.Now}] , 导表失败 {name}：{ex.Message}");
                LogDisplay.Show();
            }
        }
    }

    // ─── Lua 加载 + 校验 ──────────────────────────────────────────────────────────

    // Scoped to one validation run; cleared at the top of RunLuaValidation.
    // Caches the set of .prefab filenames present in each bundleDir's immediate subdirectories
    // so LuaCheckFileExists avoids rescanning the same directory for every Item row.
    private static Dictionary<string, HashSet<string>> _subdirPrefabCache = new();

    private static int RunLuaValidation(
        RulesRoot rules,
        string luaDir,
        HashSet<string> filterIds,
        StringBuilder report
    )
    {
        _subdirPrefabCache = new();   // reset per run
        using var lua = new Lua();
        var errorCount = 0;

        // Paths derived once from luaDir
        var luaScriptsRoot = Path.GetFullPath(Path.Combine(luaDir, "..", ".."));
        var codeRoot       = Path.GetFullPath(Path.Combine(luaDir, "..", "..", ".."));

        lua.DoString(BuildGlobalStub(rules));

        var helperFiles = new[]
        {
            @"Logics\DataStructures\Definitions\EnumCmds.lua.txt",
            @"Frameworks\LuaExtensions\LuaStringExtensions.lua.txt",
        };
        foreach (var rel in helperFiles)
        {
            var path = Path.Combine(luaScriptsRoot, rel);
            if (File.Exists(path))
            {
                try { lua.DoString(File.ReadAllText(path, Encoding.UTF8)); }
                catch { /* non-critical */ }
            }
        }

        // Load each table lua.txt and build lineMaps in the same pass (avoids double file-read)
        var lineMaps = new Dictionary<string, TableLineMaps>();
        foreach (var table in rules.Tables)
        {
            var luaFile = Path.Combine(luaDir, $"{table.Name}.lua.txt");
            if (!File.Exists(luaFile))
            {
                report.AppendLine($"[{table.Name}] ⚠ lua.txt 不存在，跳过：{luaFile}");
                continue;
            }

            NumDesAddIn.App.StatusBar = $"活动配置验证：加载 {table.Name}...";
            var luaText = File.ReadAllText(luaFile, Encoding.UTF8);
            lineMaps[table.Name] = BuildLineMap(luaText);

            try
            {
                lua.DoString(luaText);
                report.AppendLine($"[{table.Name}] ✓ 加载通过（{table.Desc}）");
            }
            catch (NLua.Exceptions.LuaException ex)
            {
                errorCount++;
                var lineNum = ExtractLuaLineNumber(ex.Message);
                var loc = ResolveLocation(table.Name, lineNum, lineMaps);
                report.AppendLine($"[{table.Name}] ✗ Lua加载错误 → {loc}");
                report.AppendLine($"    {CleanLuaError(ex.Message)}");
            }
        }

        report.AppendLine();

        // c. 动态生成规则校验脚本并执行（配置表格式 + 跨表引用检查）
        NumDesAddIn.App.StatusBar = "活动配置验证：运行跨表校验...";
        var script = BuildValidationScript(rules, filterIds);
        try
        {
            lua.DoString(script);
            errorCount += CollectLuaErrors(lua, "配置规则校验", lineMaps, report);
        }
        catch (NLua.Exceptions.LuaException ex)
        {
            report.AppendLine($"[校验脚本内部错误] {CleanLuaError(ex.Message)}");
            report.AppendLine("  → 请检查 ActivityTableRules.json 中的 customCheck 语法");
            errorCount++;
        }

        // d. 加载游戏业务 Lua 代码（codeFiles 配置项）
        //    这些文件在加载时就会访问配置表，任何配置缺失/类型错误都会在此暴露

        if (rules.CodeFiles.Count > 0)
        {
            report.AppendLine("── 业务代码加载 ──");
            foreach (var relPath in rules.CodeFiles)
            {
                var codeFile = Path.GetFullPath(Path.Combine(luaScriptsRoot, relPath));
                if (!File.Exists(codeFile))
                {
                    report.AppendLine($"[业务代码] ⚠ 文件不存在，跳过：{codeFile}");
                    continue;
                }

                NumDesAddIn.App.StatusBar = $"活动配置验证：加载业务代码 {Path.GetFileName(codeFile)}...";
                var codeText = File.ReadAllText(codeFile, Encoding.UTF8);
                try
                {
                    lua.DoString(codeText);
                    report.AppendLine($"[业务代码] ✓ {relPath}");
                }
                catch (NLua.Exceptions.LuaException ex)
                {
                    errorCount++;
                    var lineNum = ExtractLuaLineNumber(ex.Message);
                    // 业务代码错误：先尝试从错误消息推断是哪张配置表的哪个 ID 出了问题
                    var configLoc = InferConfigLocation(ex.Message, lineMaps);
                    report.AppendLine($"[业务代码] ✗ {relPath} 加载时出错");
                    report.AppendLine($"    Lua错误：{CleanLuaError(ex.Message)}（第{lineNum}行）");
                    if (configLoc != null)
                        report.AppendLine($"    推断配置问题：{configLoc}");
                    report.AppendLine();
                }
            }
            report.AppendLine();
        }

        // e. 活动配置链追踪 + 子表验证（按 typeTableMap 加载 type 对应子表）
        if (filterIds != null && filterIds.Count > 0)
        {
            // e1. 收集所有需要加载的子表名：
            //     typeSubTableRules（深度校验规则）> typeTableMap（存在性兼容）> 固定辅助表
            var loadedTables = rules.Tables.Select(t => t.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var subTablesLoaded = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // 从 typeSubTableRules 和 typeMultiSubTableRules 中收集所有子表名
            var subTableNamesFromRules = rules.TypeSubTableRules.Values
                .Select(r => r.Table)
                .Concat(rules.TypeMultiSubTableRules.Values.SelectMany(list => list.Select(e => e.Table)))
                .Select(k => k.Contains('.') ? k.Split('.')[1] : k);

            // 从 typeTableMap 收集（向后兼容）
            var subTableNamesFromMap = rules.TypeTableMap.Values
                .Select(k => k.Contains('.') ? k.Split('.')[1] : k);

            // 固定辅助表：LteData/MapDataProto 始终加载（LTE 链验证需要）
            var extraTables = new[] { "LteData", "MapDataProto" };

            // No Distinct() needed — the Contains guards below already prevent double-loading
            var subTableKeys = subTableNamesFromRules
                .Concat(subTableNamesFromMap)
                .Concat(extraTables);

            foreach (var subName in subTableKeys)
            {
                if (loadedTables.Contains(subName) || subTablesLoaded.Contains(subName)) continue;
                var subFile = Path.Combine(luaDir, $"{subName}.lua.txt");
                if (!File.Exists(subFile))
                {
                    report.AppendLine($"[子表] ⚠ {subName}.lua.txt 不存在，跳过相关验证");
                    continue;
                }
                try
                {
                    lua.DoString(File.ReadAllText(subFile, Encoding.UTF8));
                    subTablesLoaded.Add(subName);
                    report.AppendLine($"[子表] ✓ 加载 {subName}");
                }
                catch (NLua.Exceptions.LuaException ex)
                {
                    report.AppendLine($"[子表] ✗ {subName} 加载失败：{CleanLuaError(ex.Message)}");
                }
            }
            if (subTablesLoaded.Count > 0) report.AppendLine();

            // e1b. 加载 LTE 动态映射表（LteIconIdMapping_* / LteElementPrefabMapping_*）
            // 这些表名在 LteData 行里以字符串形式引用，必须预先全部加载才能在追踪脚本里访问
            // Use targeted glob patterns instead of scanning the full directory
            var lteMappingFiles = Directory.GetFiles(luaDir, "LteIconIdMapping_*.lua.txt")
                .Concat(Directory.GetFiles(luaDir, "LteElementPrefabMapping_*.lua.txt"));
            foreach (var mf in lteMappingFiles)
            {
                // Strip both extensions: "Foo.lua.txt" → GetFileNameWithoutExtension → "Foo.lua" → [..^4] → "Foo"
                var nameWithoutTxt = Path.GetFileNameWithoutExtension(mf);
                var mName = nameWithoutTxt.EndsWith(".lua", StringComparison.OrdinalIgnoreCase)
                    ? nameWithoutTxt[..^4]
                    : nameWithoutTxt;
                if (subTablesLoaded.Contains(mName)) continue;
                try
                {
                    lua.DoString(File.ReadAllText(mf, Encoding.UTF8));
                    subTablesLoaded.Add(mName);
                }
                catch { /* 映射表加载失败不阻断主流程 */ }
            }

            // e1c. 加载 LTE 活动私有子表及全局辅助表
            // 全局表：只加载一次
            foreach (var globalTbl in new[] {
                // schema tables must come first — activity-specific tables call Tables.Xxx._dataCellMetaTable
                "Item", "Icon", "Type", "ItemMerge", "ItemBuild", "ItemSpawn", "ItemAdsorb",
                "ItemBomb", "ExploitationDetail", "ExploitationGroup", "Drop", "BpMergeScore",
                "LandmarkBuilding", "Object", "LteMultiScenarioData" })
            {
                TryLoadLuaTable(lua, luaDir, globalTbl, subTablesLoaded);
            }

            // 活动私有表：每个 filterID 一套
            foreach (var fid in filterIds)
            {
                foreach (var prefix in new[] {
                    "Item_", "ItemBuild_", "ItemMerge_", "ItemSpawn_", "ItemAdsorb_",
                    "ItemBomb_", "ExploitationGroup_", "ExploitationDetail_",
                    "Drop_", "Icon_", "Type_", "BpMergeScore_",
                    "RewardGroup_", "Help_", "FindTargetTemplateData_",
                    "PictorialBookItemData_", "ItemExchangePointObstacle_" })
                {
                    TryLoadLuaTable(lua, luaDir, prefix + fid, subTablesLoaded);
                }
            }

            // e2. 注册 C# 文件存在检查函数，供 Lua 脚本校验场景 prefab 文件
            lua.RegisterFunction("_cs_file_exists", null,
                typeof(ActivityConfigTester).GetMethod(nameof(LuaCheckFileExists),
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static));
            // 通过全局变量把 codeRoot 传进去（Lua 字符串常量）
            lua.DoString($"_cs_code_root = [[{codeRoot.Replace('\\', '/')}]]");

            // e3. 运行配置链追踪脚本
            report.AppendLine("── 活动配置链追踪（指定ID） ──");
            lua.DoString("_validation_errors = {}");
            // 同时收集正向追踪日志（每条 ID 追踪成功时写入）
            lua.DoString("_trace_log = {}");
            var traceScript = BuildIdTraceScript(filterIds, rules.TypeTableMap,
                rules.TypeSubTableRules, rules.TypeMultiSubTableRules);
            try
            {
                lua.DoString(traceScript);
            }
            catch (NLua.Exceptions.LuaException ex)
            {
                report.AppendLine($"[追踪脚本内部错误] {CleanLuaError(ex.Message)}");
                errorCount++;
            }
            // 输出正向追踪日志（验证通过的路径）
            var traceLog = lua.GetTable("_trace_log");
            if (traceLog != null)
            {
                for (var i = 1; ; i++)
                {
                    var entry = traceLog[i]?.ToString();
                    if (entry == null) break;
                    report.AppendLine($"[追踪] {entry}");
                }
            }
            var traceErrors = CollectLuaErrors(lua, "配置链追踪", lineMaps, report);
            errorCount += traceErrors;
            if (traceErrors == 0)
                report.AppendLine("✓ 所有指定ID的配置链验证通过");
            report.AppendLine();
        }
        else if (rules.EntryPoints.Count > 0)
        {
            report.AppendLine("── 业务入口执行 ──");
            lua.DoString("_validation_errors = {}");
            foreach (var entry in rules.EntryPoints)
            {
                NumDesAddIn.App.StatusBar = $"活动配置验证：执行入口 {entry[..Math.Min(entry.Length, 40)]}...";
                try
                {
                    lua.DoString(entry);
                    report.AppendLine($"[入口] ✓ {entry}");
                }
                catch (NLua.Exceptions.LuaException ex)
                {
                    errorCount++;
                    var configLoc = InferConfigLocation(ex.Message, lineMaps);
                    report.AppendLine($"[入口] ✗ 执行出错：{entry}");
                    report.AppendLine($"    Lua错误：{CleanLuaError(ex.Message)}");
                    if (configLoc != null)
                        report.AppendLine($"    推断配置问题：{configLoc}");
                    report.AppendLine();
                }
            }
            errorCount += CollectLuaErrors(lua, "业务入口校验", lineMaps, report);
            report.AppendLine();
        }

        return errorCount;
    }

    // ─── 从 _validation_errors Lua 表读取并格式化错误 ────────────────────────────

    private static int CollectLuaErrors(
        Lua lua,
        string phase,
        Dictionary<string, TableLineMaps> lineMaps,
        StringBuilder report
    )
    {
        var count = 0;
        var luaErrors = lua.GetTable("_validation_errors");
        if (luaErrors == null) return 0;

        for (var i = 1; ; i++)
        {
            var entry = luaErrors[i]?.ToString();
            if (entry == null) break;

            // 格式：tableName|id|fieldName|message
            var parts = entry.Split('|', 4);
            if (parts.Length == 4)
            {
                var tbl   = parts[0];
                var id    = parts[1];
                var field = parts[2];
                var code  = parts[3];

                // ERR_TABLE_NOT_LOADED 是警告，不计入错误数
                if (code == "ERR_TABLE_NOT_LOADED")
                {
                    report.AppendLine($"[{phase}] [WARN] 表={tbl} lua.txt 未加载，跳过校验");
                    continue;
                }

                var msg = TranslateErrorCode(code, field);
                var loc = ResolveLocationById(tbl, id, lineMaps);
                report.AppendLine($"[{phase}] [ERROR] 表={tbl} | ID={id} | 字段[{field}]");
                report.AppendLine($"    问题：{msg}");
                report.AppendLine($"    位置：{loc}");
                report.AppendLine();
            }
            else
            {
                report.AppendLine($"[{phase}] {entry}");
            }
            count++;
        }
        return count;
    }

    // 将 Lua 脚本中的 ASCII 占位符翻译为中文描述（避免 Lua 字符串传递中文乱码）
    private static string TranslateErrorCode(string code, string fieldName)
    {
        if (code == "ERR_REQUIRED")
            return $"{fieldName} 不能为空";
        if (code == "ERR_TABLE_NOT_LOADED")
            return "lua.txt 未加载或加载失败，跳过校验";
        if (code.StartsWith("ERR_TYPE:"))
        {
            var p = code.Split(':');
            return p.Length >= 3
                ? $"期望类型 {p[1]}，实际类型 {p[2]}"
                : code;
        }
        if (code.StartsWith("ERR_REF:"))
        {
            // ERR_REF:Tables.Xxx_BY_TYPE:id — sub-table miss for activity type
            if (code.Contains("_BY_TYPE:"))
            {
                var p = code.Split(':', 3);
                var subTable = p.Length > 1 ? p[1].Replace("_BY_TYPE", "") : "?";
                var actId    = p.Length > 2 ? p[2] : "?";
                return $"type对应子表 {subTable} 中不存在 id={actId}，该活动类型缺少子表配置";
            }
            var parts = code.Split(':', 3);
            return parts.Length == 3
                ? $"{fieldName}={parts[2]} 在 {parts[1]} 中不存在"
                : code;
        }
        if (code == "ERR_ID_NOT_FOUND")
            return "该 ID 在配置表中不存在";
        if (code.StartsWith("ERR_NO_VALID_HIERARCHY:"))
            return $"组活动 ID={code.Split(':')[1]} 的所有 hierarchyActivityIDs 均无法找到有效的 ActivityClientData";
        if (code.StartsWith("ERR_MISMATCH_GROUP:"))
            return "isActivityGroup=1 但找不到对应的 ActivityClientHierarchyGroupData 配置";
        if (code.StartsWith("ERR_SCENE_MISSING:"))
        {
            var p = code.Split(':', 3);
            var rowId = p.Length > 1 ? p[1] : "?";
            var asset = p.Length > 2 ? p[2] : "?";
            return $"prefab 文件不存在（id={rowId}，assetName={asset}），请确认 bundleName/assetName 配置正确且资源已提交";
        }
        return code;
    }

    // ─── 从 Lua 错误消息推断配置问题位置 ─────────────────────────────────────────

    // 尝试从 Lua 错误中提取类似 "ActivityClientData[1234]" 的配置引用，定位 Excel 行
    private static readonly Regex LuaConfigRefRegex = new(
        @"ActivityClient\w*\[(\d+)\]|Tables\.(\w+)\[(\d+)\]",
        RegexOptions.Compiled
    );

    private static string InferConfigLocation(
        string errorMsg,
        Dictionary<string, TableLineMaps> lineMaps
    )
    {
        var m = LuaConfigRefRegex.Match(errorMsg);
        if (!m.Success) return null;

        string id, tableName;
        if (m.Groups[1].Success)
        {
            // ActivityClientXxx[id] 模式
            id = m.Groups[1].Value;
            tableName = lineMaps.Keys.FirstOrDefault(k =>
                errorMsg.Contains(k, StringComparison.OrdinalIgnoreCase)) ?? "";
        }
        else
        {
            tableName = m.Groups[2].Value;
            id = m.Groups[3].Value;
        }

        if (string.IsNullOrEmpty(tableName)) return $"活动ID={id}（表名未知）";
        return ResolveLocationById(tableName, id, lineMaps);
    }

    private static void TryLoadLuaTable(Lua lua, string luaDir, string tableName, HashSet<string> loaded)
    {
        if (loaded.Contains(tableName)) return;
        var path = Path.Combine(luaDir, tableName + ".lua.txt");
        if (!File.Exists(path)) return;
        try { lua.DoString(File.ReadAllText(path, Encoding.UTF8)); loaded.Add(tableName); }
        catch { /* non-critical */ }
    }

    // 供 NLua RegisterFunction 使用：检查 codeRoot/bundleName/assetName.prefab 是否存在
    // LandmarkBuilding / Item prefabs live in a subdirectory under bundleDir;
    // _subdirPrefabCache prevents rescanning the same directory for every Item row.
    private static bool LuaCheckFileExists(string codeRoot, string bundleName, string assetName)
    {
        try
        {
            var bundleDir = Path.Combine(
                codeRoot.Replace('/', Path.DirectorySeparatorChar),
                bundleName.Replace('/', Path.DirectorySeparatorChar));
            var fileName = assetName + ".prefab";
            if (File.Exists(Path.Combine(bundleDir, fileName)))
                return true;
            if (!Directory.Exists(bundleDir)) return false;
            if (!_subdirPrefabCache.TryGetValue(bundleDir, out var cached))
            {
                cached = Directory.GetFiles(bundleDir, "*.prefab", SearchOption.AllDirectories)
                    .Select(Path.GetFileName)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);
                _subdirPrefabCache[bundleDir] = cached;
            }
            return cached.Contains(fileName);
        }
        catch { return false; }
    }

    // ─── 针对指定活动 ID 追踪配置链 ────────────────────────────────────────────────
    //
    //   A. ActivityClientData[id] 必须存在（组活动也必须有）
    //   B. ActivityClientHierarchyGroupData[id]（若存在）→ 层级链完整性
    //   C. typeSubTableRules[type]（深度字段校验，优先）或 typeTableMap（存在性兼容）
    //   D. type=2(LTE)：LteData → MapDataProto → prefab 文件
    //   E. typeMultiSubTableRules[type]：多子表存在性
    private static string BuildIdTraceScript(
        HashSet<string> ids,
        Dictionary<string, string> typeTableMap,
        Dictionary<string, SubTableRule> typeSubTableRules,
        Dictionary<string, List<MultiSubTableEntry>> typeMultiSubTableRules)
    {
        var sb = new StringBuilder();
        sb.AppendLine("local _errs = _validation_errors");
        sb.AppendLine("local function _e(tbl, id, field, msg)");
        sb.AppendLine("    table.insert(_errs, tbl..'|'..tostring(id)..'|'..field..'|'..msg)");
        sb.AppendLine("end");
        sb.AppendLine();

        // 构建合并的 type→子表 Lua 表（typeSubTableRules 中没有的 type 才回退 typeTableMap）
        // 格式：[typeNum] = { tbl=Tables.Xxx, lookupField="activityID" }
        // 注意 type=2 由 LTE 专属链处理，不放入通用 map
        // [typeNum] = { tbl=Tables.Xxx, lf='lookupField' }
        // typeSubTableRules entries take priority; typeTableMap only fills gaps
        sb.AppendLine("local _typeMap = {}");
        foreach (var (typeKey, rule) in typeSubTableRules)
        {
            if (typeKey == "2") continue; // LTE chain handled separately
            sb.AppendLine($"_typeMap[{typeKey}] = {{tbl={rule.Table}, lf='{rule.LookupField}'}}");
        }
        foreach (var (typeKey, luaKey) in typeTableMap)
        {
            if (typeKey == "2") continue;
            if (typeSubTableRules.ContainsKey(typeKey)) continue;
            sb.AppendLine($"_typeMap[{typeKey}] = {{tbl={luaKey}, lf='activityID'}}");
        }
        sb.AppendLine();

        // 生成各 type 的深度字段校验函数
        sb.AppendLine(BuildSubTableCheckFunctions(typeSubTableRules));

        // ids 列表
        sb.Append("local _ids = {");
        foreach (var id in ids)
            sb.Append($"{id},");
        sb.AppendLine("}");
        sb.AppendLine();

        sb.AppendLine(@"
local _tlog = _trace_log
local function _tl(msg)
    if #_tlog < 200 then table.insert(_tlog, msg)
    elseif #_tlog == 200 then table.insert(_tlog, '... (trace truncated at 200 entries)') end
end

local function _checkPrefab(tblName, rowId, bundle, asset)
    if not bundle or bundle == '' then
        _e(tblName, rowId, 'bundleName', 'ERR_REQUIRED')
    elseif not asset or asset == '' then
        _e(tblName, rowId, 'assetName', 'ERR_REQUIRED')
    elseif not _cs_file_exists(_cs_code_root, bundle, asset) then
        _e(tblName, rowId, 'assetName', 'ERR_SCENE_MISSING:'..tostring(rowId)..':'..tostring(asset))
    end
end

local _bombRefFields = {'bombType3Target','bombType3IceBlock','bombType4Target','bombType5Target'}

local function _checkMapData(srcTable, srcId, mapId)
    local _mapData = Tables.MapDataProto and Tables.MapDataProto[mapId]
    if not _mapData then
        _e(srcTable, srcId, 'mapData', 'ERR_REF:Tables.MapDataProto:'..tostring(mapId))
        return
    end
    _checkPrefab('MapDataProto', mapId, _mapData.bundleName, _mapData.assetName)
    if _mapData.bundleName and _mapData.bundleName ~= '' and
       _mapData.assetName  and _mapData.assetName  ~= '' and
       _cs_file_exists(_cs_code_root, _mapData.bundleName, _mapData.assetName) then
        _tl('  [OK] scene prefab: mapId='..tostring(mapId)..' asset='..tostring(_mapData.assetName))
    end
end

local function _checkClientData(actId, actData)
    if not actData.type then
        _e('ActivityClientData', actId, 'type', 'ERR_REQUIRED')
        return
    end
    local _typeEntry = _typeMap[actData.type]
    if _typeEntry then
        local _lookupKey = actData[_typeEntry.lf] or actId
        local _sub = _typeEntry.tbl and _typeEntry.tbl[_lookupKey]
        if not _sub then
            _e('ActivityClientData', actId, 'activityID',
                'ERR_REF:'..tostring(_typeEntry.tbl)..'_BY_TYPE:'..tostring(_lookupKey))
        else
            _tl('  [OK] sub-table: type='..tostring(actData.type)..' subId='..tostring(_lookupKey))
            local _deepCheck = _subCheckFuncs and _subCheckFuncs[actData.type]
            if _deepCheck then _deepCheck(actId, _lookupKey, _sub) end
        end
    end
    if actData.type == 2 then
        local _activityID = actData.activityID or actId
        local _lte = Tables.LteData and Tables.LteData[_activityID]
        if not _lte then
            _e('LteData', actId, 'id', 'ERR_REF:Tables.LteData:'..tostring(_activityID))
        else
            _tl('  [OK] LteData['..tostring(_activityID)..'] bpData='..tostring(_lte.bpData or 0))
            -- IconIdMappingTableName / ElementPrefabMappingTableName: 若非空则验证对应 lua.txt 存在
            if _lte.IconIdMappingTableName and _lte.IconIdMappingTableName ~= '' then
                local _iconTbl = Tables[_lte.IconIdMappingTableName]
                if not _iconTbl then
                    _e('LteData', _activityID, 'IconIdMappingTableName',
                        'ERR_REF:Tables.'..tostring(_lte.IconIdMappingTableName)..':0')
                else
                    _tl('  [OK] IconIdMapping: '.._lte.IconIdMappingTableName)
                end
            end
            if _lte.ElementPrefabMappingTableName and _lte.ElementPrefabMappingTableName ~= '' then
                local _elemTbl = Tables[_lte.ElementPrefabMappingTableName]
                if not _elemTbl then
                    _e('LteData', _activityID, 'ElementPrefabMappingTableName',
                        'ERR_REF:Tables.'..tostring(_lte.ElementPrefabMappingTableName)..':0')
                else
                    _tl('  [OK] ElementPrefabMapping: '.._lte.ElementPrefabMappingTableName)
                end
            end
            if _lte.mapData and _lte.mapData ~= 0 then
                _checkMapData('LteData', _activityID, _lte.mapData)
            end
            if type(_lte.mapIdList) == 'table' then
                for _, _mid in ipairs(_lte.mapIdList) do
                    _checkMapData('LteData', _activityID, _mid)
                end
            end
            -- lteType==2: mapIdList entries must exist in LteMultiScenarioData
            if _lte.lteType == 2 and type(_lte.mapIdList) == 'table' then
                for _, _mid in ipairs(_lte.mapIdList) do
                    if _mid ~= 0 and (Tables.LteMultiScenarioData == nil or Tables.LteMultiScenarioData[_mid] == nil) then
                        _e('LteMultiScenarioData', _activityID, 'mapIdList',
                            'ERR_REF:Tables.LteMultiScenarioData:'..tostring(_mid))
                    end
                end
            end
            -- ── Item 子系统完整验证 ──
            local _actId_s   = tostring(_activityID)
            -- 分表：遍历本期新增数据用
            local _itemTbl        = Tables['Item_'.._actId_s]
            local _mergeTbl       = Tables['ItemMerge_'.._actId_s]
            local _buildTbl       = Tables['ItemBuild_'.._actId_s]
            local _spawnTbl       = Tables['ItemSpawn_'.._actId_s]
            local _adsorbTbl      = Tables['ItemAdsorb_'.._actId_s]
            local _exploitDetailTbl = Tables['ExploitationDetail_'.._actId_s]
            -- 全局总表：所有跨表引用的 ID 存在性检查必须查总表（包含历史期数据）
            local _gItem          = Tables.Item
            local _gMerge         = Tables.ItemMerge
            local _gBuild         = Tables.ItemBuild
            local _gSpawn         = Tables.ItemSpawn
            local _gBomb          = Tables.ItemBomb
            local _gBpScore       = Tables.BpMergeScore
            local _gExploitDetail = Tables.ExploitationDetail
            local _gDrop          = Tables.Drop
            local _gIcon          = Tables.Icon
            local _gType          = Tables.Type
            -- 模版表中明确不需要 Icon 的物品类型（兑-产/兑-产-障碍无 Icon 行）
            -- 由于 lua 无法在运行时区分这两类，对 Icon 缺失统一降级为 WARNING
            local function _warnMissingIcon(tblName, iid)
                _tl('  [WARN] '..tblName..'['..tostring(iid)..'] 在 Icon 总表中无对应行（兑-产类可忽略）')
            end

            if _itemTbl then
                local _itemCnt = 0
                for _iid, _item in pairs(_itemTbl) do
                    if type(_item) ~= 'table' then goto _item_next end
                    _itemCnt = _itemCnt + 1
                    _checkPrefab('Item_'.._actId_s, _iid, _item.bundleName, _item.assetName)
                    if _gIcon and _gIcon[_iid] == nil then
                        _warnMissingIcon('Item_'.._actId_s, _iid)
                    end
                    if _gType and _gType[_iid] == nil then
                        _e('Item_'.._actId_s, _iid, 'id(Type)', 'ERR_REF:Type:'..tostring(_iid))
                    end
                    if _item.item_merge and _item.item_merge ~= 0 then
                        if _gMerge == nil or _gMerge[_item.item_merge] == nil then
                            _e('Item_'.._actId_s, _iid, 'item_merge',
                                'ERR_REF:ItemMerge:'..tostring(_item.item_merge))
                        end
                    end
                    if _item.item_build and _item.item_build ~= 0 then
                        if _gBuild == nil or _gBuild[_item.item_build] == nil then
                            _e('Item_'.._actId_s, _iid, 'item_build',
                                'ERR_REF:ItemBuild:'..tostring(_item.item_build))
                        end
                    end
                    if _item.item_bomb and _item.item_bomb ~= 0 then
                        if _gBomb == nil or _gBomb[_item.item_bomb] == nil then
                            _e('Item_'.._actId_s, _iid, 'item_bomb',
                                'ERR_REF:ItemBomb:'..tostring(_item.item_bomb))
                        end
                    end
                    if _item.exploitation_id and _item.exploitation_id ~= 0 then
                        if _gExploitDetail == nil or _gExploitDetail[_item.exploitation_id] == nil then
                            _e('Item_'.._actId_s, _iid, 'exploitation_id',
                                'ERR_REF:ExploitationDetail:'..tostring(_item.exploitation_id))
                        end
                    end
                    if type(_item.item_spawn) == 'table' then
                        for _, _sid in ipairs(_item.item_spawn) do
                            if _sid ~= 0 and (_gSpawn == nil or _gSpawn[_sid] == nil) then
                                _e('Item_'.._actId_s, _iid, 'item_spawn',
                                    'ERR_REF:ItemSpawn:'..tostring(_sid))
                            end
                        end
                    end
                    if _item.merge_score_target and _item.merge_score_target ~= 0 then
                        if _gBpScore == nil or _gBpScore[_item.merge_score_target] == nil then
                            _e('Item_'.._actId_s, _iid, 'merge_score_target',
                                'ERR_REF:BpMergeScore:'..tostring(_item.merge_score_target))
                        end
                    end
                    for _, _btf in ipairs(_bombRefFields) do
                        local _btv = _item[_btf]
                        if _btv and _btv ~= 0 and (_gItem == nil or _gItem[_btv] == nil) then
                            _e('Item_'.._actId_s, _iid, _btf,
                                'ERR_REF:Item:'..tostring(_btv))
                        end
                    end
                    ::_item_next::
                end
                _tl('  [OK] Item_'.._actId_s..': '.._itemCnt..' rows')
            end

            -- spwan_id points to Item directly (not ItemSpawn — common mistake)
            if _mergeTbl then
                local _mc = 0
                for _mid, _merge in pairs(_mergeTbl) do
                    if type(_merge) ~= 'table' then goto _merge_next end
                    _mc = _mc + 1
                    if _merge.spwan_id and _merge.spwan_id ~= 0 then
                        if _gItem == nil or _gItem[_merge.spwan_id] == nil then
                            _e('ItemMerge_'.._actId_s, _mid, 'spwan_id',
                                'ERR_REF:Item:'..tostring(_merge.spwan_id))
                        end
                    end
                    ::_merge_next::
                end
                _tl('  [OK] ItemMerge_'.._actId_s..': '.._mc..' rows')
            end

            if _spawnTbl then
                local _sc = 0
                for _sid, _spawn in pairs(_spawnTbl) do
                    if type(_spawn) ~= 'table' then goto _spawn_next end
                    _sc = _sc + 1
                    if _spawn.spawn_id and _spawn.spawn_id ~= 0 then
                        if _gItem == nil or _gItem[_spawn.spawn_id] == nil then
                            _e('ItemSpawn_'.._actId_s, _sid, 'spawn_id',
                                'ERR_REF:Item:'..tostring(_spawn.spawn_id))
                        end
                    end
                    ::_spawn_next::
                end
                _tl('  [OK] ItemSpawn_'.._actId_s..': '.._sc..' rows')
            end

            if _buildTbl then
                local _bc = 0
                for _bid, _build in pairs(_buildTbl) do
                    if type(_build) ~= 'table' then goto _build_next end
                    _bc = _bc + 1
                    if _build.spawn_id and _build.spawn_id ~= 0 then
                        if _gItem == nil or _gItem[_build.spawn_id] == nil then
                            _e('ItemBuild_'.._actId_s, _bid, 'spawn_id',
                                'ERR_REF:Item:'..tostring(_build.spawn_id))
                        end
                    end
                    ::_build_next::
                end
                _tl('  [OK] ItemBuild_'.._actId_s..': '.._bc..' rows')
            end

            if _adsorbTbl then
                local _ac = 0
                for _aid, _adsorb in pairs(_adsorbTbl) do
                    if type(_adsorb) ~= 'table' then goto _adsorb_next end
                    _ac = _ac + 1
                    if type(_adsorb.adsorb) == 'table' then
                        for _, _tgt in ipairs(_adsorb.adsorb) do
                            if _tgt ~= 0 and (_gItem == nil or _gItem[_tgt] == nil) then
                                _e('ItemAdsorb_'.._actId_s, _aid, 'adsorb',
                                    'ERR_REF:Item:'..tostring(_tgt))
                            end
                        end
                    end
                    ::_adsorb_next::
                end
                _tl('  [OK] ItemAdsorb_'.._actId_s..': '.._ac..' rows')
            end

            local _exploitGrpTbl = Tables['ExploitationGroup_'.._actId_s]
            if _exploitGrpTbl then
                local _egc = 0
                for _egid, _eg in pairs(_exploitGrpTbl) do
                    if type(_eg) ~= 'table' then goto _eg_next end
                    _egc = _egc + 1
                    if type(_eg.ExploitationDetail_id) == 'table' then
                        for _, _edid in ipairs(_eg.ExploitationDetail_id) do
                            if _edid ~= 0 and (_gExploitDetail == nil or _gExploitDetail[_edid] == nil) then
                                _e('ExploitationGroup_'.._actId_s, _egid, 'ExploitationDetail_id',
                                    'ERR_REF:ExploitationDetail:'..tostring(_edid))
                            end
                        end
                    end
                    if type(_eg.may_drop) == 'table' then
                        for _, _dropItemId in ipairs(_eg.may_drop) do
                            if _dropItemId ~= 0 and (_gItem == nil or _gItem[_dropItemId] == nil) then
                                _e('ExploitationGroup_'.._actId_s, _egid, 'may_drop',
                                    'ERR_REF:Item:'..tostring(_dropItemId))
                            end
                        end
                    end
                    ::_eg_next::
                end
                _tl('  [OK] ExploitationGroup_'.._actId_s..': '.._egc..' rows')
            end

            if _exploitDetailTbl then
                local _edc = 0
                for _edid, _ed in pairs(_exploitDetailTbl) do
                    if type(_ed) ~= 'table' then goto _ed_next end
                    _edc = _edc + 1
                    if _ed.dorp_id and _ed.dorp_id ~= 0 then
                        if _gDrop == nil or _gDrop[_ed.dorp_id] == nil then
                            _e('ExploitationDetail_'.._actId_s, _edid, 'dorp_id',
                                'ERR_REF:Drop:'..tostring(_ed.dorp_id))
                        end
                    end
                    ::_ed_next::
                end
                _tl('  [OK] ExploitationDetail_'.._actId_s..': '.._edc..' rows')
            end

            -- LandmarkBuilding rows belong to an activity by ID prefix convention (no explicit link field).
            -- bundleName is the parent dir; prefabs live one level deeper — _checkPrefab recurses one level.
            if Tables.LandmarkBuilding then
                local _lbc = 0
                for _lbid, _lb in pairs(Tables.LandmarkBuilding) do
                    if type(_lb) ~= 'table' then goto _lb_next end
                    if string.sub(tostring(_lbid), 1, #_actId_s) ~= _actId_s then goto _lb_next end
                    _lbc = _lbc + 1
                    if _lb.bundleName and _lb.bundleName ~= '' then
                        _checkPrefab('LandmarkBuilding', _lbid, _lb.bundleName, _lb.assetName)
                    end
                    ::_lb_next::
                end
                _tl('  [OK] LandmarkBuilding (actId='.._actId_s..'): '.._lbc..' rows checked')
            end
        end
    end
end

for _, _id in ipairs(_ids) do
    local _clientData = Tables.ActivityClientData and Tables.ActivityClientData[_id]
    if not _clientData then
        _e('ActivityClientData', _id, 'id', 'ERR_ID_NOT_FOUND')
    else
        _tl('[ID='..tostring(_id)..'] type='..tostring(_clientData.type)..' activityID='..tostring(_clientData.activityID or _id))
        _checkClientData(_id, _clientData)
    end

    -- ② 层级组链（可选，存在则额外验证）
    local _groupData = Tables.ActivityClientHierarchyGroupData
                       and Tables.ActivityClientHierarchyGroupData[_id]
    if _groupData then
        if type(_groupData.hierarchyActivityIDs) ~= 'table' then
            _e('ActivityClientHierarchyGroupData', _id, 'hierarchyActivityIDs', 'ERR_REQUIRED')
        else
            for _, _hid in ipairs(_groupData.hierarchyActivityIDs) do
                local _hData = Tables.ActivityClientHierarchyData
                               and Tables.ActivityClientHierarchyData[_hid]
                if not _hData then
                    _e('ActivityClientHierarchyGroupData', _id, 'hierarchyActivityIDs',
                        'ERR_REF:Tables.ActivityClientHierarchyData:'..tostring(_hid))
                else
                    local _actId = _hData.activityIds
                    if not _actId then
                        _e('ActivityClientHierarchyData', _hid, 'activityIds', 'ERR_REQUIRED')
                    else
                        local _act = Tables.ActivityClientData and Tables.ActivityClientData[_actId]
                        if not _act then
                            _e('ActivityClientHierarchyData', _hid, 'activityIds',
                                'ERR_REF:Tables.ActivityClientData:'..tostring(_actId))
                        else
                            _checkClientData(_actId, _act)
                        end
                    end
                    if _hData.openConditions ~= nil and type(_hData.openConditions) ~= 'table' then
                        _e('ActivityClientHierarchyData', _hid, 'openConditions',
                           'ERR_TYPE:table:'..type(_hData.openConditions))
                    end
                    if _hData.closeConditions ~= nil and type(_hData.closeConditions) ~= 'table' then
                        _e('ActivityClientHierarchyData', _hid, 'closeConditions',
                           'ERR_TYPE:table:'..type(_hData.closeConditions))
                    end
                end
            end
        end
    end
end
");
        sb.AppendLine(BuildMultiSubTableCheckScript(typeMultiSubTableRules));

        return sb.ToString();
    }

    // 生成各 type 的深度字段校验函数集合（返回 Lua 代码片段）
    private static string BuildSubTableCheckFunctions(Dictionary<string, SubTableRule> typeSubTableRules)
    {
        if (typeSubTableRules.Count == 0) return "";
        var sb = new StringBuilder();
        sb.AppendLine("local _subCheckFuncs = {}");
        foreach (var (typeKey, rule) in typeSubTableRules)
        {
            if (typeKey == "2") continue; // LTE 链专属
            if (rule.Fields == null || rule.Fields.Count == 0) continue;
            var tableName = rule.Table.Contains('.') ? rule.Table.Split('.')[1] : rule.Table;
            sb.AppendLine($"_subCheckFuncs[{typeKey}] = function(actId, subId, row)");
            foreach (var field in rule.Fields)
            {
                var fa = $"row.{field.Name}";
                if (field.Required)
                {
                    sb.AppendLine($"    if {fa} == nil or tostring({fa}) == '' then");
                    sb.AppendLine($"        _e('{tableName}', subId, '{field.Name}', 'ERR_REQUIRED')");
                    sb.AppendLine($"    else");
                }
                else
                {
                    sb.AppendLine($"    if {fa} ~= nil and tostring({fa}) ~= '' then");
                }
                if (field.Type is "number" or "string" or "boolean" or "table")
                {
                    sb.AppendLine($"        if type({fa}) ~= '{field.Type}' then");
                    sb.AppendLine($"            _e('{tableName}', subId, '{field.Name}',");
                    sb.AppendLine($"                'ERR_TYPE:{field.Type}:'..type({fa}))");
                    sb.AppendLine($"        end");
                }
                if (!string.IsNullOrEmpty(field.RefTable))
                {
                    sb.AppendLine($"        local _rv = tonumber({fa}) or {fa}");
                    // 0 是配置表通用"未设置"哨兵值，跳过引用检查
                    sb.AppendLine($"        if _rv ~= 0 and ({field.RefTable} == nil or {field.RefTable}[_rv] == nil) then");
                    sb.AppendLine($"            _e('{tableName}', subId, '{field.Name}',");
                    sb.AppendLine($"                'ERR_REF:{EscapeLua(field.RefTable)}:'..tostring({fa}))");
                    sb.AppendLine($"        end");
                }
                if (!string.IsNullOrEmpty(field.CustomCheck))
                {
                    sb.AppendLine($"        do");
                    sb.AppendLine($"            local id = subId");
                    sb.AppendLine($"            local function err(f,m) _e('{tableName}', subId, f, m) end");
                    sb.AppendLine($"            {field.CustomCheck}");
                    sb.AppendLine($"        end");
                }
                sb.AppendLine($"    end");
            }
            sb.AppendLine("end");
        }
        return sb.ToString();
    }

    // 生成多子表存在性验证脚本（typeMultiSubTableRules）
    private static string BuildMultiSubTableCheckScript(
        Dictionary<string, List<MultiSubTableEntry>> rules)
    {
        if (rules.Count == 0) return "";
        var sb = new StringBuilder();
        sb.AppendLine("for _, _mid in ipairs(_ids) do  -- multi-sub-table existence check");
        sb.AppendLine("    local _cd = Tables.ActivityClientData and Tables.ActivityClientData[_mid]");
        sb.AppendLine("    if _cd then");
        sb.AppendLine("        local _t = tostring(_cd.type)");
        foreach (var (typeKey, entries) in rules)
        {
            sb.AppendLine($"        if _t == '{typeKey}' then");
            foreach (var entry in entries)
            {
                var tblName = entry.Table.Contains('.') ? entry.Table.Split('.')[1] : entry.Table;
                var lf = entry.LookupField ?? "activityID";
                sb.AppendLine($"            local _lk = _cd.{lf} or _mid");
                sb.AppendLine($"            if {entry.Table} == nil or {entry.Table}[_lk] == nil then");
                sb.AppendLine($"                _e('ActivityClientData', _mid, '{lf}',");
                sb.AppendLine($"                    'ERR_REF:{EscapeLua(entry.Table)}:'..tostring(_lk))");
                sb.AppendLine($"            end");
            }
            sb.AppendLine($"        end");
        }
        sb.AppendLine("    end");
        sb.AppendLine("end");
        return sb.ToString();
    }

    // ─── 动态生成 Lua 校验脚本（完全由 JSON 规则驱动）──────────────────────────

    private static string ToLocalVar(string luaKey)
    {
        const string prefix = "Tables.";
        var name = luaKey.StartsWith(prefix, StringComparison.Ordinal)
            ? luaKey[prefix.Length..]
            : luaKey.Replace(".", "_");
        return "_t_" + name;
    }

    private static string BuildValidationScript(RulesRoot rules, HashSet<string> filterIds)
    {
        var sb = new StringBuilder();
        sb.AppendLine("_validation_errors = {}");
        sb.AppendLine();
        sb.AppendLine("local function _err(tbl, id, field, msg)");
        sb.AppendLine("    table.insert(_validation_errors, tbl..\"|\"..tostring(id)..\"|\"..field..\"|\"..msg)");
        sb.AppendLine("end");
        sb.AppendLine();

        // ID 过滤函数
        if (filterIds != null && filterIds.Count > 0)
        {
            sb.Append("local _filter = {");
            foreach (var id in filterIds)
                sb.Append($"[\"{id}\"]=true,");
            sb.AppendLine("}");
            sb.AppendLine("local function _shouldCheck(id) return _filter[tostring(id)] == true end");
        }
        else
        {
            sb.AppendLine("local function _shouldCheck(id) return true end");
        }
        sb.AppendLine();

        // ── 关键修正：把所有表的局部变量提升到脚本顶层 ──
        // 这样跨表引用时，被引用表的变量在任意校验块中都可见
        sb.AppendLine("-- 所有配置表的顶层局部引用（必须在校验块之外声明，以保证跨块可见）");
        foreach (var t in rules.Tables)
        {
            var lv = ToLocalVar(t.LuaKey);
            sb.AppendLine($"local {lv} = {t.LuaKey}");
            // 若表未加载（nil）则输出警告，而不是把它当空表跳过
            sb.AppendLine($"if {lv} == nil then");
            sb.AppendLine($"    table.insert(_validation_errors, \"{t.Name}||_load|ERR_TABLE_NOT_LOADED\")");
            sb.AppendLine($"end");
        }
        sb.AppendLine();

        foreach (var table in rules.Tables)
        {
            var localVar = ToLocalVar(table.LuaKey);

            sb.AppendLine($"-- ════ 校验：{table.Name} ({table.Desc}) ════");
            sb.AppendLine($"if type({localVar}) == \"table\" then");
            sb.AppendLine($"    for _id, _row in pairs({localVar}) do");
            sb.AppendLine($"        if not _shouldCheck(_id) then goto _next_{localVar} end");
            sb.AppendLine();

            foreach (var field in table.Fields)
            {
                var isKeyField = field.Name == table.KeyField;

                // ── required 检查（主键字段用 _id 而非 _row.xxx，防止默认值压缩导致字段缺失）──
                var fieldAccess = isKeyField ? "_id" : $"_row.{field.Name}";

                if (field.Required)
                {
                    sb.AppendLine($"        -- [{field.Name}] required: {field.Desc}");
                    sb.AppendLine($"        if {fieldAccess} == nil or tostring({fieldAccess}) == \"\" then");
                    sb.AppendLine($"            _err(\"{table.Name}\", _id, \"{field.Name}\", \"ERR_REQUIRED\")");
                    sb.AppendLine($"        else");
                }
                else
                {
                    sb.AppendLine($"        -- [{field.Name}]: {field.Desc}");
                    sb.AppendLine($"        if {fieldAccess} ~= nil and tostring({fieldAccess}) ~= \"\" then");
                }

                // ── 类型检查 ──
                if (field.Type is "number" or "string" or "boolean" or "table")
                {
                    sb.AppendLine($"            if type({fieldAccess}) ~= \"{field.Type}\" then");
                    sb.AppendLine($"                _err(\"{table.Name}\", _id, \"{field.Name}\",");
                    sb.AppendLine($"                    \"ERR_TYPE:{field.Type}:\"..type({fieldAccess}))");
                    sb.AppendLine($"            end");
                }

                // ── 跨表引用检查 ──
                if (!string.IsNullOrEmpty(field.RefTable))
                {
                    var refLocalVar = ToLocalVar(field.RefTable);
                    if (field.RefIsArray)
                    {
                        sb.AppendLine($"            if type({fieldAccess}) == \"table\" then");
                        sb.AppendLine($"                for _, _refId in ipairs({fieldAccess}) do");
                        sb.AppendLine($"                    local _key = tonumber(_refId) or _refId");
                        sb.AppendLine($"                    if {refLocalVar} == nil or {refLocalVar}[_key] == nil then");
                        sb.AppendLine($"                        _err(\"{table.Name}\", _id, \"{field.Name}\",");
                        sb.AppendLine($"                            \"ERR_REF:{EscapeLua(field.RefTable)}:\"..tostring(_refId))");
                        sb.AppendLine($"                    end");
                        sb.AppendLine($"                end");
                        sb.AppendLine($"            end");
                    }
                    else
                    {
                        sb.AppendLine($"            local _refKey = tonumber({fieldAccess}) or {fieldAccess}");
                        // 0 是配置表通用"未设置"哨兵值，跳过引用检查，避免假阳性
                        sb.AppendLine($"            if _refKey ~= 0 and ({refLocalVar} == nil or {refLocalVar}[_refKey] == nil) then");
                        sb.AppendLine($"                _err(\"{table.Name}\", _id, \"{field.Name}\",");
                        sb.AppendLine($"                    \"ERR_REF:{EscapeLua(field.RefTable)}:\"..tostring({fieldAccess}))");
                        sb.AppendLine($"            end");
                    }
                }

                // ── 自定义检查 ──
                if (!string.IsNullOrEmpty(field.CustomCheck))
                {
                    sb.AppendLine($"            do");
                    sb.AppendLine($"                local row, id, tbl_name = _row, _id, \"{table.Name}\"");
                    sb.AppendLine($"                local function err(f, m) _err(tbl_name, id, f, m) end");
                    sb.AppendLine($"                {field.CustomCheck}");
                    sb.AppendLine($"            end");
                }

                sb.AppendLine($"        end");
                sb.AppendLine();
            }

            sb.AppendLine($"        ::_next_{localVar}::");
            sb.AppendLine($"    end");
            sb.AppendLine($"end");
            sb.AppendLine();
        }

        return sb.ToString();
    }

    // ─── 全局桩 ───────────────────────────────────────────────────────────────────

    private static string BuildGlobalStub(RulesRoot rules)
    {
        var sb = new StringBuilder();

        // ── 配置表容器 ──
        sb.AppendLine("Tables = {}");
        // SetDataTableMetatable(targetTable, dataTable, metatable, name, flag)
        // 真实实现：把 dataTable 里的每一行复制进 targetTable，并为每行设置元表（默认值）。
        // 桩必须忠实模拟：否则 Tables.ActivityClientData[id] 永远为 nil，导致所有 ID 报"不存在"。
        sb.AppendLine(@"
Tables.SetDataTableMetatable = function(target, data, mt, name, flag)
    for k, v in pairs(data) do
        if type(v) == 'table' then
            setmetatable(v, mt)
        end
        target[k] = v
    end
end
Tables.SetSubTableMetatable = Tables.SetDataTableMetatable
Tables.CheckNullValue = function() end
");

        // ── 引擎/框架基础桩 ──
        sb.AppendLine("IsEditor = false");
        sb.AppendLine("Debug = { Log=function()end, LogError=function()end, LogWarning=function()end, LogFormat=function()end }");
        sb.AppendLine("print = function() end");

        // ── Lua OOP 框架桩（class / handler / import 等常见 pattern）──
        sb.AppendLine(@"
-- 极简 class() 实现，支持 class('Name') 和 class('Name', Base)
function class(name, base)
    local cls = {}
    cls.__name = name
    cls.__index = cls
    if base then setmetatable(cls, { __index = base }) end
    cls.new  = function(...) local o = setmetatable({}, cls); if o.ctor then o:ctor(...) end; return o end
    cls.super = base or {}
    return cls
end
function handler(obj, fn) return function(...) return fn(obj, ...) end end
function import(m) return {} end
function require(m) return _G[m] or {} end
function ipairs_safe(t) if type(t)~='table' then return function()end end return ipairs(t) end
function pairs_safe(t)  if type(t)~='table' then return function()end end return pairs(t) end
function CheckMultiCommonConditions(conds) return true end
function CheckCommonCondition(cond) return true end
function isset(v) return v ~= nil end
-- 常见全局对象桩
ArchiveManager   = { IsExistDataTable=function()return false end, LoadDataTable=function()end }
ArchiveRoot      = setmetatable({}, { __index=function() return {} end })
EventManager     = { AddEventListener=function()end, RemoveEventListener=function()end, DispatchEvent=function()end }
TimerManager     = { AddTimer=function()return 0 end, RemoveTimer=function()end }
CS               = setmetatable({}, { __index=function(t,k) t[k]=setmetatable({},{__index=function(t2,k2) t2[k2]=function()end; return t2[k2] end}); return t[k] end })
UnityEngine      = CS.UnityEngine or setmetatable({}, { __index=function(t,k) t[k]=function()end; return t[k] end })
SolarRoot        = setmetatable({}, { __index=function(t,k) t[k]=setmetatable({},{__index=function()return '' end}); return t[k] end })
");

        // ── 用户自定义引擎 API 桩（来自 JSON globalStubExtras）──
        if (!string.IsNullOrWhiteSpace(rules.GlobalStubExtras))
            sb.AppendLine(rules.GlobalStubExtras);

        return sb.ToString();
    }

    // ─── 工具方法 ─────────────────────────────────────────────────────────────────

    private static string FindLuaOutputDir(string workbookPath)
    {
        var dir = Path.GetDirectoryName(workbookPath);
        for (var depth = 0; depth <= 6; depth++)
        {
            if (dir == null) break;
            var candidate = Path.Combine(dir, "Code", "Assets", "LuaScripts", "Tables");
            if (Directory.Exists(candidate))
                return candidate;
            dir = Path.GetDirectoryName(dir);
        }
        return null;
    }

    private static TableLineMaps BuildLineMap(string luaText)
    {
        var byLine = new Dictionary<int, LineMapEntry>();
        var byId   = new Dictionary<string, LineMapEntry>(StringComparer.Ordinal);
        using var reader = new StringReader(luaText);
        int lineNum = 0, dataRowIndex = 0;
        string line;
        while ((line = reader.ReadLine()) != null)
        {
            lineNum++;
            var m = LuaRowKeyRegex.Match(line);
            if (!m.Success) continue;
            // Excel行：1行标题 + 3行字段定义 + 1偏移 + N = 5 + dataRowIndex
            var entry = new LineMapEntry(m.Groups[1].Value, 5 + dataRowIndex);
            byLine[lineNum] = entry;
            byId.TryAdd(entry.Id, entry);
            dataRowIndex++;
        }
        return new TableLineMaps(byLine, byId);
    }

    private static string ResolveLocation(
        string tableName,
        int luaLineNum,
        Dictionary<string, TableLineMaps> lineMaps
    )
    {
        if (luaLineNum < 0 || !lineMaps.TryGetValue(tableName, out var maps))
            return "（位置未知）";
        var best = maps.ByLine.Keys.Where(k => k <= luaLineNum).DefaultIfEmpty(-1).Max();
        if (best < 0) return $"Lua第{luaLineNum}行（表头区域）";
        var e = maps.ByLine[best];
        return $"{tableName}.xlsx 第{e.ExcelDisplayRow}行（ID={e.Id}，Lua第{luaLineNum}行）";
    }

    private static string ResolveLocationById(
        string tableName,
        string activityId,
        Dictionary<string, TableLineMaps> lineMaps
    )
    {
        if (!lineMaps.TryGetValue(tableName, out var maps)) return "（位置未知）";
        return maps.ById.TryGetValue(activityId, out var entry)
            ? $"{tableName}.xlsx 第{entry.ExcelDisplayRow}行（ID={activityId}）"
            : $"{tableName}.xlsx（ID={activityId}，行号未知）";
    }

    private static int ExtractLuaLineNumber(string errorMsg)
    {
        var m = LuaErrorLineRegex.Match(errorMsg);
        return m.Success ? int.Parse(m.Groups[1].Value) : -1;
    }

    private static string CleanLuaError(string msg)
    {
        var idx = msg.IndexOf('\n');
        return idx > 0 ? msg[..idx].Trim() : msg.Trim();
    }

    // Lua 字符串内转义（双引号、反斜杠）
    private static string EscapeLua(string s) =>
        s?.Replace("\\", "\\\\").Replace("\"", "\\\"") ?? "";
}
