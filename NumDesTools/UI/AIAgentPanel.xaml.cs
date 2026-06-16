using System.IO;
using System.Runtime.Versioning;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Windows;
using System.Windows.Input;
using Markdig;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace NumDesTools.UI;

[SupportedOSPlatform("windows")]
public partial class AIAgentPanel
{
    private CancellationTokenSource _cts;
    private readonly List<object> _history = [];
    private const int MaxHistoryMessages = 60;
    private bool _isRunning;
    // 执行中的插话队列：Enter 发送时入队，每步循环开头注入对话
    private readonly System.Collections.Concurrent.ConcurrentQueue<string> _interjectQueue = new();

    private static readonly MarkdownPipeline MdPipeline = new MarkdownPipelineBuilder()
        .UseAdvancedExtensions()
        .Build();

    private static readonly string HtmlTemplate =
        @"<html><head><meta charset='utf-8'><style>
body{background:#1c1c1c;color:#d4d4d4;font-family:'微软雅黑',sans-serif;font-size:10pt;line-height:1.5;margin:0;padding:8px 10px;overflow-y:auto}
.msg{margin:5px 0;max-width:98%}
.msg.user{margin-left:auto;text-align:right}
.msg.assistant{margin-left:0}
.role{font-size:.72em;color:#555;margin-bottom:2px}
.role .ts{color:#3d3d3d}
.content{display:inline-block;padding:6px 10px;border-radius:7px;word-wrap:break-word;text-align:left;max-width:100%}
.user .content{background:#1e3a5f;color:#e8e8e8}
.assistant .content{background:#2a2a30;color:#e8e8e8}
.content p{margin:3px 0}
.content h1,.content h2,.content h3{font-size:1em;font-weight:bold;margin:4px 0 2px}
pre{background:#252526;color:#dcdcdc;padding:7px;border-radius:5px;overflow-x:auto;font-size:10pt;margin:4px 0}
code{font-family:Consolas,monospace;background:#252526;padding:1px 3px;border-radius:3px;font-size:10pt}
table{border-collapse:collapse;font-size:.88em;margin:4px 0;width:auto}
th,td{border:1px solid #555;padding:3px 8px;text-align:left;white-space:nowrap}
th{background:#2a2d2e;color:#c6c6c6;font-weight:bold}
tr:nth-child(even) td{background:#2a2a2a}
ul,ol{margin:3px 0;padding-left:18px}
li{margin:1px 0}
a[href^='excel://']{color:#4ec9b0;text-decoration:none;border-bottom:1px dashed #4ec9b0}
a[href^='excel://']:hover{background:#1a3a35;border-radius:2px}
</style></head><body></body></html>";

    private static readonly object[] ToolDefinitions =
    [
        new
        {
            type = "function",
            function = new
            {
                name = "read_selection",
                description = "读取 Excel 当前选中区域的单元格值，返回二维数组字符串",
                parameters = new
                {
                    type = "object",
                    properties = new { },
                    required = Array.Empty<string>(),
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "write_range",
                description = "向指定单元格地址写入值",
                parameters = new
                {
                    type = "object",
                    properties = new
                    {
                        address = new
                        {
                            type = "string",
                            description = "单元格地址，如 A1 或 B2:B10",
                        },
                        value = new
                        {
                            type = "string",
                            description = "要写入的值，多行用\\n分隔，多列用\\t分隔",
                        },
                    },
                    required = new[] { "address", "value" },
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "run_formula",
                description = "在指定单元格写入公式并返回计算结果",
                parameters = new
                {
                    type = "object",
                    properties = new
                    {
                        address = new { type = "string", description = "目标单元格地址" },
                        formula = new { type = "string", description = "Excel 公式，以 = 开头" },
                    },
                    required = new[] { "address", "formula" },
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "list_udfs",
                description = "列出插件所有可用的 UDF 自定义函数名称",
                parameters = new
                {
                    type = "object",
                    properties = new { },
                    required = Array.Empty<string>(),
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "read_sheet",
                description = "读取指定 Sheet 的数据（前 N 行）",
                parameters = new
                {
                    type = "object",
                    properties = new
                    {
                        sheet_name = new
                        {
                            type = "string",
                            description = "Sheet 名称，留空则读取当前活动 Sheet",
                        },
                        max_rows = new { type = "integer", description = "最多读取行数，默认 50" },
                    },
                    required = Array.Empty<string>(),
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "list_sheets",
                description = "列出当前工作簿所有 Sheet 名称",
                parameters = new
                {
                    type = "object",
                    properties = new { },
                    required = Array.Empty<string>(),
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "get_workbook_structure",
                description = "获取工作簿结构：每个 Sheet 的行列数和前两行内容，用于快速了解数据布局",
                parameters = new
                {
                    type = "object",
                    properties = new { },
                    required = Array.Empty<string>(),
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "batch_write",
                description = "批量向多个单元格写入数据，一次调用写多个地址",
                parameters = new
                {
                    type = "object",
                    properties = new
                    {
                        sheet_name = new
                        {
                            type = "string",
                            description = "Sheet 名称，留空则用当前活动 Sheet",
                        },
                        writes = new
                        {
                            type = "array",
                            description = "写入列表，每项含 address 和 value",
                            items = new
                            {
                                type = "object",
                                properties = new
                                {
                                    address = new { type = "string" },
                                    value = new { type = "string" },
                                },
                            },
                        },
                    },
                    required = new[] { "writes" },
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "run_vba_macro",
                description = "执行一段 VBA 代码完成复杂操作（格式、筛选、复制、跨表等），代码在 Excel 中直接运行",
                parameters = new
                {
                    type = "object",
                    properties = new
                    {
                        code = new
                        {
                            type = "string",
                            description = "完整的 VBA Sub 代码，包含 Sub...End Sub",
                        },
                    },
                    required = new[] { "code" },
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "apply_format",
                description = "对指定单元格区域应用格式：背景色、字体色、粗体、斜体、边框、列宽、行高等。不依赖 VBA，xlsx 文件可用。",
                parameters = new
                {
                    type = "object",
                    properties = new
                    {
                        range = new
                        {
                            type = "string",
                            description = "单元格区域地址，如 'A1'、'B2:D5'",
                        },
                        sheet_name = new
                        {
                            type = "string",
                            description = "Sheet 名，空则用当前活动 Sheet",
                        },
                        bg_color = new
                        {
                            type = "string",
                            description = "背景色，十六进制 RGB，如 'FF0000' 表示红色",
                        },
                        font_color = new { type = "string", description = "字体色，十六进制 RGB" },
                        bold = new { type = "boolean", description = "是否加粗" },
                        italic = new { type = "boolean", description = "是否斜体" },
                        font_size = new { type = "number", description = "字号" },
                        wrap_text = new { type = "boolean", description = "是否自动换行" },
                        h_align = new
                        {
                            type = "string",
                            description = "水平对齐：left / center / right",
                        },
                        col_width = new
                        {
                            type = "number",
                            description = "列宽（仅对单列或区域第一列有效）",
                        },
                        row_height = new
                        {
                            type = "number",
                            description = "行高（仅对单行或区域第一行有效）",
                        },
                    },
                    required = new[] { "range" },
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "check_cross_ref",
                description = "跨 Sheet 检查外键合法性：验证 source_sheet 的 source_col 列中每个值是否都存在于 target_sheet 的 target_col 列中，返回缺失值及其行号",
                parameters = new
                {
                    type = "object",
                    properties = new
                    {
                        source_sheet = new { type = "string", description = "要检查的 Sheet 名称" },
                        source_col = new
                        {
                            type = "string",
                            description = "要检查的列名或列号（如 'activityID' 或 'B'），从第2行开始（第1行为列名）",
                        },
                        target_sheet = new
                        {
                            type = "string",
                            description = "合法值所在的 Sheet 名称",
                        },
                        target_col = new
                        {
                            type = "string",
                            description = "合法值所在的列名或列号",
                        },
                    },
                    required = new[] { "source_sheet", "source_col", "target_sheet", "target_col" },
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "read_lua_table",
                description = "读取 C:\\M1Work\\Code\\Assets\\LuaScripts\\Tables\\ 目录下的 Lua 导出数据文件，返回结构化内容供分析对比",
                parameters = new
                {
                    type = "object",
                    properties = new
                    {
                        table_name = new
                        {
                            type = "string",
                            description = "Lua 表文件名（不含扩展名），如 LteData、ActivityBpData",
                        },
                        max_rows = new { type = "integer", description = "最多返回行数，默认 100" },
                    },
                    required = new[] { "table_name" },
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "list_lua_tables",
                description = "列出 C:\\M1Work\\Code\\Assets\\LuaScripts\\Tables\\ 目录下所有可用的 Lua 导出表名，用于确认哪些表可以被 read_lua_table 读取",
                parameters = new
                {
                    type = "object",
                    properties = new { },
                    required = Array.Empty<string>(),
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "describe_data",
                description = "统计指定范围的数据概况：行列数、空值率、类型分布、数值范围/均值/标准差",
                parameters = new
                {
                    type = "object",
                    properties = new
                    {
                        sheet_name = new
                        {
                            type = "string",
                            description = "Sheet 名，留空用当前 Sheet",
                        },
                        range = new
                        {
                            type = "string",
                            description = "单元格范围如 A1:D100，留空用当前选区",
                        },
                    },
                    required = Array.Empty<string>(),
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "detect_patterns",
                description = "对指定列检测：异常值（3σ）、趋势（递增/递减）、重复值",
                parameters = new
                {
                    type = "object",
                    properties = new
                    {
                        sheet_name = new
                        {
                            type = "string",
                            description = "Sheet 名，留空用当前 Sheet",
                        },
                        col_range = new { type = "string", description = "列范围如 B2:B100" },
                    },
                    required = new[] { "col_range" },
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "sim_progression",
                description = "模拟数值增长曲线：给定初始值和增长方式，生成 N 步数据，可直接写入表格",
                parameters = new
                {
                    type = "object",
                    properties = new
                    {
                        init_val = new { type = "number", description = "初始值" },
                        growth_rate = new
                        {
                            type = "number",
                            description = "增长率（线性为每步增量，倍率为倍数，幂次为指数）",
                        },
                        growth_type = new
                        {
                            type = "string",
                            description = "增长类型：linear/multiply/power，默认 multiply",
                        },
                        steps = new { type = "integer", description = "步数，默认 10" },
                        write_sheet = new
                        {
                            type = "string",
                            description = "写入 Sheet 名，留空只输出不写入",
                        },
                        write_start_cell = new
                        {
                            type = "string",
                            description = "写入起始单元格如 A2",
                        },
                    },
                    required = new[] { "init_val", "growth_rate" },
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "calc_drop_expectation",
                description = "分析掉落表：计算每个物品的期望产出、标准差、至少掉落1次的概率",
                parameters = new
                {
                    type = "object",
                    properties = new
                    {
                        sheet_name = new
                        {
                            type = "string",
                            description = "Sheet 名，留空用当前 Sheet",
                        },
                        item_col = new { type = "string", description = "物品名列名或列字母" },
                        prob_col = new { type = "string", description = "概率列名或列字母" },
                        trials = new { type = "integer", description = "模拟抽取次数，默认 100" },
                    },
                    required = new[] { "item_col", "prob_col" },
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "balance_check",
                description = "检查数值列的相邻增长比是否在合理范围内，用于验证关卡/升级曲线是否平衡",
                parameters = new
                {
                    type = "object",
                    properties = new
                    {
                        sheet_name = new
                        {
                            type = "string",
                            description = "Sheet 名，留空用当前 Sheet",
                        },
                        col_range = new { type = "string", description = "列范围如 C2:C50" },
                        min_ratio = new { type = "number", description = "最小增长比，默认 1.0" },
                        max_ratio = new { type = "number", description = "最大增长比，默认 2.0" },
                    },
                    required = new[] { "col_range" },
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "cost_curve_fit",
                description = "对升级消耗/关卡数值列拟合增长曲线，输出线性和指数拟合公式及 R² 值",
                parameters = new
                {
                    type = "object",
                    properties = new
                    {
                        sheet_name = new
                        {
                            type = "string",
                            description = "Sheet 名，留空用当前 Sheet",
                        },
                        col_range = new { type = "string", description = "列范围如 D2:D30" },
                    },
                    required = new[] { "col_range" },
                },
            },
        },

        // ── P0 新工具定义 ─────────────────────────────────────────────────────
        new { type="function", function=new { name="get_active_cell",
            description="获取当前活动单元格的地址、值及所在 Sheet 名",
            parameters=new { type="object", properties=new { }, required=Array.Empty<string>() } } },

        new { type="function", function=new { name="navigate_to",
            description="跳转到指定 Sheet 的指定单元格并激活",
            parameters=new { type="object", properties=new {
                sheet_name=new{type="string",description="Sheet 名"},
                address=new{type="string",description="单元格地址如 A1"} },
                required=new[]{"address"} } } },

        new { type="function", function=new { name="manage_sheet",
            description="新建/删除/重命名/复制/移动 Sheet",
            parameters=new { type="object", properties=new {
                action=new{type="string",description="create|delete|rename|copy|move"},
                sheet_name=new{type="string",description="操作的 Sheet 名"},
                new_name=new{type="string",description="新名称（rename 时用）"},
                target_position=new{type="integer",description="目标位置 1-based（move/copy 时）"},
                copy_before=new{type="string",description="复制到哪个 Sheet 前面"} },
                required=new[]{"action","sheet_name"} } } },

        new { type="function", function=new { name="save_workbook",
            description="保存当前工作簿，或另存为指定路径",
            parameters=new { type="object", properties=new {
                save_as_path=new{type="string",description="另存路径，不传则原路径保存"},
                file_format=new{type="string",description="xlsx|xlsm|csv"} },
                required=Array.Empty<string>() } } },

        new { type="function", function=new { name="find_replace",
            description="在指定区域执行查找替换，返回完成提示",
            parameters=new { type="object", properties=new {
                find=new{type="string",description="查找内容"},
                replace=new{type="string",description="替换内容"},
                sheet_name=new{type="string",description="Sheet 名，不传用当前"},
                address=new{type="string",description="范围，不传用整 Sheet"},
                match_case=new{type="boolean",description="是否区分大小写"},
                match_entire_cell=new{type="boolean",description="是否整单元格匹配"} },
                required=new[]{"find","replace"} } } },

        new { type="function", function=new { name="insert_delete_rows_cols",
            description="插入或删除整行/整列",
            parameters=new { type="object", properties=new {
                action=new{type="string",description="insert_rows|delete_rows|insert_cols|delete_cols"},
                sheet_name=new{type="string",description="Sheet 名"},
                start=new{type="integer",description="起始行/列号（1-based）"},
                count=new{type="integer",description="数量，默认 1"} },
                required=new[]{"action","sheet_name","start"} } } },

        new { type="function", function=new { name="sort_range",
            description="对指定区域按列排序",
            parameters=new { type="object", properties=new {
                sheet_name=new{type="string",description="Sheet 名"},
                address=new{type="string",description="排序区域如 A1:D100"},
                key_column=new{type="integer",description="排序键相对列号（1-based）"},
                ascending=new{type="boolean",description="true 升序 false 降序"},
                has_header=new{type="boolean",description="是否有标题行"} },
                required=new[]{"sheet_name","address","key_column"} } } },

        new { type="function", function=new { name="set_number_format",
            description="设置单元格数字格式，如日期、百分比、货币、小数位数",
            parameters=new { type="object", properties=new {
                sheet_name=new{type="string",description="Sheet 名"},
                address=new{type="string",description="单元格/区域地址"},
                format_string=new{type="string",description="格式字符串如 0.00% / yyyy-mm-dd / #,##0.00"} },
                required=new[]{"sheet_name","address","format_string"} } } },

        new { type="function", function=new { name="download_image",
            description="从 URL 下载图片到本地临时文件，返回本地路径供 insert_image 使用",
            parameters=new { type="object", properties=new {
                url=new{type="string",description="图片 URL"} },
                required=new[]{"url"} } } },

        new { type="function", function=new { name="generate_image",
            description="根据文字描述用 AI 生成一张图片，保存到本地临时文件，返回文件路径。调用后必须立即用 insert_image 将该路径插入 Excel 单元格。",
            parameters=new { type="object", properties=new {
                prompt=new{type="string",description="图片内容描述，支持中英文，越详细越好，例如：一只在草地上奔跑的橙色猫咪，卡通风格"},
                model=new{type="string",description="生图模型，可选：gemini-3.1-flash-image-preview（快速，默认）、gemini-3-pro-image-preview（高质量）"} },
                required=new[]{"prompt"} } } },

        new { type="function", function=new { name="insert_image",
            description="将本地图片文件插入 Excel 指定单元格位置，图片左上角对齐单元格左上角。通常与 generate_image 或 download_image 配合使用。",
            parameters=new { type="object", properties=new {
                sheet_name=new{type="string",description="工作表名称，留空则用当前活动表"},
                file_path=new{type="string",description="本地图片完整路径，通常来自 generate_image 返回值，如 C:\\Users\\cent\\AppData\\Local\\Temp\\ndtools_gen_xxx.png"},
                anchor_address=new{type="string",description="插入位置的单元格地址，例如 B3"},
                width_pt=new{type="number",description="图片宽度（磅），不填则保持原始尺寸"},
                height_pt=new{type="number",description="图片高度（磅），不填则保持原始尺寸"} },
                required=new[]{"file_path","anchor_address"} } } },
    ];

    // 匹配 Sheet1!A1:B5 / A1:B5 / A1 形式的单元格地址
    private static readonly System.Text.RegularExpressions.Regex CellAddressRegex = new(
        @"(?<![""'#>\/])(?:([A-Za-z0-9_一-龥]+)!)?([A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?)\b",
        System.Text.RegularExpressions.RegexOptions.None
    );

    public AIAgentPanel()
    {
        InitializeComponent();
        PopulateModelList();
        ChatOutput.NavigateToString(HtmlTemplate);
        ChatOutput.Navigating += ChatOutput_Navigating;
        LoadAgentHistory();
        // CTP 内 WebBrowser 的键盘事件会被 Excel 截走：
        // cancelBubble 阻止事件冒泡到 Excel，并主动调用 execCommand 保证 Ctrl+C/A/X 可用
        ChatOutput.LoadCompleted += (_, _) =>
        {
            try
            {
                dynamic doc = ChatOutput.Document;
                // 点击时确保文档获得焦点，Ctrl+C 才能稳定工作
                doc.attachEvent("onmousedown", new Action<dynamic>(_ => doc.body.focus()));
                doc.attachEvent(
                    "onkeydown",
                    new Action<dynamic>(e =>
                    {
                        if ((int)e.ctrlKey != 1) return;
                        e.cancelBubble = true;
                        e.returnValue = false;
                        int keyCode = (int)e.keyCode;
                        if (keyCode == 67) doc.execCommand("Copy",     false, null); // Ctrl+C
                        if (keyCode == 65) doc.execCommand("SelectAll", false, null); // Ctrl+A
                        if (keyCode == 88) doc.execCommand("Cut",      false, null); // Ctrl+X
                    })
                );
            }
            catch { }
        };
        // TextBox 也需要拦截，防止 Excel 截走 Ctrl+C/A/X/Z/Y
        AttachClipboardKeys(TaskInput);
        AttachClipboardKeys(CustomInstructionInput);
        // 深色主题：输入框默认背景
        var darkBg = new System.Windows.Media.SolidColorBrush(
            System.Windows.Media.Color.FromRgb(0x1e, 0x1e, 0x1e));
        TaskInput.Background = darkBg;
        CustomInstructionInput.Background = darkBg;
        TaskInput.CaretBrush = new System.Windows.Media.SolidColorBrush(
            System.Windows.Media.Colors.White);
        CustomInstructionInput.Text = AppServices.Config.Agent.CustomInstruction;
        CustomInstructionInput.LostFocus += (_, _) =>
        {
            AppServices.Config.Agent.CustomInstruction = CustomInstructionInput.Text;
            AppServices.Config.Save("AgentCustomInstruction", CustomInstructionInput.Text);
        };
    }

    private static void ChatOutput_Navigating(
        object sender,
        System.Windows.Navigation.NavigatingCancelEventArgs e
    )
    {
        var uri = e.Uri?.ToString() ?? "";
        if (!uri.StartsWith("excel://cell/"))
            return;
        e.Cancel = true;
        var address = Uri.UnescapeDataString(uri["excel://cell/".Length..]);
        try
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                dynamic app = AppServices.App;
                if (address.Contains('!'))
                {
                    var parts = address.Split('!');
                    dynamic ws = app.ActiveWorkbook.Sheets[parts[0]];
                    ws.Activate();
                    ws.Range[parts[1]].Select();
                }
                else
                {
                    app.ActiveSheet.Range[address].Select();
                }
            });
        }
        catch { }
    }

    private const string AgentModelKey = "AgentModel";
    private const int AgentHistoryPageSize = 40;

    private void PopulateModelList()
    {
        ModelComboBox.Items.Clear();
        var models = AppServices.Config.Llm.ModelList;
        if (models.Count == 0)
            models = [AppServices.Config.Llm.Model];
        foreach (var m in models)
            ModelComboBox.Items.Add(m);

        // 优先恢复上次 Agent 选用的模型，fallback 到全局模型
        AppServices.GlobalValue.Value.TryGetValue(AgentModelKey, out var savedAgentModel);
        var target = !string.IsNullOrEmpty(savedAgentModel) && ModelComboBox.Items.Contains(savedAgentModel)
            ? savedAgentModel
            : AppServices.Config.Llm.Model;
        ModelComboBox.SelectedItem = ModelComboBox.Items.Contains(target) ? target : ModelComboBox.Items[0];
    }

    private void ModelComboBox_SelectionChanged(
        object sender,
        System.Windows.Controls.SelectionChangedEventArgs e
    )
    {
        if (ModelComboBox.SelectedItem is string model)
            AppServices.GlobalValue.SaveValue(AgentModelKey, model);
    }

    private void LoadAgentHistory()
    {
        var db = new ChatHistoryManager();
        var history = db.LoadChatHistory(AgentHistoryPageSize, isAgent: true);
        if (history.Count == 0)
            return;

        // 重建 _history（最近 MaxHistoryMessages 条 user/assistant 对）
        foreach (var m in history.TakeLast(MaxHistoryMessages))
            _history.Add(new { role = m.IsUser ? "user" : "assistant", content = m.Message });

        // 渲染到 WebBrowser（等 LoadCompleted 后执行）
        var sb = new System.Text.StringBuilder();
        foreach (var m in history)
            sb.Append(BuildAgentMessageHtml(m.IsUser ? "user" : "assistant", m.Message, m.IsUser, m.Role, m.Timestamp));
        var escaped = System.Web.HttpUtility.JavaScriptStringEncode(sb.ToString());
        ChatOutput.LoadCompleted += OnFirstLoadInsertHistory;
        _pendingHistoryHtml = escaped;
    }

    private string _pendingHistoryHtml;

    private void OnFirstLoadInsertHistory(object sender, System.Windows.Navigation.NavigationEventArgs e)
    {
        ChatOutput.LoadCompleted -= OnFirstLoadInsertHistory;
        if (string.IsNullOrEmpty(_pendingHistoryHtml))
            return;
        try
        {
            ChatOutput.InvokeScript(
                "eval",
                $"document.body.insertAdjacentHTML('afterbegin','{_pendingHistoryHtml}');window.scrollTo(0,document.body.scrollHeight);"
            );
        }
        catch { }
        _pendingHistoryHtml = null;
    }

    private static string BuildAgentMessageHtml(string role, string markdown, bool isUser, string label, DateTime? timestamp)
    {
        var html = InjectCellLinks(Markdown.ToHtml(markdown, MdPipeline));
        var cls = isUser ? "user" : "assistant";
        var ts = (timestamp ?? DateTime.Now).ToString("HH:mm:ss");
        return $"<div class='msg {cls}'><div class='role'>{label} <span class='ts'>{ts}</span></div><div class='content'>{html}</div></div>";
    }

    /// <summary>给 TextBox 注入 Ctrl+C/X/A/Z/Y 拦截，防止 CTP 内 Excel 截走快捷键</summary>
    private static void AttachClipboardKeys(System.Windows.Controls.TextBox tb)
    {
        tb.PreviewKeyDown += (_, e) =>
        {
            if (Keyboard.Modifiers != ModifierKeys.Control) return;
            switch (e.Key)
            {
                case Key.C: tb.Copy();  e.Handled = true; break;
                case Key.X: tb.Cut();   e.Handled = true; break;
                case Key.A: tb.SelectAll(); e.Handled = true; break;
                case Key.Z: tb.Undo(); e.Handled = true; break;
            }
        };
    }

    private void TaskInput_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter)
            return;
        e.Handled = true;
        if ((e.KeyboardDevice.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
        {
            TaskInput.AppendText(Environment.NewLine);
            return;
        }
        var text = TaskInput.Text.Trim();
        if (string.IsNullOrEmpty(text)) return;
        if (_isRunning)
        {
            // 执行中：插话，不启动新任务
            _interjectQueue.Enqueue(text);
            AppendChat("user", $"💬 {text}");
            TaskInput.Clear();
        }
        else
        {
            RunButton_Click(RunButton, new RoutedEventArgs());
        }
    }

    private void StopButton_Click(object sender, RoutedEventArgs e)
    {
        _cts?.Cancel();
        SetStatus("已停止");
    }

    private async void RunButton_Click(object sender, RoutedEventArgs e)
    {
        var task = TaskInput.Text.Trim();
        if (string.IsNullOrEmpty(task))
            return;

        _cts = new CancellationTokenSource();
        _isRunning = true;
        RunButton.IsEnabled = false;
        StopButton.IsEnabled = true;
        StepsList.Items.Clear();
        AppendChat("user", task);
        TaskInput.Clear();
        TaskInput.IsEnabled = true; // 执行中保持可输入（用于插话）
        TaskInput.Background = new System.Windows.Media.SolidColorBrush(
            System.Windows.Media.Color.FromRgb(0x1a, 0x1a, 0x28)); // 微调颜色提示"插话模式"
        SetStatus("执行中… (输入可插话)");

        try
        {
            await RunAgentLoopAsync(task, _cts.Token);
        }
        catch (OperationCanceledException)
        {
            AddStep("⛔ 已取消");
        }
        catch (Exception ex)
        {
            AddStep($"❌ 错误：{ex.Message}");
            SetStatus("出错");
        }
        finally
        {
            _isRunning = false;
            RunButton.IsEnabled = true;
            StopButton.IsEnabled = false;
            TaskInput.Background = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(0x1e, 0x1e, 0x1e));
        }
    }

    private async Task RunAgentLoopAsync(string userTask, CancellationToken ct)
    {
        var model = ModelComboBox.SelectedItem as string ?? AppServices.Config.Llm.Model;
        var apiKey = AppServices.Config.Llm.ApiKey;
        var apiUrl = AppServices.Config.Llm.ChatCompletionsUrl;
        var maxSteps = (int)(MaxStepsInput.Value ?? 10);

        if (_history.Count == 0)
        {
            var customInstruction = Dispatcher.Invoke(() => CustomInstructionInput.Text.Trim());
            var systemContent =
                "你是一个专业的 Excel 数据助手兼游戏数值策划助手，可以对当前工作簿进行全面操作和分析。\n"
                + "工作流程：1) 先用 get_workbook_structure 或 list_sheets 了解所有打开工作簿的结构；"
                + "2) 用 read_sheet 读取相关数据；"
                + "3) 用 write_range/batch_write 写入结果。\n"
                + "工具选择原则：\n"
                + "- 数据统计分析：用 describe_data / detect_patterns\n"
                + "- 游戏数值计算：用 sim_progression / calc_drop_expectation / balance_check / cost_curve_fit\n"
                + "- 背景色/字体色/粗体/列宽/行高等单元格格式：优先用 apply_format（xlsx 文件可用，无需宏权限）\n"
                + "- 复杂格式/条件格式/图表/去重/筛选等：用 run_vba_macro 编写 VBA 代码执行，不要询问是否可以\n"
                + "- Lua 配置对比：用 list_lua_tables 查看可用表，再用 read_lua_table 读取\n"
                + "- 跨表外键验证：用 check_cross_ref\n"
                + "每次只调用一个工具，等待结果后再决定下一步。完成后用 Markdown 输出简洁的结果说明。\n"
                + "图片操作规则：当用户要求生成图片并插入 Excel 时，必须先调用 generate_image 得到文件路径，再将路径传给 insert_image，两步必须连续执行，不能只做其中一步。\n"
                + "回复规则：每次任务结束时，把工具返回的关键数据（表头、关键值、操作的 Sheet 名、行列号、单元格地址等）写进最终回复，不要只说完成或成功——这些细节是后续修正任务的上下文依据。\n"
                + "重要：如果你在上一条消息中给出了编号选项（如 1. 2. 3.），用户回复单个数字时，请将其理解为选择对应选项，直接执行该选项对应的操作，不要再次询问。";
            if (!string.IsNullOrEmpty(customInstruction))
                systemContent += $"\n\n用户自定义指令（始终遵守）：{customInstruction}";
            _history.Add(new { role = "system", content = systemContent });
        }
        _history.Add(new { role = "user", content = userTask });

        // 防止历史消息无限增长导致 token 超限：保留 system 消息 + 最近 MaxHistoryMessages 条
        if (_history.Count > MaxHistoryMessages + 1)
        {
            var systemMsg = _history.FirstOrDefault();
            var recent = _history.Skip(_history.Count - MaxHistoryMessages).ToList();
            _history.Clear();
            if (systemMsg != null)
                _history.Add(systemMsg);
            _history.AddRange(recent);
        }

        // 工作列表：历史 + 本次任务中间步骤，任务结束后只把最终 assistant 消息写回 _history
        var messages = _history.ToList();
        var historyCountBefore = _history.Count;

        AddStep($"📋 {userTask[..Math.Min(40, userTask.Length)]}…");

        for (var step = 1; step <= maxSteps; step++)
        {
            ct.ThrowIfCancellationRequested();

            // 插话注入：每步开头检查用户是否在执行中发送了新消息
            while (_interjectQueue.TryDequeue(out var injection))
            {
                messages.Add(new { role = "user", content = injection });
                AddStep($"💬 插话：{injection[..Math.Min(40, injection.Length)]}");
            }

            SetStatus($"步骤 {step}/{maxSteps}…");

            var (content, toolCalls) = await CallWithToolsAsync(
                model,
                messages,
                apiKey,
                apiUrl,
                ct
            );

            if (toolCalls is { Count: > 0 })
            {
                messages.Add(
                    new
                    {
                        role = "assistant",
                        content = content ?? "",
                        tool_calls = toolCalls,
                    }
                );

                foreach (var tc in toolCalls)
                {
                    var toolName = tc["function"]?["name"]?.ToString() ?? "";
                    var argsJson = tc["function"]?["arguments"]?.ToString() ?? "{}";
                    var toolCallId = tc["id"]?.ToString() ?? $"tc_{step}";

                    AddStep($"🔧 {toolName}({argsJson[..Math.Min(50, argsJson.Length)]})");
                    var tcs = new TaskCompletionSource<string>();
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try
                        {
                            tcs.SetResult(ExecuteTool(toolName, argsJson));
                        }
                        catch (Exception ex)
                        {
                            tcs.SetResult($"工具执行异常: {ex.Message}");
                        }
                    });
                    var result = await tcs.Task;
                    AddStep($"   ↳ {result[..Math.Min(70, result.Length)]}");

                    messages.Add(
                        new
                        {
                            role = "tool",
                            tool_call_id = toolCallId,
                            content = result,
                        }
                    );
                }
            }
            else
            {
                AddStep($"✅ 完成（{step} 步）");
                SetStatus("完成");
                var finalContent = content ?? "（无输出）";
                AppendChat("assistant", finalContent);
                // 只把最终文本消息写回持久历史（中间 tool_calls/tool 消息丢弃）
                while (_history.Count > historyCountBefore)
                    _history.RemoveAt(_history.Count - 1);
                _history.Add(new { role = "assistant", content = finalContent });
                // 持久化用户消息 + 最终 assistant 消息
                var db = new ChatHistoryManager();
                var model2 = Dispatcher.Invoke(() => ModelComboBox.SelectedItem as string ?? "Agent");
                _ = db.SaveChatMessageAsync(new ChatMessage
                {
                    Role = Environment.UserName, Message = userTask,
                    IsUser = true, Timestamp = DateTime.Now, IsAgent = true,
                });
                _ = db.SaveChatMessageAsync(new ChatMessage
                {
                    Role = model2, Message = finalContent,
                    IsUser = false, Timestamp = DateTime.Now, IsAgent = true,
                });
                return;
            }
        }

        AddStep("⚠️ 已达步骤上限");
        SetStatus("超出步骤上限");
        // 超限：丢弃中间步骤，历史不新增（避免污染下一轮）
        while (_history.Count > historyCountBefore)
            _history.RemoveAt(_history.Count - 1);
    }

    private static async Task<(string content, List<JObject> toolCalls)> CallWithToolsAsync(
        string model,
        List<object> messages,
        string apiKey,
        string apiUrl,
        CancellationToken ct
    )
    {
        var body = new
        {
            model,
            messages,
            tools = ToolDefinitions,
            tool_choice = "auto",
            max_tokens = 4000,
        };
        using var http = new System.Net.Http.HttpClient { Timeout = TimeSpan.FromMinutes(3) };
        using var req = new System.Net.Http.HttpRequestMessage(
            System.Net.Http.HttpMethod.Post,
            apiUrl
        );
        req.Content = new System.Net.Http.StringContent(
            JsonConvert.SerializeObject(body),
            System.Text.Encoding.UTF8,
            "application/json"
        );
        req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue(
            "Bearer",
            apiKey
        );
        using var resp = await http.SendAsync(req, ct);
        var json = JObject.Parse(await resp.Content.ReadAsStringAsync(ct));
        if (!resp.IsSuccessStatusCode)
            throw new Exception(json.ToString());
        var msg = json["choices"]?[0]?["message"];
        return (msg?["content"]?.ToString(), msg?["tool_calls"]?.ToObject<List<JObject>>());
    }

    private static string ExecuteTool(string toolName, string argsJson)
    {
        try
        {
            var args = JObject.Parse(argsJson);
            return toolName switch
            {
                "read_selection" => PubMetToExcel.ArrayToArrayStr(AppServices.App.Selection.Value2),
                "write_range" => ToolWriteRange(
                    args["address"]?.ToString() ?? "",
                    args["value"]?.ToString() ?? ""
                ),
                "run_formula" => ToolRunFormula(
                    args["address"]?.ToString() ?? "",
                    args["formula"]?.ToString() ?? ""
                ),
                "list_udfs" => ToolListUdfs(),
                "read_sheet" => ToolReadSheet(
                    args["sheet_name"]?.ToString() ?? "",
                    (int)(args["max_rows"] ?? 50)
                ),
                "list_sheets" => ToolListSheets(),
                "get_workbook_structure" => ToolGetWorkbookStructure(),
                "batch_write" => ToolBatchWrite(
                    args["sheet_name"]?.ToString() ?? "",
                    args["writes"] as JArray ?? []
                ),
                "run_vba_macro" => ToolRunVbaMacro(args["code"]?.ToString() ?? ""),
                "apply_format" => ToolApplyFormat(args),
                "check_cross_ref" => ToolCheckCrossRef(
                    args["source_sheet"]?.ToString() ?? "",
                    args["source_col"]?.ToString() ?? "",
                    args["target_sheet"]?.ToString() ?? "",
                    args["target_col"]?.ToString() ?? ""
                ),
                "read_lua_table" => ToolReadLuaTable(
                    args["table_name"]?.ToString() ?? "",
                    (int)(args["max_rows"] ?? 100)
                ),
                "list_lua_tables" => ToolListLuaTables(),
                "describe_data" => ToolDescribeData(
                    args["sheet_name"]?.ToString() ?? "",
                    args["range"]?.ToString() ?? ""
                ),
                "detect_patterns" => ToolDetectPatterns(
                    args["sheet_name"]?.ToString() ?? "",
                    args["col_range"]?.ToString() ?? ""
                ),
                "sim_progression" => ToolSimProgression(
                    (double)(args["init_val"] ?? 100),
                    (double)(args["growth_rate"] ?? 1.1),
                    args["growth_type"]?.ToString() ?? "multiply",
                    (int)(args["steps"] ?? 10),
                    args["write_sheet"]?.ToString() ?? "",
                    args["write_start_cell"]?.ToString() ?? ""
                ),
                "calc_drop_expectation" => ToolCalcDropExpectation(
                    args["sheet_name"]?.ToString() ?? "",
                    args["item_col"]?.ToString() ?? "",
                    args["prob_col"]?.ToString() ?? "",
                    (int)(args["trials"] ?? 100)
                ),
                "balance_check" => ToolBalanceCheck(
                    args["sheet_name"]?.ToString() ?? "",
                    args["col_range"]?.ToString() ?? "",
                    (double)(args["min_ratio"] ?? 1.0),
                    (double)(args["max_ratio"] ?? 2.0)
                ),
                "cost_curve_fit" => ToolCostCurveFit(
                    args["sheet_name"]?.ToString() ?? "",
                    args["col_range"]?.ToString() ?? ""
                ),
                // ── P0 新工具 ────────────────────────────────────────────────────
                "get_active_cell"   => ToolGetActiveCell(),
                "navigate_to"       => ToolNavigateTo(
                    args["sheet_name"]?.ToString() ?? "",
                    args["address"]?.ToString() ?? "A1"
                ),
                "manage_sheet"      => ToolManageSheet(
                    args["action"]?.ToString() ?? "",
                    args["sheet_name"]?.ToString() ?? "",
                    args["new_name"]?.ToString(),
                    (int?)(args["target_position"]),
                    args["copy_before"]?.ToString()
                ),
                "save_workbook"     => ToolSaveWorkbook(
                    args["save_as_path"]?.ToString(),
                    args["file_format"]?.ToString()
                ),
                "find_replace"      => ToolFindReplace(
                    args["find"]?.ToString() ?? "",
                    args["replace"]?.ToString() ?? "",
                    args["sheet_name"]?.ToString(),
                    args["address"]?.ToString(),
                    (bool)(args["match_case"] ?? false),
                    (bool)(args["match_entire_cell"] ?? false)
                ),
                "insert_delete_rows_cols" => ToolInsertDeleteRowsCols(
                    args["action"]?.ToString() ?? "",
                    args["sheet_name"]?.ToString() ?? "",
                    (int)(args["start"] ?? 1),
                    (int)(args["count"] ?? 1)
                ),
                "sort_range"        => ToolSortRange(
                    args["sheet_name"]?.ToString() ?? "",
                    args["address"]?.ToString() ?? "",
                    (int)(args["key_column"] ?? 1),
                    (bool)(args["ascending"] ?? true),
                    (bool)(args["has_header"] ?? true)
                ),
                "set_number_format" => ToolSetNumberFormat(
                    args["sheet_name"]?.ToString() ?? "",
                    args["address"]?.ToString() ?? "",
                    args["format_string"]?.ToString() ?? "General"
                ),
                // ── 图片插入 ─────────────────────────────────────────────────────
                "insert_image"      => ToolInsertImage(
                    args["sheet_name"]?.ToString() ?? "",
                    args["file_path"]?.ToString() ?? "",
                    args["anchor_address"]?.ToString() ?? "A1",
                    (double?)(args["width_pt"]),
                    (double?)(args["height_pt"])
                ),
                "download_image"    => ToolDownloadImage(args["url"]?.ToString() ?? ""),
                "generate_image"    => ToolGenerateImage(
                    args["prompt"]?.ToString() ?? "",
                    args["model"]?.ToString() ?? "gemini-3.1-flash-image-preview",
                    AppServices.Config.Llm.ApiKey,
                    AppServices.Config.Llm.ChatCompletionsUrl
                ),
                _ => $"未知工具: {toolName}",
            };
        }
        catch (Exception ex)
        {
            return $"工具执行失败: {ex.Message}";
        }
    }

    private static string ToolWriteRange(string address, string value)
    {
        dynamic ws = AppServices.App.ActiveSheet;
        var range = ws.Range[address];
        var lines = value.Split('\n');
        for (var r = 0; r < lines.Length; r++)
        {
            var cols = lines[r].Split('\t');
            for (var c = 0; c < cols.Length; c++)
                range.Cells[r + 1, c + 1] = cols[c];
        }
        return $"已写入 {address}";
    }

    private static string ToolRunFormula(string address, string formula)
    {
        dynamic ws = AppServices.App.ActiveSheet;
        var cell = ws.Range[address];
        cell.Formula = formula;
        return $"{address} = {cell.Value2}";
    }

    private static string ToolListUdfs()
    {
        var udfs = System
            .Reflection.Assembly.GetExecutingAssembly()
            .GetTypes()
            .SelectMany(t =>
                t.GetMethods(
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static
                )
            )
            .Where(m =>
                m.GetCustomAttributes(
                    typeof(ExcelDna.Integration.ExcelFunctionAttribute),
                    false
                ).Length > 0
            )
            .Select(m => m.Name)
            .OrderBy(n => n)
            .ToList();
        return udfs.Count > 0 ? string.Join(", ", udfs) : "（未找到 UDF）";
    }

    // 推断 Lua 导出目录（复用 ExcelExporter.JsonBaseFolder 同套逻辑）
    private static string GetLuaTablesDir()
    {
        var basePath = AppServices.Config.Paths.BasePath;
        if (string.IsNullOrEmpty(basePath))
            return "";
        string jsonBase;
        if (
            basePath.Contains("Lte资源映射")
            || basePath.Contains("二合")
            || basePath.Contains("工会")
            || basePath.Contains("克朗代克")
        )
            jsonBase = Path.GetFullPath(Path.Combine(basePath, "./../../../../"));
        else if (
            basePath.Contains("Configs")
            || basePath.Contains("UIs")
            || basePath.Contains("Localizations")
        )
            jsonBase = Path.GetFullPath(Path.Combine(basePath, "./../../"));
        else
            jsonBase = Path.GetFullPath(Path.Combine(basePath, "./../../../"));
        return Path.Combine(jsonBase, "Code", "Assets", "LuaScripts", "Tables");
    }

    // 工作簿名→Lua 表名规则：
    //   工作簿名含 $ → 每个 Sheet 各自导出，Lua 表名 = Sheet 名
    //   工作簿名不含 $ → 单 Sheet 模式，Lua 表名 = 工作簿文件名（不含扩展名）
    private static string GetLuaTableName(dynamic ws)
    {
        var wbName = (string)ws.Parent.Name;
        var wbBaseName = Path.GetFileNameWithoutExtension(wbName);
        if (wbBaseName.Contains('$'))
            return (string)ws.Name; // 多Sheet模式：用 Sheet 名
        return wbBaseName; // 单Sheet模式：用工作簿名
    }

    // 在所有打开的工作簿中查找指定 Sheet
    // sheetName 可以是 Sheet 名，也可以是 Lua 表名（会自动匹配）
    private static dynamic FindSheet(string sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            return AppServices.App.ActiveSheet;
        foreach (dynamic wb in AppServices.App.Workbooks)
        {
            foreach (dynamic ws in wb.Sheets)
            {
                // 精确匹配 Sheet 名
                if (string.Equals((string)ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    return ws;
                // 匹配对应的 Lua 表名
                if (
                    string.Equals(
                        GetLuaTableName(ws),
                        sheetName,
                        StringComparison.OrdinalIgnoreCase
                    )
                )
                    return ws;
            }
        }
        return null;
    }

    private static string ToolReadSheet(string sheetName, int maxRows)
    {
        var ws = FindSheet(sheetName);
        if (ws is null)
            return $"找不到 Sheet '{sheetName}'，当前打开的工作簿中无此表";
        var used = ws.UsedRange;
        var rows = Math.Min((int)used.Rows.Count, maxRows);
        var cols = (int)used.Columns.Count;
        var sb = new System.Text.StringBuilder();
        for (var r = 1; r <= rows; r++)
        {
            var line = new List<string>();
            for (var c = 1; c <= cols; c++)
                line.Add(used.Cells[r, c].Value2?.ToString() ?? "");
            sb.AppendLine(string.Join("\t", line));
        }
        return sb.ToString();
    }

    private static string ToolListSheets()
    {
        var sb = new System.Text.StringBuilder();
        foreach (dynamic wb in AppServices.App.Workbooks)
        {
            sb.AppendLine($"[{wb.Name}]");
            foreach (dynamic ws in wb.Sheets)
            {
                var luaName = GetLuaTableName(ws);
                sb.AppendLine($"  Sheet: {ws.Name}  →  Lua表名: {luaName}");
            }
        }
        return sb.ToString();
    }

    private static string ToolGetWorkbookStructure()
    {
        var sb = new System.Text.StringBuilder();
        foreach (dynamic wb in AppServices.App.Workbooks)
        {
            sb.AppendLine($"=== {wb.Name} ===");
            foreach (dynamic ws in wb.Sheets)
            {
                try
                {
                    var used = ws.UsedRange;
                    int rows = used.Rows.Count;
                    int cols = used.Columns.Count;
                    var luaName = GetLuaTableName(ws);
                    sb.AppendLine($"  [{ws.Name}] (Lua:{luaName}) {rows}行×{cols}列");
                    var previewRows = Math.Min(rows, 2);
                    for (var r = 1; r <= previewRows; r++)
                    {
                        var line = new List<string>();
                        for (var c = 1; c <= Math.Min(cols, 10); c++)
                            line.Add(used.Cells[r, c].Value2?.ToString() ?? "");
                        sb.AppendLine("    " + string.Join("\t", line));
                    }
                }
                catch
                {
                    sb.AppendLine($"  [{ws.Name}] (无法读取)");
                }
            }
        }
        return sb.ToString();
    }

    private static string ToolBatchWrite(string sheetName, JArray writes)
    {
        dynamic ws = string.IsNullOrEmpty(sheetName)
            ? AppServices.App.ActiveSheet
            : AppServices.App.ActiveWorkbook.Sheets[sheetName];
        var count = 0;
        foreach (var item in writes)
        {
            var address = item["address"]?.ToString() ?? "";
            var value = item["value"]?.ToString() ?? "";
            if (string.IsNullOrEmpty(address))
                continue;
            var lines = value.Split('\n');
            var range = ws.Range[address];
            for (var r = 0; r < lines.Length; r++)
            {
                var cols = lines[r].Split('\t');
                for (var c = 0; c < cols.Length; c++)
                    range.Cells[r + 1, c + 1] = cols[c];
            }
            count++;
        }
        return $"已写入 {count} 个区域";
    }

    private static string ToolRunVbaMacro(string code)
    {
        try
        {
            dynamic wb = AppServices.App.ActiveWorkbook;
            // VBProject 仅在宏工作簿(.xlsm/.xls/.xlam)中存在；xlsx 访问会抛 COMException
            string wbName = wb.Name?.ToString() ?? "";
            if (
                !wbName.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase)
                && !wbName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase)
                && !wbName.EndsWith(".xlam", StringComparison.OrdinalIgnoreCase)
            )
                return "VBA 执行失败: 当前工作簿不支持宏（需要 .xlsm/.xls/.xlam 格式）";
            var wasSaved = (bool)wb.Saved;
            var vbProj = wb.VBProject;
            var module = vbProj.VBComponents.Add(1); // vbext_ct_StdModule = 1
            module.CodeModule.AddFromString(code);
            var subName = "Main";
            var match = System.Text.RegularExpressions.Regex.Match(code, @"Sub\s+(\w+)\s*\(");
            if (match.Success)
                subName = match.Groups[1].Value;
            // 不带工作簿前缀，避免中文/特殊字符文件名导致 App.Run 报错
            var macroName = $"{module.Name}.{subName}";
            PluginLog.Write($"[VBA] App.Run({macroName})");
            AppServices.App.Run(macroName);
            vbProj.VBComponents.Remove(module);
            wb.Saved = wasSaved;
            return "VBA 执行完成";
        }
        catch (Exception ex)
        {
            PluginLog.Write($"[VBA] 失败: {ex.GetType().Name}: {ex.Message}");
            return $"VBA 执行失败: {ex.Message}";
        }
    }

    private static string ToolApplyFormat(JObject args)
    {
        try
        {
            var rangeAddr = args["range"]?.ToString() ?? "";
            if (string.IsNullOrEmpty(rangeAddr))
                return "apply_format 失败：缺少 range 参数";

            var sheetName = args["sheet_name"]?.ToString() ?? "";
            dynamic ws = string.IsNullOrEmpty(sheetName)
                ? AppServices.App.ActiveSheet
                : AppServices.App.ActiveWorkbook.Sheets[sheetName];
            dynamic rng = ws.Range[rangeAddr];

            if (args["bg_color"] is { } bgToken)
            {
                var hex = bgToken.ToString().TrimStart('#');
                var rgb = Convert.ToInt32(hex, 16);
                // Excel Interior.Color = BGR int
                var r = (rgb >> 16) & 0xFF;
                var g = (rgb >> 8) & 0xFF;
                var b = rgb & 0xFF;
                rng.Interior.Color = b << 16 | g << 8 | r;
            }

            if (args["font_color"] is { } fcToken)
            {
                var hex = fcToken.ToString().TrimStart('#');
                var rgb = Convert.ToInt32(hex, 16);
                var r = (rgb >> 16) & 0xFF;
                var g = (rgb >> 8) & 0xFF;
                var b = rgb & 0xFF;
                rng.Font.Color = b << 16 | g << 8 | r;
            }

            if (args["bold"] is { } boldToken)
                rng.Font.Bold = boldToken.ToObject<bool>();

            if (args["italic"] is { } italicToken)
                rng.Font.Italic = italicToken.ToObject<bool>();

            if (args["font_size"] is { } fsToken)
                rng.Font.Size = fsToken.ToObject<double>();

            if (args["wrap_text"] is { } wrapToken)
                rng.WrapText = wrapToken.ToObject<bool>();

            if (args["h_align"] is { } alignToken)
                rng.HorizontalAlignment = alignToken.ToString().ToLower() switch
                {
                    "left" => -4131, // xlLeft
                    "right" => -4152, // xlRight
                    "center" => -4108, // xlCenter
                    _ => -4108,
                };

            if (args["col_width"] is { } cwToken)
                rng.Columns[1].ColumnWidth = cwToken.ToObject<double>();

            if (args["row_height"] is { } rhToken)
                rng.Rows[1].RowHeight = rhToken.ToObject<double>();

            return $"格式已应用到 {rangeAddr}";
        }
        catch (Exception ex)
        {
            return $"apply_format 失败: {ex.Message}";
        }
    }

    private void AddStep(string text) =>
        Dispatcher.Invoke(() =>
        {
            StepsList.Items.Add(text);
            StepsScroll.ScrollToBottom();
        });

    private static string ToolCheckCrossRef(
        string sourceSheet,
        string sourceCol,
        string targetSheet,
        string targetCol
    )
    {
        var srcWs = FindSheet(sourceSheet);
        var tgtWs = FindSheet(targetSheet);
        if (srcWs is null)
            return $"找不到 Sheet '{sourceSheet}'";
        if (tgtWs is null)
            return $"找不到 Sheet '{targetSheet}'";

        // 找列索引（支持列名如"activityID"或列字母如"B"）
        int FindColIndex(dynamic ws, string colName)
        {
            var used = ws.UsedRange;
            int cols = used.Columns.Count;
            // 先尝试按标题行匹配
            for (var c = 1; c <= cols; c++)
            {
                var header = used.Cells[1, c].Value2?.ToString() ?? "";
                if (string.Equals(header, colName, StringComparison.OrdinalIgnoreCase))
                    return c;
            }
            // 再尝试按列字母（A/B/C...）
            if (colName.Length <= 2 && colName.All(char.IsLetter))
            {
                var idx = 0;
                foreach (var ch in colName.ToUpper())
                    idx = idx * 26 + (ch - 'A' + 1);
                return idx;
            }
            return -1;
        }

        var srcColIdx = FindColIndex(srcWs, sourceCol);
        var tgtColIdx = FindColIndex(tgtWs, targetCol);
        if (srcColIdx < 0)
            return $"找不到列 '{sourceCol}' 在 Sheet '{sourceSheet}'";
        if (tgtColIdx < 0)
            return $"找不到列 '{targetCol}' 在 Sheet '{targetSheet}'";

        // 收集合法值集合
        var tgtUsed = tgtWs.UsedRange;
        int tgtRows = tgtUsed.Rows.Count;
        var validSet = new HashSet<string>();
        for (var r = 2; r <= tgtRows; r++)
        {
            var v = tgtUsed.Cells[r, tgtColIdx].Value2?.ToString() ?? "";
            if (!string.IsNullOrEmpty(v))
                validSet.Add(v);
        }

        // 检查来源列
        var srcUsed = srcWs.UsedRange;
        int srcRows = srcUsed.Rows.Count;
        var missing = new List<string>();
        for (var r = 2; r <= srcRows; r++)
        {
            var v = srcUsed.Cells[r, srcColIdx].Value2?.ToString() ?? "";
            if (string.IsNullOrEmpty(v))
                continue;
            if (!validSet.Contains(v))
                missing.Add($"行{r}: {v}");
        }

        if (missing.Count == 0)
            return $"✅ 全部 {srcRows - 1} 条记录均合法，无缺失引用";
        return $"❌ 发现 {missing.Count} 条无效引用：\n" + string.Join("\n", missing);
    }

    private static string ToolDescribeData(string sheetName, string rangeAddress)
    {
        dynamic ws = string.IsNullOrEmpty(sheetName)
            ? AppServices.App.ActiveSheet
            : FindSheet(sheetName);
        if (ws is null)
            return $"找不到 Sheet '{sheetName}'";
        dynamic range = string.IsNullOrEmpty(rangeAddress)
            ? AppServices.App.Selection
            : ws.Range[rangeAddress];
        var values = range.Value2 as object[,];
        if (values is null)
            return "选区为空";

        int rows = values.GetLength(0);
        int cols = values.GetLength(1);
        int total = rows * cols;
        int empty = 0;
        int numeric = 0;
        int text = 0;
        var nums = new List<double>();

        for (var r = 1; r <= rows; r++)
        for (var c = 1; c <= cols; c++)
        {
            var v = values[r, c];
            if (v is null || v.ToString() == "")
                empty++;
            else if (v is double d)
            {
                numeric++;
                nums.Add(d);
            }
            else if (double.TryParse(v.ToString(), out var p))
            {
                numeric++;
                nums.Add(p);
            }
            else
                text++;
        }

        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"范围：{rows}行 × {cols}列，共 {total} 个单元格");
        sb.AppendLine($"空值：{empty}（{empty * 100 / total}%）");
        sb.AppendLine($"数值型：{numeric}，文本型：{text}");
        if (nums.Count > 0)
        {
            nums.Sort();
            var avg = nums.Average();
            var min = nums[0];
            var max = nums[^1];
            var median = nums[nums.Count / 2];
            var variance = nums.Average(x => (x - avg) * (x - avg));
            sb.AppendLine($"数值范围：{min} ~ {max}");
            sb.AppendLine($"均值：{avg:F2}，中位数：{median}，标准差：{Math.Sqrt(variance):F2}");
        }
        return sb.ToString();
    }

    private static string ToolDetectPatterns(string sheetName, string colAddress)
    {
        dynamic ws = string.IsNullOrEmpty(sheetName)
            ? AppServices.App.ActiveSheet
            : FindSheet(sheetName);
        if (ws is null)
            return $"找不到 Sheet '{sheetName}'";
        dynamic range = ws.Range[colAddress];
        var values = range.Value2 as object[,];
        if (values is null)
            return "无数据";

        int rows = values.GetLength(0);
        var nums = new List<(int row, double val)>();
        for (var r = 1; r <= rows; r++)
        {
            var v = values[r, 1];
            if (v is double d)
                nums.Add((r, d));
            else if (double.TryParse(v?.ToString(), out var p))
                nums.Add((r, p));
        }
        if (nums.Count < 2)
            return "数值不足，无法分析";

        var sb = new System.Text.StringBuilder();
        var vals = nums.Select(x => x.val).ToList();
        var avg = vals.Average();
        var std = Math.Sqrt(vals.Average(x => (x - avg) * (x - avg)));

        // 异常值（偏离均值 3σ）
        var anomalies = nums.Where(x => Math.Abs(x.val - avg) > 3 * std).ToList();
        if (anomalies.Count > 0)
            sb.AppendLine(
                $"⚠️ 异常值（3σ）："
                    + string.Join(", ", anomalies.Select(x => $"行{x.row}={x.val}"))
            );
        else
            sb.AppendLine("✅ 无异常值（3σ内）");

        // 趋势
        var diffs = Enumerable
            .Range(0, nums.Count - 1)
            .Select(i => nums[i + 1].val - nums[i].val)
            .ToList();
        var posCount = diffs.Count(d => d > 0);
        var negCount = diffs.Count(d => d < 0);
        if (posCount > nums.Count * 0.8)
            sb.AppendLine("📈 趋势：持续递增");
        else if (negCount > nums.Count * 0.8)
            sb.AppendLine("📉 趋势：持续递减");
        else
            sb.AppendLine("↕️ 趋势：无明显单调性");

        // 重复值
        var dupes = vals.GroupBy(x => x)
            .Where(g => g.Count() > 1)
            .Select(g => $"{g.Key}×{g.Count()}")
            .ToList();
        if (dupes.Count > 0)
            sb.AppendLine($"🔁 重复值：" + string.Join(", ", dupes));

        return sb.ToString();
    }

    private static string ToolSimProgression(
        double initVal,
        double growthRate,
        string growthType,
        int steps,
        string writeSheet,
        string writeStartCell
    )
    {
        var result = new List<double>();
        var cur = initVal;
        for (var i = 0; i < steps; i++)
        {
            result.Add(cur);
            cur = growthType switch
            {
                "linear" => cur + growthRate,
                "multiply" => cur * growthRate,
                "power" => initVal * Math.Pow(i + 2, growthRate),
                _ => cur * growthRate,
            };
        }
        if (!string.IsNullOrEmpty(writeSheet) && !string.IsNullOrEmpty(writeStartCell))
        {
            var ws = FindSheet(writeSheet) ?? AppServices.App.ActiveSheet;
            dynamic range = ws.Range[writeStartCell];
            for (var i = 0; i < result.Count; i++)
                range.Cells[i + 1, 1] = Math.Round(result[i], 2);
            return $"已写入 {steps} 步数据到 {writeSheet}!{writeStartCell}";
        }
        return "模拟结果：\n" + string.Join("\n", result.Select((v, i) => $"第{i + 1}步: {v:F2}"));
    }

    private static string ToolCalcDropExpectation(
        string sheetName,
        string itemCol,
        string probCol,
        int trials
    )
    {
        var ws = FindSheet(sheetName) ?? AppServices.App.ActiveSheet;
        dynamic used = ws.UsedRange;
        int rows = used.Rows.Count;

        // 找列索引
        int FindCol(string colName)
        {
            int cols2 = used.Columns.Count;
            for (var c = 1; c <= cols2; c++)
                if (
                    string.Equals(
                        used.Cells[1, c].Value2?.ToString(),
                        colName,
                        StringComparison.OrdinalIgnoreCase
                    )
                )
                    return c;
            if (colName.Length <= 2 && colName.All(char.IsLetter))
            {
                var idx = 0;
                foreach (var ch in colName.ToUpper())
                    idx = idx * 26 + (ch - 'A' + 1);
                return idx;
            }
            return -1;
        }

        var ic = FindCol(itemCol);
        var pc = FindCol(probCol);
        if (ic < 0 || pc < 0)
            return $"找不到列 '{itemCol}' 或 '{probCol}'";

        var items = new List<(string name, double prob)>();
        double totalProb = 0;
        for (var r = 2; r <= rows; r++)
        {
            var name = used.Cells[r, ic].Value2?.ToString() ?? "";
            if (string.IsNullOrEmpty(name))
                continue;
            if (!double.TryParse(used.Cells[r, pc].Value2?.ToString(), out double prob))
                continue;
            items.Add((name, prob));
            totalProb += prob;
        }
        if (items.Count == 0)
            return "未找到有效掉落数据";

        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"掉落表分析（{trials} 次抽取期望）：");
        var probStatus =
            Math.Abs(totalProb - 1) < 0.01 ? "正常✅" : $"⚠️偏差{Math.Abs(totalProb - 1):F4}";
        sb.AppendLine($"总概率：{totalProb:F4}（{probStatus}）");
        foreach (var (name, prob) in items)
        {
            var exp = prob * trials;
            var stdDev = Math.Sqrt(trials * prob * (1 - prob));
            var pAtLeastOne = 1 - Math.Pow(1 - prob, trials);
            sb.AppendLine(
                $"  {name}: 概率={prob:P2}, 期望={exp:F1}次, σ={stdDev:F1}, 至少1次概率={pAtLeastOne:P1}"
            );
        }
        return sb.ToString();
    }

    private static string ToolBalanceCheck(
        string sheetName,
        string colAddress,
        double minRatio,
        double maxRatio
    )
    {
        var ws = FindSheet(sheetName) ?? AppServices.App.ActiveSheet;
        dynamic range = ws.Range[colAddress];
        var values = range.Value2 as object[,];
        if (values is null)
            return "无数据";

        int rows = values.GetLength(0);
        var nums = new List<(int row, double val)>();
        for (var r = 1; r <= rows; r++)
        {
            var v = values[r, 1];
            if (v is double d)
                nums.Add((r, d));
            else if (double.TryParse(v?.ToString(), out var p))
                nums.Add((r, p));
        }
        if (nums.Count < 2)
            return "数值不足";

        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"平衡性检查（相邻增长比约束：{minRatio:F2} ~ {maxRatio:F2}）：");
        var violations = new List<string>();
        for (var i = 1; i < nums.Count; i++)
        {
            if (nums[i - 1].val == 0)
                continue;
            var ratio = nums[i].val / nums[i - 1].val;
            if (ratio < minRatio || ratio > maxRatio)
                violations.Add(
                    $"行{nums[i].row}: {nums[i - 1].val}→{nums[i].val}，比值={ratio:F3}"
                );
        }
        if (violations.Count == 0)
            sb.AppendLine($"✅ 全部 {nums.Count - 1} 个相邻比值均在范围内");
        else
        {
            sb.AppendLine($"❌ 发现 {violations.Count} 处违规：");
            violations.ForEach(v => sb.AppendLine("  " + v));
        }
        return sb.ToString();
    }

    private static string ToolCostCurveFit(string sheetName, string colAddress)
    {
        var ws = FindSheet(sheetName) ?? AppServices.App.ActiveSheet;
        dynamic range = ws.Range[colAddress];
        var values = range.Value2 as object[,];
        if (values is null)
            return "无数据";

        int rows = values.GetLength(0);
        var nums = new List<double>();
        for (var r = 1; r <= rows; r++)
        {
            var v = values[r, 1];
            if (v is double d)
                nums.Add(d);
            else if (double.TryParse(v?.ToString(), out var p))
                nums.Add(p);
        }
        if (nums.Count < 3)
            return "数据点不足（需至少3个）";

        int n = nums.Count;
        var x = Enumerable.Range(1, n).Select(i => (double)i).ToList();

        // 线性拟合 y = a + bx
        double sx = x.Sum(),
            sy = nums.Sum();
        double sxx = x.Sum(v => v * v),
            sxy = x.Zip(nums, (a, b) => a * b).Sum();
        double b_lin = (n * sxy - sx * sy) / (n * sxx - sx * sx);
        double a_lin = (sy - b_lin * sx) / n;
        double r2_lin =
            1
            - nums.Zip(x, (y, xi) => Math.Pow(y - (a_lin + b_lin * xi), 2)).Sum()
                / nums.Select(y => Math.Pow(y - nums.Average(), 2)).Sum();

        // 指数拟合 y = a * e^(bx)，对 ln(y) 做线性拟合
        var lnY = nums.Where(v => v > 0).Select(v => Math.Log(v)).ToList();
        double r2_exp = 0;
        double a_exp = 0,
            b_exp = 0;
        if (lnY.Count == n)
        {
            double sly = lnY.Sum(),
                slxy = x.Zip(lnY, (a, b) => a * b).Sum();
            b_exp = (n * slxy - sx * sly) / (n * sxx - sx * sx);
            a_exp = Math.Exp((sly - b_exp * sx) / n);
            r2_exp =
                1
                - nums.Zip(x, (y, xi) => Math.Pow(y - a_exp * Math.Exp(b_exp * xi), 2)).Sum()
                    / nums.Select(y => Math.Pow(y - nums.Average(), 2)).Sum();
        }

        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"曲线拟合结果（{n} 个数据点）：");
        sb.AppendLine($"线性：y = {a_lin:F2} + {b_lin:F2}x，R²={r2_lin:F4}");
        sb.AppendLine($"指数：y = {a_exp:F2} × e^({b_exp:F4}x)，R²={r2_exp:F4}");
        var best = r2_lin >= r2_exp ? "线性" : "指数";
        sb.AppendLine($"推荐：{best}（R²更高）");
        return sb.ToString();
    }

    private static string ToolListLuaTables()
    {
        var luaDir = GetLuaTablesDir();
        if (string.IsNullOrEmpty(luaDir))
            return "无法推断 Lua 目录，请确认 BasePath 配置正确";
        if (!Directory.Exists(luaDir))
            return $"目录不存在：{luaDir}";
        var files = Directory
            .GetFiles(luaDir, "*.lua")
            .Select(f => Path.GetFileNameWithoutExtension(f))
            .OrderBy(n => n)
            .ToList();
        return files.Count > 0
            ? $"共 {files.Count} 个 Lua 表：\n" + string.Join(", ", files)
            : "目录为空";
    }

    private static string ToolReadLuaTable(string tableName, int maxRows)
    {
        var luaDir = GetLuaTablesDir();
        if (string.IsNullOrEmpty(luaDir))
            return "无法推断 Lua 目录，请确认 BasePath 配置正确";
        // 支持 .lua 和 .lua.txt 两种后缀
        var luaFile = Path.Combine(luaDir, tableName + ".lua");
        if (!File.Exists(luaFile))
            luaFile = Path.Combine(luaDir, tableName + ".lua.txt");
        if (!File.Exists(luaFile))
            return $"找不到 Lua 文件：{luaFile}";

        try
        {
            using var lua = new NLua.Lua();
            lua.DoFile(luaFile);

            // Lua 导出文件通常是 local data = {...} return data 或 tableName = {...}
            var tableObj = lua[tableName] ?? lua["data"] ?? lua.GetTable(tableName);
            if (tableObj is not NLua.LuaTable table)
                return $"无法解析 Lua 表 '{tableName}'，请确认文件结构";

            var sb = new System.Text.StringBuilder();
            var count = 0;
            foreach (var key in table.Keys)
            {
                if (count >= maxRows)
                {
                    sb.AppendLine($"... (已截断，共显示 {maxRows} 行)");
                    break;
                }
                var row = table[key];
                if (row is NLua.LuaTable rowTable)
                {
                    var fields = new List<string>();
                    foreach (var fk in rowTable.Keys)
                        fields.Add($"{fk}={rowTable[fk]}");
                    sb.AppendLine($"[{key}] " + string.Join(", ", fields));
                }
                else
                {
                    sb.AppendLine($"[{key}] {row}");
                }
                count++;
            }
            return sb.Length > 0 ? sb.ToString() : "(空表)";
        }
        catch (Exception ex)
        {
            return $"Lua 解析失败: {ex.Message}";
        }
    }

    private void SetStatus(string text) => Dispatcher.Invoke(() => StatusText.Text = text);

    private static string InjectCellLinks(string html) =>
        CellAddressRegex.Replace(
            html,
            m =>
            {
                var sheet = m.Groups[1].Value;
                var cell = m.Groups[2].Value;
                var address = string.IsNullOrEmpty(sheet) ? cell : $"{sheet}!{cell}";
                var encoded = Uri.EscapeDataString(address);
                return $"<a href='excel://cell/{encoded}' style='color:#4ec9b0;text-decoration:none' title='定位到 {address}'>{m.Value}</a>";
            }
        );

    private void AppendChat(string role, string markdown)
    {
        Dispatcher.Invoke(() =>
        {
            var html = InjectCellLinks(Markdown.ToHtml(markdown, MdPipeline));
            var cls = role == "user" ? "user" : "assistant";
            var label =
                role == "user"
                    ? Environment.UserName
                    : (ModelComboBox.SelectedItem as string ?? "Agent");
            var ts = DateTime.Now.ToString("HH:mm:ss");
            var block =
                $"<div class='msg {cls}'><div class='role'>{label} <span class='ts'>{ts}</span></div><div class='content'>{html}</div></div>";
            ChatOutput.InvokeScript(
                "eval",
                $"document.body.insertAdjacentHTML('beforeend','{HttpUtility.JavaScriptStringEncode(block)}');window.scrollTo(0,document.body.scrollHeight);"
            );
        });
    }

    // ── P0 新工具实现 ─────────────────────────────────────────────────────────

    private static string ToolGetActiveCell()
    {
        var cell = AppServices.App.ActiveCell;
        var sheet = (string)AppServices.App.ActiveSheet.Name;
        return $"Sheet: {sheet}, Address: {cell.Address}, Value: {cell.Value2}";
    }

    private static string ToolNavigateTo(string sheetName, string address)
    {
        var ws = FindSheet(sheetName) ?? AppServices.App.ActiveSheet;
        ws.Activate();
        ws.Range[address].Select();
        return $"已跳转到 {ws.Name}!{address}";
    }

    private static string ToolManageSheet(string action, string sheetName, string? newName,
        int? targetPosition, string? copyBefore)
    {
        var wb = AppServices.App.ActiveWorkbook;
        switch (action)
        {
            case "create":
                var ns = wb.Sheets.Add();
                ns.Name = string.IsNullOrEmpty(newName) ? sheetName : newName;
                return $"已创建 Sheet: {ns.Name}";
            case "delete":
                var ds = FindSheet(sheetName);
                if (ds == null) return $"Sheet 不存在: {sheetName}";
                AppServices.App.DisplayAlerts = false;
                ds.Delete();
                AppServices.App.DisplayAlerts = true;
                return $"已删除 Sheet: {sheetName}";
            case "rename":
                var rs = FindSheet(sheetName);
                if (rs == null) return $"Sheet 不存在: {sheetName}";
                rs.Name = newName ?? sheetName;
                return $"已重命名为: {rs.Name}";
            case "copy":
                var cs = FindSheet(sheetName);
                if (cs == null) return $"Sheet 不存在: {sheetName}";
                var before = string.IsNullOrEmpty(copyBefore) ? null : FindSheet(copyBefore);
                if (before != null) cs.Copy(before); else cs.Copy();
                return $"已复制 Sheet: {sheetName}";
            case "move":
                var ms = FindSheet(sheetName);
                if (ms == null) return $"Sheet 不存在: {sheetName}";
                if (targetPosition.HasValue)
                    ms.Move(wb.Sheets[targetPosition.Value]);
                return $"已移动 Sheet: {sheetName}";
            default:
                return $"未知 action: {action}";
        }
    }

    private static string ToolSaveWorkbook(string? saveAsPath, string? fileFormat)
    {
        var wb = AppServices.App.ActiveWorkbook;
        if (string.IsNullOrEmpty(saveAsPath))
        {
            wb.Save();
            return "工作簿已保存";
        }
        var fmt = fileFormat?.ToLower() switch
        {
            "csv"  => Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV,
            "xlsm" => Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled,
            _      => Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook,
        };
        wb.SaveAs(saveAsPath, fmt);
        return $"已另存为: {saveAsPath}";
    }

    private static string ToolFindReplace(string find, string replace, string? sheetName,
        string? address, bool matchCase, bool matchEntireCell)
    {
        dynamic rng;
        if (!string.IsNullOrEmpty(sheetName))
        {
            var ws = FindSheet(sheetName) ?? AppServices.App.ActiveSheet;
            rng = string.IsNullOrEmpty(address) ? ws.UsedRange : ws.Range[address];
        }
        else
        {
            rng = AppServices.App.ActiveSheet.UsedRange;
        }
        rng.Replace(find, replace,
            matchCase ? Microsoft.Office.Interop.Excel.XlLookAt.xlWhole : Microsoft.Office.Interop.Excel.XlLookAt.xlPart,
            Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
            matchCase, Type.Missing, Type.Missing, Type.Missing);
        return $"查找替换完成：\"{find}\" → \"{replace}\"";
    }

    private static string ToolInsertDeleteRowsCols(string action, string sheetName, int start, int count)
    {
        dynamic ws = FindSheet(sheetName) ?? AppServices.App.ActiveSheet;
        var end = start + count - 1;
        switch (action)
        {
            case "insert_rows":
                ws.Range[ws.Cells[start, 1], ws.Cells[end, 1]].EntireRow.Insert(
                    Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                return $"已在第 {start} 行上方插入 {count} 行";
            case "delete_rows":
                ws.Range[ws.Cells[start, 1], ws.Cells[end, 1]].EntireRow.Delete(
                    Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);
                return $"已删除第 {start}–{end} 行";
            case "insert_cols":
                ws.Range[ws.Cells[1, start], ws.Cells[1, end]].EntireColumn.Insert(
                    Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftToRight);
                return $"已在第 {start} 列左侧插入 {count} 列";
            case "delete_cols":
                ws.Range[ws.Cells[1, start], ws.Cells[1, end]].EntireColumn.Delete(
                    Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftToLeft);
                return $"已删除第 {start}–{end} 列";
            default:
                return $"未知 action: {action}";
        }
    }

    private static string ToolSortRange(string sheetName, string address, int keyColumn,
        bool ascending, bool hasHeader)
    {
        var ws = FindSheet(sheetName) ?? AppServices.App.ActiveSheet;
        dynamic rng = ws.Range[address];
        var key = rng.Columns[keyColumn];
        var order = ascending
            ? Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending
            : Microsoft.Office.Interop.Excel.XlSortOrder.xlDescending;
        var header = hasHeader
            ? Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes
            : Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo;
        rng.Sort(key, order, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, header);
        return $"已排序 {address}，键列 {keyColumn}，{(ascending ? "升序" : "降序")}";
    }

    private static string ToolSetNumberFormat(string sheetName, string address, string formatString)
    {
        var ws = FindSheet(sheetName) ?? AppServices.App.ActiveSheet;
        ws.Range[address].NumberFormat = formatString;
        return $"已设置 {address} 数字格式为: {formatString}";
    }

    private static string ToolInsertImage(string sheetName, string filePath, string anchorAddress,
        double? widthPt, double? heightPt)
    {
        if (!File.Exists(filePath))
            return $"文件不存在: {filePath}";
        var ws = FindSheet(sheetName) ?? AppServices.App.ActiveSheet;
        dynamic anchor = ws.Range[anchorAddress];
        float left = (float)(double)anchor.Left;
        float top  = (float)(double)anchor.Top;
        float w    = widthPt.HasValue  ? (float)widthPt.Value  : -1;
        float h    = heightPt.HasValue ? (float)heightPt.Value : -1;
        ws.Shapes.AddPicture(filePath,
            Microsoft.Office.Core.MsoTriState.msoFalse,
            Microsoft.Office.Core.MsoTriState.msoTrue,
            left, top, w, h);
        return $"已插入图片 {Path.GetFileName(filePath)} 到 {anchorAddress}";
    }

    /// <summary>从 URL 下载图片到本地临时文件，返回本地路径</summary>
    private static string ToolDownloadImage(string url)
    {
        try
        {
            using var client = new System.Net.Http.HttpClient();
            client.Timeout = TimeSpan.FromSeconds(30);
            var bytes = client.GetByteArrayAsync(url).GetAwaiter().GetResult();
            var ext   = Path.GetExtension(new Uri(url).AbsolutePath);
            if (string.IsNullOrEmpty(ext)) ext = ".png";
            var path = Path.Combine(Path.GetTempPath(), $"ndtools_img_{DateTime.Now.Ticks}{ext}");
            File.WriteAllBytes(path, bytes);
            return $"下载成功: {path}";
        }
        catch (Exception ex)
        {
            return $"下载失败: {ex.Message}";
        }
    }

    private static string SaveBase64ToPng(string b64)
    {
        var bytes = Convert.FromBase64String(b64);
        var path = Path.Combine(Path.GetTempPath(), $"ndtools_gen_{DateTime.Now.Ticks}.png");
        File.WriteAllBytes(path, bytes);
        return $"图片已生成: {path}";
    }

    /// <summary>
    /// 生成图片三阶段：
    /// 1. /images/generations（LiteLLM 图片端点，gemini-3.1-flash-image-preview 已验证可用）
    /// 2. chat/completions multimodal 回退（解析 message.images[0].image_url.url）
    /// 3. Pollinations.ai 免费公网兜底
    /// </summary>
    private static string ToolGenerateImage(string prompt, string model, string apiKey, string apiUrl)
    {
        try
        {
            // 先尝试 OpenAI images/generations 端点
            var imgUrl = apiUrl.Replace("/chat/completions", "/images/generations");
            using var client = new System.Net.Http.HttpClient();
            client.Timeout = TimeSpan.FromSeconds(60);
            client.DefaultRequestHeaders.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);

            var body = Newtonsoft.Json.JsonConvert.SerializeObject(new
            {
                model,
                prompt,
                n = 1,
                size = "1024x1024",
                response_format = "b64_json",
            });
            var resp = client.PostAsync(imgUrl,
                new System.Net.Http.StringContent(body, System.Text.Encoding.UTF8, "application/json"))
                .GetAwaiter().GetResult();
            var json = Newtonsoft.Json.Linq.JObject.Parse(resp.Content.ReadAsStringAsync().GetAwaiter().GetResult());

            // 阶段1: /images/generations（已验证 gemini-3.1-flash-image-preview 可用）
            if (resp.IsSuccessStatusCode)
            {
                var b64 = json["data"]?[0]?["b64_json"]?.ToString();
                if (!string.IsNullOrEmpty(b64)) return SaveBase64ToPng(b64);
                var imgSrc = json["data"]?[0]?["url"]?.ToString();
                if (!string.IsNullOrEmpty(imgSrc)) return ToolDownloadImage(imgSrc);
            }

            // 阶段2: chat/completions multimodal 回退（LiteLLM OpenAI 兼容层格式）
            var chatBody = Newtonsoft.Json.JsonConvert.SerializeObject(new
            {
                model,
                messages = new[] { new { role = "user", content = prompt } },
                max_tokens = 4096,
            });
            var chatResp = client.PostAsync(apiUrl,
                new System.Net.Http.StringContent(chatBody, System.Text.Encoding.UTF8, "application/json"))
                .GetAwaiter().GetResult();
            var chatJson = Newtonsoft.Json.Linq.JObject.Parse(
                chatResp.Content.ReadAsStringAsync().GetAwaiter().GetResult());

            // LiteLLM 封装后路径: message.images[0].image_url.url = "data:image/png;base64,..."
            var images = chatJson["choices"]?[0]?["message"]?["images"] as Newtonsoft.Json.Linq.JArray;
            if (images is { Count: > 0 })
            {
                var dataUrl = images[0]?["image_url"]?["url"]?.ToString();
                if (!string.IsNullOrEmpty(dataUrl))
                {
                    var idx = dataUrl.IndexOf("base64,", StringComparison.Ordinal);
                    if (idx >= 0) return SaveBase64ToPng(dataUrl[(idx + 7)..]);
                }
            }
            // 兜底：content 里内联的 data URL
            var content = chatJson["choices"]?[0]?["message"]?["content"]?.ToString() ?? "";
            if (content.Contains("data:image"))
            {
                var start = content.IndexOf("base64,", StringComparison.Ordinal);
                if (start >= 0)
                {
                    var end = content.IndexOfAny(new[] { '"', ' ', '\n' }, start);
                    var slice = end > 0 ? content[(start + 7)..end] : content[(start + 7)..];
                    if (!string.IsNullOrEmpty(slice)) return SaveBase64ToPng(slice);
                }
            }

            // 阶段3: Pollinations.ai 免费兜底（无需密钥）
            PluginLog.Write("[generate_image] LiteLLM 不支持，fallback → Pollinations.ai");
            var pollinationsUrl =
                $"https://image.pollinations.ai/prompt/{Uri.EscapeDataString(prompt)}?width=1024&height=1024&nologo=true";
            return ToolDownloadImage(pollinationsUrl);
        }
        catch (Exception ex)
        {
            return $"生成失败: {ex.Message}";
        }
    }
}
