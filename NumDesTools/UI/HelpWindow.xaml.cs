using System.Collections.Generic;
using System.Windows.Controls;
using WpfWindow = System.Windows.Window;

namespace NumDesTools.UI
{
    public partial class HelpWindow : WpfWindow
    {
        // ── 数据模型 ─────────────────────────────────────────────────────────────

        private record HelpItem(string Title, string Html);

        private record HelpGroup(string Name, List<HelpItem> Items);

        // ── 帮助内容（覆盖所有 Ribbon 命令）────────────────────────────────────

        private static readonly List<HelpGroup> HelpData =
            new()
            {
                new(
                    "格式检查",
                    new()
                    {
                        new(
                            "标准格式",
                            @"
<h2>📐 标准格式</h2>
<p class='summary'>整理当前激活 Sheet 的格式：标准化单元格大小、文本对齐、边框等，让表格看起来整洁统一。</p>
<h3>使用步骤</h3>
<ol>
  <li>切换到要整理的 Sheet</li>
  <li>点击「标准格式」大按钮</li>
  <li>等待完成提示即可</li>
</ol>"
                        ),
                        new(
                            "放大镜",
                            @"
<h2>🔍 放大镜</h2>
<p class='summary'>将当前选中单元格的内容放大显示在浮窗中，方便查看长文本或小字体内容。</p>
<h3>使用步骤</h3>
<ol>
  <li>点击「放大镜」按钮开启（按钮变为「关闭」状态）</li>
  <li>选中任意单元格，浮窗自动弹出并显示放大内容</li>
  <li>再次点击按钮关闭</li>
</ol>
<div class='tip'>⚠️ 配置路径：<code>\Documents\NumDesGlobalKey.json</code>。若报错请删除该文件重新生成。</div>"
                        ),
                        new(
                            "聚光灯",
                            @"
<h2>💡 聚光灯</h2>
<p class='summary'>在屏幕上叠加十字色条（黄色横条 + 蓝色纵条），高亮当前选中单元格所在的行和列，便于跟踪视线。色条支持点击穿透，不影响正常编辑。</p>
<h3>使用步骤</h3>
<ol>
  <li>点击「聚光灯」按钮开启</li>
  <li>选择任意单元格即可看到十字色条跟随</li>
  <li>再次点击关闭</li>
</ol>
<h3>特性</h3>
<ul>
  <li>切换工作表、工作簿不影响功能</li>
  <li>无活动工作簿时色条自动隐藏</li>
  <li>插件重载后若上次开启，会自动恢复</li>
</ul>
<div class='tip'>⚠️ 配置路径：<code>\Documents\NumDesGlobalKey.json</code></div>"
                        ),
                        new(
                            "表格目录",
                            @"
<h2>📋 表格目录</h2>
<p class='summary'>在侧边栏显示当前工作簿的所有 Sheet 列表，点击名称快速跳转，省去手动翻找标签页的麻烦。</p>
<h3>使用步骤</h3>
<ol>
  <li>点击「表格目录」按钮开启侧边栏</li>
  <li>点击 Sheet 名称直接跳转</li>
  <li>再次点击按钮关闭</li>
</ol>
<div class='tip'>⚠️ 配置路径：<code>\Documents\NumDesGlobalKey.json</code></div>"
                        ),
                        new(
                            "高亮单元格",
                            @"
<h2>🌟 高亮单元格</h2>
<p class='summary'>在一定范围内高亮显示与当前选中单元格内容相同的所有单元格，方便快速定位重复值。</p>
<h3>使用步骤</h3>
<ol>
  <li>点击「高亮单元格」按钮开启</li>
  <li>选中任意单元格，范围内同值单元格自动高亮</li>
  <li>再次点击关闭高亮</li>
</ol>
<div class='tip'>⚠️ 配置路径：<code>\Documents\NumDesGlobalKey.json</code></div>"
                        ),
                        new(
                            "数据自检",
                            @"
<h2>✅ 数据自检开关</h2>
<p class='summary'>开启后，工作簿关闭时自动触发数据合法性检查，检测重复 ID、类型错误、缺失引用等问题并输出日志。</p>
<h3>使用步骤</h3>
<ol>
  <li>点击「数据自检」按钮开启</li>
  <li>正常操作表格，关闭工作簿时自动检查</li>
  <li>再次点击关闭自检</li>
</ol>
<div class='tip'>⚠️ 配置路径：<code>\Documents\NumDesGlobalKey.json</code></div>"
                        ),
                        new(
                            "公式检查",
                            @"
<h2>🔗 公式检查</h2>
<p class='summary'>检查当前工作簿所有 Sheet 中的公式，找出错误的外部连接或损坏的引用。推荐在合完表后执行。</p>
<h3>使用步骤</h3>
<ol>
  <li>打开要检查的工作簿</li>
  <li>点击「检查工具」→「公式检查」</li>
  <li>查看弹出结果列表，逐一处理问题公式</li>
</ol>
<div class='tip'>💡 仅对 <code>\Public\Excels\Tables\</code> 路径下的配置表生效；文件名或 Sheet 名含 <code>#</code> 的跳过检测。</div>"
                        ),
                        new(
                            "检查隐藏（当前）",
                            @"
<h2>🔎 检查隐藏（当前）</h2>
<p class='summary'>检查当前变动 Excel 文件是否含有隐藏行/列/Sheet。使用 VSTO 方式，需要打开 Excel，速度较慢。</p>
<h3>使用步骤</h3>
<ol>
  <li>确保目标文件已在 Excel 中打开</li>
  <li>点击「检查工具」→「检查隐藏(当前)」</li>
  <li>查看输出结果</li>
</ol>"
                        ),
                        new(
                            "检查隐藏（全局）",
                            @"
<h2>🔎 检查隐藏（全局）</h2>
<p class='summary'>批量检查全量 Excel 文件是否含有隐藏内容。VSTO 方式，逐文件打开，速度较慢，适合定期巡检。</p>
<h3>使用步骤</h3>
<ol>
  <li>点击「检查工具」→「检查隐藏(全局)」</li>
  <li>等待批量检测完成</li>
  <li>查看汇总报告</li>
</ol>"
                        ),
                        new(
                            "检查表格式",
                            @"
<h2>📊 检查表格式</h2>
<p class='summary'>检查全量 Excel 文件是否存在重复 Key、单元格格式不合法等问题，保障配置表规范性。</p>
<h3>使用步骤</h3>
<ol>
  <li>点击「检查工具」→「检查格式」</li>
  <li>等待扫描完成</li>
  <li>查看错误报告，修复问题单元格</li>
</ol>"
                        ),
                        new(
                            "更新表格路径",
                            @"
<h2>🔄 更新表格路径</h2>
<p class='summary'>更新工作簿中 Power Query 表连接的路径，用于换机器或迁移目录后修复数据源连接。</p>
<h3>使用步骤</h3>
<ol>
  <li>点击「检查工具」→「更新表格路径」</li>
  <li>按提示输入新路径</li>
  <li>等待所有查询路径更新完成</li>
</ol>"
                        ),
                        new(
                            "导出 Lua（当前）",
                            @"
<h2>📤 导出 Lua（当前）</h2>
<p class='summary'>将当前激活工作表的数据导出为 Lua 表格文件，供游戏客户端/服务端直接使用。</p>
<h3>使用步骤</h3>
<ol>
  <li>切换到要导出的 Sheet</li>
  <li>点击「导出Lua」→「Lua表格(当前)」</li>
  <li>选择输出目录，等待完成提示</li>
</ol>"
                        ),
                        new(
                            "导出 Lua（全量）",
                            @"
<h2>📤 导出 Lua（全量）</h2>
<p class='summary'>批量将所有配置表数据导出为 Lua 文件，一次性更新全部客户端/服务端数据。</p>
<h3>使用步骤</h3>
<ol>
  <li>点击「导出Lua」→「Lua表格(全量)」</li>
  <li>等待批量导出完成</li>
  <li>检查输出目录中的 .lua 文件</li>
</ol>"
                        ),
                    }
                ),
                new(
                    "自动填表",
                    new()
                    {
                        new(
                            "自动数据LTE",
                            @"
<h2>⚙️ 自动数据LTE</h2>
<p class='summary'>根据模板数据行，自动批量覆写 N 条类似数据到目标 xlsx（EPPlus 实现，单线程）。适合 LTE 类配置表的批量写入。</p>
<h3>使用步骤</h3>
<ol>
  <li>在配置表中选中要作为模板的行</li>
  <li>点击「自动数据」→「自动数据LTE」</li>
  <li>等待写入完成，查看日志</li>
</ol>
<div class='tip'>💡 目标 xlsx 须在配置的 Tables 目录下；含公式的格不会被覆盖。</div>"
                        ),
                        new(
                            "自动数据LTE（多线程）",
                            @"
<h2>⚙️ 自动数据LTE（多线程）</h2>
<p class='summary'>与「自动数据LTE」功能相同，但使用多线程并发处理，适合数据量大时加速写入。</p>
<h3>使用步骤</h3>
<ol>
  <li>在配置表中选中模板行</li>
  <li>点击「自动数据」→「自动数据LTE(多线程)」</li>
  <li>等待写入完成，查看日志</li>
</ol>"
                        ),
                        new(
                            "自动数据LTE（New）",
                            @"
<h2>⚙️ 自动数据LTE（New）</h2>
<p class='summary'>新版自动数据写入，逻辑与 LTE 版本类似，但采用了更新的写入机制，修复了部分旧版问题。</p>"
                        ),
                        new(
                            "自动数据LTE（多线程 New）",
                            @"
<h2>⚙️ 自动数据LTE（多线程 New）</h2>
<p class='summary'>新版多线程自动数据写入，结合了 New 版机制与多线程加速。</p>"
                        ),
                        new(
                            "特殊写入",
                            @"
<h2>✏️ 特殊写入</h2>
<p class='summary'>修正已存在表格的指定值（如 LTE 皮肤 xxx【模版】），不支持自增或批量替换，适合定点修正场景。</p>
<h3>使用步骤</h3>
<ol>
  <li>打开目标配置表</li>
  <li>点击「自动数据」→「特殊写入」</li>
  <li>按提示指定修正的 key 和目标值</li>
</ol>"
                        ),
                        new(
                            "自动数据（对话类）",
                            @"
<h2>💬 自动数据（对话类）</h2>
<p class='summary'>专门用于对话类数据表的自动填写，按模板覆写对话内容配置。</p>
<h3>使用步骤</h3>
<ol>
  <li>在对话配置表中准备好模板行</li>
  <li>点击「自动数据(对话类)」</li>
  <li>等待写入完成</li>
</ol>"
                        ),
                        new(
                            "合并Excel (Alice-Cove)",
                            @"
<h2>🔀 合并Excel (Alice-Cove)</h2>
<p class='summary'>将 Alice 与 Cove 两个工程的表格数据互相拷贝，实现双向同步合并。</p>
<h3>使用步骤</h3>
<ol>
  <li>在「文件信息」组配置好源表和目标路径</li>
  <li>点击「合并Excel(Alice-Cove)」</li>
  <li>等待合并完成，检查输出</li>
</ol>"
                        ),
                        new(
                            "查验模版写入数据",
                            @"
<h2>🧪 查验模版写入数据</h2>
<p class='summary'>预览模板表写入后的结果数据，用于验证写入逻辑是否正确，不实际写入文件。</p>
<h3>使用步骤</h3>
<ol>
  <li>点击「查验模版写入数据」</li>
  <li>选择模板表和目标配置</li>
  <li>查看预览输出结果</li>
</ol>"
                        ),
                        new(
                            "活动奖励写入",
                            @"
<h2>🎁 活动奖励写入</h2>
<p class='summary'>针对各类活动奖励内容反复修改的自动化填写工具，防止漏填数据。</p>
<h3>使用步骤</h3>
<ol>
  <li>准备好奖励配置数据</li>
  <li>点击「活动奖励写入」</li>
  <li>按向导填写或导入奖励内容</li>
</ol>"
                        ),
                        new(
                            "图片Fix",
                            @"
<h2>🖼️ 图片Fix</h2>
<p class='summary'>针对 Icon.xlsx 表格的图片资源修正数据同步。先复制目标内容，再点击按钮执行。</p>
<h3>使用步骤</h3>
<ol>
  <li>复制要修正的图片资源数据（Ctrl+C）</li>
  <li>点击「图片Fix」按钮执行同步</li>
</ol>"
                        ),
                    }
                ),
                new(
                    "Excel搜索",
                    new()
                    {
                        new(
                            "全局搜索",
                            @"
<h2>🔍 全局搜索</h2>
<p class='summary'>在搜索框输入关键词后，全列遍历所有 Excel 文件（按 A-Z 顺序），输出包含该关键词的结果。</p>
<h3>使用步骤</h3>
<ol>
  <li>在 Ribbon 搜索框中输入搜索编号/关键字</li>
  <li>点击「全局搜索」→「全局搜索」</li>
  <li>等待结果输出到结果表</li>
</ol>
<div class='tip'>💡 模糊搜索：关键词前加 <code>*</code>，如 <code>*皮肤</code></div>"
                        ),
                        new(
                            "全局搜索（多线程）",
                            @"
<h2>🔍 全局搜索（多线程）</h2>
<p class='summary'>与「全局搜索」功能相同，使用多线程加速，适合文件数量多时使用。</p>"
                        ),
                        new(
                            "编号搜索（多线程）",
                            @"
<h2>🔢 编号搜索（多线程）</h2>
<p class='summary'>专门搜索第二列（编号列），多线程遍历文件，比全列搜索更精准快速。</p>
<h3>使用步骤</h3>
<ol>
  <li>在搜索框输入编号</li>
  <li>点击「全局搜索」→「编号搜索(多线程)」</li>
  <li>查看结果</li>
</ol>"
                        ),
                        new(
                            "合并数据-关键词",
                            @"
<h2>🔀 合并数据-关键词</h2>
<p class='summary'>以搜索框当前关键词为条件，执行 Alice-Cove 合并操作，只处理匹配关键词的数据行。</p>"
                        ),
                        new(
                            "常规模版（创建数据）",
                            @"
<h2>📄 常规模版</h2>
<p class='summary'>全局搜索关键词，找到所有匹配行后，按常规模版格式（如各类 LTE、主岛活动【模版】）输出表格数据。</p>
<h3>使用步骤</h3>
<ol>
  <li>在搜索框输入关键词</li>
  <li>点击「创建数据」→「常规模版」</li>
  <li>查看输出的模版数据表</li>
</ol>"
                        ),
                        new(
                            "特殊模版（创建数据）",
                            @"
<h2>📄 特殊模版</h2>
<p class='summary'>指定表格搜索关键词，按特殊模版格式（如 LTE 皮肤【模版】）输出表格数据。</p>"
                        ),
                        new(
                            "替换文本（当前表格）",
                            @"
<h2>🔁 替换文本</h2>
<p class='summary'>在当前激活的 Sheet 中搜索特定字符串并替换，比 Excel 原生替换支持更多规则。</p>
<h3>使用步骤</h3>
<ol>
  <li>在搜索框输入要查找的文本</li>
  <li>点击「当前表操作」→「替换文本」</li>
  <li>按提示输入替换内容并执行</li>
</ol>"
                        ),
                        new(
                            "搜索文本（当前表格）",
                            @"
<h2>🔍 搜索文本</h2>
<p class='summary'>在当前激活的 Sheet 中搜索特定字符串，高亮显示所有匹配位置。</p>"
                        ),
                        new(
                            "搜索Sheet名",
                            @"
<h2>📑 搜索Sheet名</h2>
<p class='summary'>搜索 txt 导出文件所在的工作簿，特别针对以 <code>$</code> 开头的 Sheet 名进行定位。</p>"
                        ),
                        new(
                            "搜索公式名",
                            @"
<h2>📐 搜索公式名</h2>
<p class='summary'>搜索当前 Excel 文件中所有使用的公式，列出公式名和位置，便于公式审查。</p>"
                        ),
                        new(
                            "Excel数据DB化",
                            @"
<h2>🗄️ Excel数据DB化</h2>
<p class='summary'>将 Public 目录下的配置 Excel 文件备份并数据库化（SQLite），提供更高效的查询能力，替代逐文件遍历搜索。</p>
<h3>使用步骤</h3>
<ol>
  <li>点击「数据DB化」</li>
  <li>等待扫描并导入数据库完成</li>
  <li>后续全局搜索自动优先使用数据库</li>
</ol>"
                        ),
                    }
                ),
                new(
                    "运营工具",
                    new()
                    {
                        new(
                            "生成活动（活动名）",
                            @"
<h2>🎉 生成活动（活动名）</h2>
<p class='summary'>根据运营排期表，以活动名为索引，自动生成对应的活动配置数据，填写到相关配置表。</p>
<h3>使用步骤</h3>
<ol>
  <li>确保运营排期表已配置完毕</li>
  <li>点击「生成活动」→「生成活动(活动名)」</li>
  <li>输入或选择活动名</li>
  <li>等待自动生成完成</li>
</ol>"
                        ),
                        new(
                            "生成活动（活动ID）",
                            @"
<h2>🎉 生成活动（活动ID）</h2>
<p class='summary'>与「生成活动(活动名)」类似，改为以活动 ID 为索引进行生成，适合已知 ID 的场景。</p>"
                        ),
                        new(
                            "更新活动",
                            @"
<h2>🔄 更新活动</h2>
<p class='summary'>根据运营排期表，自动更新已存在的活动配置数据（修改而非新建）。</p>
<h3>使用步骤</h3>
<ol>
  <li>确保排期表中的活动信息已更新</li>
  <li>点击「生成活动」→「更新活动」</li>
  <li>等待更新完成</li>
</ol>"
                        ),
                        new(
                            "对比Excel",
                            @"
<h2>📊 对比Excel</h2>
<p class='summary'>对比不同版本、同路径下的非 <code>#</code> 开头 Excel 文件，找出差异行/列，输出至 <code>【文档\#表格比对结果.xlsx】</code>。</p>
<h3>使用步骤</h3>
<ol>
  <li>确保两个版本的文件位于相同路径结构下</li>
  <li>点击「对比Excel」</li>
  <li>等待对比完成，查看输出文件</li>
</ol>"
                        ),
                        new(
                            "解决Git冲突（xlsx）",
                            @"
<h2>⚔️ 解决Git冲突（xlsx）</h2>
<p class='summary'>自动检测 Git 工作区中冲突的 xlsx 文件，以可视化界面逐格选择保留「我的」或「他的」版本，支持 git add 写回。</p>
<h3>使用步骤</h3>
<ol>
  <li>执行 git merge/rebase 产生 xlsx 冲突后</li>
  <li>点击「xlsx冲突解决」→「解决Git冲突」</li>
  <li>在冲突窗口逐行/逐列选择保留版本</li>
  <li>点击「应用并git add」写回并标记已解决</li>
</ol>
<div class='tip'>💡 带 <code>#</code> 的工作簿/Sheet 为辅助表，不参与对比。窗口默认最大化打开。</div>"
                        ),
                        new(
                            "手动两文件对比",
                            @"
<h2>📑 手动两文件对比</h2>
<p class='summary'>手动选择两个 xlsx 文件，以可视化界面逐格对比差异并选择合并结果，不依赖 Git。</p>
<h3>使用步骤</h3>
<ol>
  <li>点击「xlsx冲突解决」→「手动两文件对比」</li>
  <li>选择「我的」xlsx 和「他的」xlsx</li>
  <li>在冲突窗口逐行逐列处理差异</li>
  <li>保存合并结果</li>
</ol>"
                        ),
                        new(
                            "查看历史版本对比",
                            @"
<h2>📜 查看历史版本对比</h2>
<p class='summary'>浏览指定 xlsx 的 Git 提交历史，选择任意历史版本与工作区当前版本或另一历史版本进行对比。</p>
<h3>使用步骤</h3>
<ol>
  <li>点击「xlsx冲突解决」→「查看历史版本对比」</li>
  <li>选择目标 xlsx 文件（须在 Git 仓库内）</li>
  <li>选择历史版本</li>
  <li>选择对比模式（vs 工作区 或 vs 另一历史）</li>
  <li>在冲突窗口解决差异</li>
</ol>"
                        ),
                        new(
                            "溯源改动",
                            @"
<h2>🔍 溯源改动</h2>
<p class='summary'>根据对比结果和 <code>【文档\表格关系.json】</code> 溯源每条改动最终影响的表格（一般为 ActivityClientData.xlsx），输出至 <code>【文档\#溯源结果.xlsx】</code>。</p>
<h3>使用步骤</h3>
<ol>
  <li>先执行「对比Excel」获取对比结果</li>
  <li>确保 <code>表格关系.json</code> 已维护</li>
  <li>点击「溯源改动」</li>
  <li>查看溯源结果文件</li>
</ol>"
                        ),
                        new(
                            "检查数据格式",
                            @"
<h2>🔧 检查数据格式</h2>
<p class='summary'>手动触发数据格式检查，自动取消隐藏行/列，过滤规则配置在 <code>\Document\NumDesToolsConfig.json</code>。</p>"
                        ),
                        new(
                            "克隆活动",
                            @"
<h2>🖨️ 克隆活动</h2>
<p class='summary'>根据 <code>ActivityTableRules</code> 规则，以指定源活动 ID 为模板，将新期次数据克隆到所有相关配置表。</p>
<h3>使用步骤</h3>
<ol>
  <li>确保 ActivityTableRules.json 已配置好规则</li>
  <li>点击「克隆活动」</li>
  <li>输入源活动 ID 和新期次 ID</li>
  <li>等待克隆完成，检查各表数据</li>
</ol>"
                        ),
                        new(
                            "更新活动规则",
                            @"
<h2>📋 更新活动规则</h2>
<p class='summary'>扫描 EnumCmds / ActivityManager / LogicBase 代码，自动补全 <code>ActivityTableRules.json</code> 中的 typeTableMap（只增不删改）。</p>
<h3>使用步骤</h3>
<ol>
  <li>点击「更新活动规则」</li>
  <li>等待扫描完成</li>
  <li>检查 ActivityTableRules.json 新增条目</li>
</ol>"
                        ),
                        new(
                            "验证活动（全量）",
                            @"
<h2>✅ 验证活动（全量）</h2>
<p class='summary'>验证所有活动配置表的跨表引用和字段规则合法性，输出完整的验证报告。</p>
<h3>使用步骤</h3>
<ol>
  <li>点击「验证活动」→「验证活动(全量)」</li>
  <li>等待扫描所有活动表完成</li>
  <li>查看验证日志，处理错误项</li>
</ol>"
                        ),
                        new(
                            "验证活动（指定ID）",
                            @"
<h2>✅ 验证活动（指定ID）</h2>
<p class='summary'>输入指定活动 ID，只验证该活动的配置链完整性，比全量验证更快。</p>"
                        ),
                        new(
                            "验证活动（Git改动）",
                            @"
<h2>✅ 验证活动（Git改动）</h2>
<p class='summary'>只验证当前 Git 工作区有改动的活动相关配置表，精准覆盖本次修改影响范围。</p>"
                        ),
                    }
                ),
                new(
                    "文件信息",
                    new()
                    {
                        new(
                            "文件名称",
                            @"
<h2>📋 文件名称</h2>
<p class='summary'>复制当前激活工作簿的文件名（不含路径）到剪贴板，方便填写配置或写文档时引用。</p>"
                        ),
                        new(
                            "文件路径",
                            @"
<h2>📁 文件路径</h2>
<p class='summary'>复制当前激活工作簿的完整文件路径到剪贴板。</p>"
                        ),
                        new(
                            "源表/目标路径框",
                            @"
<h2>🗂️ 源表 / 目标路径</h2>
<p class='summary'>两个可编辑文本框，分别用于配置「源表根目录」和「目标根目录」。各功能（合并Excel、全局搜索等）会读取这两个路径。输入后自动保存到本地配置。</p>
<h3>使用步骤</h3>
<ol>
  <li>在「源表/目标」标签下方找到两个输入框</li>
  <li>分别输入本地的源表根目录和目标根目录</li>
  <li>输入后自动保存，下次打开 Excel 无需重新输入</li>
</ol>"
                        ),
                    }
                ),
                new(
                    "玩法计算",
                    new()
                    {
                        new(
                            "Alice大富翁",
                            @"
<h2>🎲 Alice大富翁</h2>
<p class='summary'>大富翁玩法方案整理工具，用于计算和整理大富翁活动配置数据。</p>"
                        ),
                        new(
                            "TM目标元素",
                            @"
<h2>🎯 TM目标元素</h2>
<p class='summary'>生成 TM（Target Map）目标元素配置数据。</p>"
                        ),
                        new(
                            "TM非目标元素",
                            @"
<h2>🎯 TM非目标元素</h2>
<p class='summary'>生成 TM 非目标元素配置数据。</p>"
                        ),
                        new(
                            "移动魔瓶",
                            @"
<h2>🧪 移动魔瓶</h2>
<p class='summary'>移动魔瓶玩法消耗模拟计算器，输入参数后输出消耗曲线和数值方案。</p>"
                        ),
                        new(
                            "移动转盘",
                            @"
<h2>🎰 移动转盘</h2>
<p class='summary'>移动转盘玩法随机方案机选工具，模拟转盘概率分布并输出方案建议。</p>"
                        ),
                        new(
                            "万能卡概率",
                            @"
<h2>🃏 万能卡概率</h2>
<p class='summary'>相册万能卡概率模拟器，输入卡池配置后模拟概率分布，辅助设计概率方案。</p>"
                        ),
                        new(
                            "删除冗余列",
                            @"
<h2>🗑️ 删除冗余列</h2>
<p class='summary'>批量删除表格中的冗余列，同时清理 Chart1 类的非法配置表。打开的表格不能有格式错误。</p>
<h3>使用步骤</h3>
<ol>
  <li>确保要处理的表格格式正确（无合并错误等）</li>
  <li>点击「删除冗余列」</li>
  <li>按提示选择或确认要处理的文件</li>
  <li>等待处理完成</li>
</ol>"
                        ),
                    }
                ),
                new(
                    "杂项 & AI",
                    new()
                    {
                        new(
                            "插件日志",
                            @"
<h2>📝 插件日志</h2>
<p class='summary'>展示插件运行日志窗口，查看各功能的执行记录、错误信息等，便于排查问题。</p>
<h3>使用步骤</h3>
<ol>
  <li>点击「插件日志」开启日志窗口</li>
  <li>执行其他功能时日志实时滚动</li>
  <li>再次点击关闭日志窗口</li>
</ol>"
                        ),
                        new(
                            "默认配置",
                            @"
<h2>🔄 默认配置</h2>
<p class='summary'>将插件所有用户配置恢复为默认值，适合配置文件损坏或想重置设置时使用。</p>
<div class='tip'>⚠️ 执行后当前所有自定义配置（路径、开关状态等）将被清除，请谨慎操作。</div>"
                        ),
                        new(
                            "AI对话",
                            @"
<h2>🤖 AI对话</h2>
<p class='summary'>内置 AI 对话面板，可直接在 Excel 插件中与 AI 进行交互，辅助数据分析、内容生成等工作。</p>
<h3>使用步骤</h3>
<ol>
  <li>在「AI配置选择」下拉框选择模型（ChatGPT-4o 或 DeepSeek-V3）</li>
  <li>点击「AI对话」开启对话面板</li>
  <li>在输入框中输入问题，回车发送</li>
</ol>
<div class='tip'>💡 使用前确保已在配置文件中填写对应的 API Key。</div>"
                        ),
                    }
                ),
            };

        // ── CSS ───────────────────────────────────────────────────────────────────

        private const string Css =
            @"<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
    font-family: '微软雅黑', 'Segoe UI', sans-serif;
    background: #1e1e1e;
    color: #d4d4d4;
    padding: 28px 36px;
    font-size: 20px;
    line-height: 1.8;
}
h2 {
    font-size: 30px;
    color: #ffffff;
    margin-bottom: 12px;
    padding-bottom: 10px;
    border-bottom: 2px solid #3a3a5a;
}
h3 {
    font-size: 20px;
    color: #9cdcfe;
    margin: 22px 0 10px;
    letter-spacing: 0.05em;
}
p.summary {
    color: #b5b5b5;
    font-size: 20px;
    margin-bottom: 12px;
    line-height: 1.8;
}
ol, ul {
    padding-left: 28px;
    margin-bottom: 14px;
}
li {
    margin-bottom: 7px;
    color: #cccccc;
    font-size: 20px;
}
code {
    background: #2d2d2d;
    color: #ce9178;
    padding: 2px 7px;
    border-radius: 3px;
    font-family: Consolas, monospace;
    font-size: 18px;
}
.tip {
    margin-top: 20px;
    padding: 12px 18px;
    background: #252540;
    border-left: 4px solid #5a5aaa;
    border-radius: 4px;
    color: #aaaacc;
    font-size: 18px;
}
#placeholder {
    display: flex;
    align-items: center;
    justify-content: center;
    height: 80vh;
    color: #555;
    font-size: 20px;
}
</style>";

        private const string PlaceholderHtml =
            "<!DOCTYPE html><html><head><meta charset='utf-8'>"
            + Css
            + "</head><body>"
            + "<div id='placeholder'>← 从左侧选择功能查看说明</div>"
            + "</body></html>";

        // ── 构造 ─────────────────────────────────────────────────────────────────

        public HelpWindow()
        {
            WpfUiHelper.EnsureApplication();
            InitializeComponent();
            Loaded += (_, _) => WpfUiHelper.ApplyDarkTitleBar(this);
            BuildNavTree();
            HelpBrowser.NavigateToString(PlaceholderHtml);
        }

        // ── 构建左侧导航树 ────────────────────────────────────────────────────────

        private void BuildNavTree()
        {
            foreach (var group in HelpData)
            {
                var groupNode = new TreeViewItem
                {
                    Header = group.Name,
                    Style = (System.Windows.Style)FindResource("GroupItem"),
                    IsExpanded = true
                };

                foreach (var item in group.Items)
                {
                    var leafNode = new TreeViewItem
                    {
                        Header = item.Title,
                        Tag = item,
                        Style = (System.Windows.Style)FindResource("LeafItem")
                    };
                    groupNode.Items.Add(leafNode);
                }

                NavTreeView.Items.Add(groupNode);
            }
        }

        // ── 选中事件 ──────────────────────────────────────────────────────────────

        private void NavTreeView_SelectedItemChanged(
            object sender,
            System.Windows.RoutedPropertyChangedEventArgs<object> e
        )
        {
            if (e.NewValue is not TreeViewItem { Tag: HelpItem item })
                return;

            var html =
                "<!DOCTYPE html><html><head><meta charset='utf-8'>"
                + Css
                + "</head><body>"
                + item.Html
                + "</body></html>";
            HelpBrowser.NavigateToString(html);
        }

        private void Window_EscClose(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Escape)
                Close();
        }
    }
}
