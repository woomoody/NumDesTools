using System.Text;
using System.Text.RegularExpressions;

namespace NumDesTools.Scanner;

public static class CommentBuilder
{
    public const string StoryMarker = "[AI分析]";
    public const string IssueMarker = "[AI缺陷分析]";

    // ── 需求分析 ─────────────────────────────────────────────────────────────

    public static string BuildStoryComment(
        WorkItem story, List<TableMatch> tables, int phase, List<string>? historyNotes = null)
    {
        var sb = new StringBuilder();
        bool isFollowup = phase > 1;

        sb.AppendLine($"{StoryMarker} 配置表分析");
        sb.AppendLine(isFollowup
            ? $"第{phase}期续期，建议用【克隆活动】工具，重点核查："
            : "首期/全新活动，完整配置：");
        sb.AppendLine();

        for (int i = 0; i < tables.Count; i++)
        {
            var t = tables[i];
            if (isFollowup)
            {
                sb.AppendLine($"{i + 1}. {t.Excel}");
            }
            else
            {
                sb.Append($"{i + 1}. {t.Excel}");
                if (!string.IsNullOrEmpty(t.Desc) && !t.Desc.StartsWith("type="))
                    sb.Append($" — {t.Desc}");
                if (t.RequiredFields.Count > 0)
                    sb.Append($"（必填：{string.Join(", ", t.RequiredFields)}）");
                sb.AppendLine();
            }
        }

        if (historyNotes?.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("历史经验：");
            foreach (var note in historyNotes)
                sb.AppendLine($"• {note}");
        }

        return sb.ToString().TrimEnd();
    }

    // ── 缺陷分析 ─────────────────────────────────────────────────────────────

    private static readonly (string[] Keywords, string[] HintExcels, string Fix)[] BugKeywordHints =
    [
        (["寻找不到", "找不到目标", "寻不到", "寻找 寻了", "寻找有报错", "寻找加一下",
          "先寻", "寻找顺序", "寻找优先级", "寻找配置"],
         ["FindTargetTemplateData.xlsx", "ItemExchangeMat.xlsx", "LteExchangeLandmarkData.xlsx"],
         "FindTargetTemplateData.findTargets（定义寻什么/顺序）；ItemExchangeMat.findTargetsId；LteExchangeLandmarkData.exchange_find_data"),

        (["兑换点", "兑换", "icon不同", "资源和icon", "详情页报错", "帮助界面报错",
          "详情页有报错", "气泡"],
         ["LteExchangeLandmark.xlsx", "LteExchangeLandmarkData.xlsx", "ItemExchange.xlsx", "ItemExchangeMat.xlsx"],
         "数据链：LteExchangeLandmark→LteExchangeLandmarkData.exchange_need/reward_data→ItemExchangeMat；详情页报错优先查引用是否存在；icon检查ItemExchangeMat→Item.iconId"),

        (["地组", "锁头", "点击锁头", "地组解锁", "地组锁", "解锁前", "锁头寻找", "地组未解锁"],
         ["LteExchangeLandmark.xlsx", "MapDataProto.xlsx"],
         "解锁逻辑：LteExchangeLandmark.unlock_item；点击报错查unlock_item引用的道具ID是否在Item.xlsx中有效"),

        (["任务", "主链", "建造", "链条", "未建造", "建造前", "建造后", "合成",
          "地图任务", "任务不能完成", "任务流程"],
         ["LteData.xlsx", "LteTaskNPCLandmark.xlsx"],
         "入口：LteData.allTasks/LteStageTaskId；NPC地标：LteTaskNPCLandmark.task_id；建造卡死检查前置任务ID有效性"),

        (["矿", "采矿", "无限矿", "矿产出", "矿的白色底板"],
         ["LteExchangeLandmarkData.xlsx", "MapDataProto.xlsx"],
         "矿产出：LteExchangeLandmarkData.exchange_find_data→FindTargetTemplateData；无限矿地标在MapDataProto"),

        (["炸弹", "被炸障碍物", "障碍物", "爆炸"],
         ["LteExchangeLandmark.xlsx", "FindTargetTemplateData.xlsx"],
         "障碍物寻找：FindTargetTemplateData；点击报错查LteExchangeLandmark.repair_id/exchange_data"),

        (["图鉴", "模型不同", "道具", "icon", "图标"],
         ["Item.xlsx"],
         "Item.xlsx：iconId/atlasId/modelId；模型不同通常是modelId与关卡assetName不一致"),

        (["奖励", "阶段奖励", "进度条", "领取", "宝箱"],
         ["ActivityStageRewardData.xlsx"],
         "ActivityStageRewardData：stageCount/threshold/itemId；进度条上限查maxStage"),

        (["奖励区", "宝箱", "collect"],
         ["LteFullCollectData.xlsx"],
         "LteFullCollectData.findTargetID/redPointID"),

        (["红点", "入口红点", "入口无红点", "red_point", "redpoint"],
         ["ActivityClientData.xlsx", "LteFullCollectData.xlsx"],
         "入口红点：ActivityClientData；内部红点：LteFullCollectData.redPointID"),

        (["报错", "崩溃", "错误", "初始化", "进入lte", "进入有报错", "进图有报错"],
         ["ActivityClientData.xlsx", "LteData.xlsx"],
         "优先查ActivityClientData.followUpActivityId/subActivityId；LteData首图mapData/mapIdList"),

        (["UI", "界面", "布局", "穿帮", "黑边", "适配", "相机视角"],
         [],
         "非配置表问题（美术/前端）；ActivityClientData中UI模板ID与同类活动一致；黑边查MapDataProto.tiled_seting"),

        (["礼包", "触发礼包", "付费", "购买"],
         ["ActivityStageRewardData.xlsx"],
         "礼包触发：ActivityStageRewardData关联道具ID；icon不符查Item.xlsx.iconId"),

        (["多语言", "文案", "翻译"],
         [],
         "确认所有文案key均已添加到多语言表"),
    ];

    public static string BuildIssueComment(
        WorkItem issue, List<TableMatch> tables, string activityId,
        int? typeNum, string typeNote, List<string>? historyNotes = null,
        string? findChainAnalysis = null)
    {
        var (locationExcels, fixSuggestion, _) = InferBugHints(issue.Name);
        var sb = new StringBuilder();

        sb.AppendLine($"{IssueMarker} 缺陷配置分析");
        if (!string.IsNullOrEmpty(activityId))
        {
            sb.Append($"activityID={activityId}");
            if (typeNum.HasValue)
                sb.Append(string.IsNullOrEmpty(typeNote)
                    ? $" · type={typeNum}"
                    : $" · type={typeNum}（{typeNote}）");
            sb.AppendLine();
        }
        sb.AppendLine();

        if (locationExcels.Count > 0)
        {
            sb.AppendLine("涉及配置表：" + string.Join("、", locationExcels));
            sb.AppendLine();
        }

        sb.AppendLine("修复建议：");
        sb.AppendLine(fixSuggestion);

        if (!string.IsNullOrEmpty(findChainAnalysis))
        {
            sb.AppendLine();
            sb.AppendLine("寻找链分析：");
            sb.AppendLine(findChainAnalysis);
        }

        if (historyNotes?.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("历史经验：");
            foreach (var note in historyNotes)
                sb.AppendLine($"• {note}");
        }

        return sb.ToString().TrimEnd();
    }

    /// <summary>
    /// 返回 (命中表格列表, 修复建议, 是否命中关键词)。
    /// </summary>
    public static (List<string> Excels, string Fix, bool Matched) InferBugHints(string issueName)
    {
        var nameLower   = issueName.ToLower();
        var locations   = new List<string>();
        var suggestions = new List<string>();

        foreach (var (kws, hints, fix) in BugKeywordHints)
        {
            if (!kws.Any(kw => nameLower.Contains(kw.ToLower()))) continue;
            foreach (var h in hints)
                if (!locations.Contains(h))
                    locations.Add(h);
            suggestions.Add(fix);
        }

        bool matched = suggestions.Count > 0;
        string fixText = matched
            ? string.Join("\n", suggestions.Select(s => $"• {s}"))
            : "标题关键词未匹配到已知配置规则，请人工排查";
        return (locations, fixText, matched);
    }
}
