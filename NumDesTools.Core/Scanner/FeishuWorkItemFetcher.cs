using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;

namespace NumDesTools.Scanner;

/// <summary>
/// 从飞书拉取 story / issue 工作项，过滤已关闭状态。
/// 对应 Python 版本的 fetch_stories() / fetch_issues() / _parse_stories() / _parse_issues()。
/// </summary>
public static class FeishuWorkItemFetcher
{
    private static readonly HashSet<string> SkipStatuses =
        ["已完成", "已关闭", "废弃", "关闭", "测试通过✅"];

    private static readonly string SelectFields =
        "`work_item_id`, `name`, `description`, `work_item_status`";

    // story 额外拉需求内容字段（人工正文）
    private static readonly string SelectFieldsStory =
        "`work_item_id`, `name`, `description`, `field_d9a4cd`, `work_item_status`";

    // ── story ────────────────────────────────────────────────────────────────

    public static async Task<List<WorkItem>> FetchStoriesAsync(
        List<string>? itemIds = null, bool fetchAll = false)
    {
        if (itemIds?.Count > 0)
        {
            var ids  = string.Join(", ", itemIds);
            var mql  = $"SELECT {SelectFieldsStory} FROM `{FeishuMcpClient.ProjectKey}`.`story` WHERE `work_item_id` IN ({ids})";
            var data = await FeishuMcpClient.SearchByMqlAsync(mql);
            return ParseItems(data, storyContentKey: "field_d9a4cd");
        }

        if (fetchAll)
            return await FetchAllPaged("story");

        var single = await FeishuMcpClient.SearchByMqlAsync(
            $"SELECT {SelectFieldsStory} FROM `{FeishuMcpClient.ProjectKey}`.`story` LIMIT 50");
        return ParseItems(single, storyContentKey: "field_d9a4cd");
    }

    // ── issue ────────────────────────────────────────────────────────────────

    public static async Task<List<WorkItem>> FetchIssuesAsync(
        List<string>? itemIds = null, bool fetchAll = false)
    {
        if (itemIds?.Count > 0)
        {
            var ids  = string.Join(", ", itemIds);
            var mql  = $"SELECT {SelectFields} FROM `{FeishuMcpClient.ProjectKey}`.`issue` WHERE `work_item_id` IN ({ids})";
            var data = await FeishuMcpClient.SearchByMqlAsync(mql);
            return ParseItems(data);
        }

        if (fetchAll)
            return await FetchAllPaged("issue");

        var single = await FeishuMcpClient.SearchByMqlAsync(
            $"SELECT {SelectFields} FROM `{FeishuMcpClient.ProjectKey}`.`issue` LIMIT 50");
        return ParseItems(single);
    }

    // ── 通用 ─────────────────────────────────────────────────────────────────

    private static async Task<List<WorkItem>> FetchAllPaged(string workItemType)
    {
        var all = new List<WorkItem>();
        int offset = 0;
        bool isStory = workItemType == "story";
        var selectFields = isStory ? SelectFieldsStory : SelectFields;
        Console.WriteLine($"  正在从飞书拉取{(isStory ? "需求" : "缺陷")}数据...");
        while (true)
        {
            var mql = $"SELECT {selectFields} FROM `{FeishuMcpClient.ProjectKey}`.`{workItemType}` LIMIT 50 OFFSET {offset}";
            var data = await FeishuMcpClient.SearchByMqlAsync(mql);
            var page = ParseItems(data, isStory ? "field_d9a4cd" : null);
            if (page.Count == 0) break;
            all.AddRange(page);
            Console.WriteLine($"  已拉取 {all.Count} 条...");
            offset += 50;
        }
        Console.WriteLine($"  拉取完成，共 {all.Count} 条。\n");
        return all;
    }

    // storyContentKey: story 时传 "field_d9a4cd"，issue 时传 null 用 description
    private static List<WorkItem> ParseItems(JToken data, string? storyContentKey = null)
    {
        var result = new List<WorkItem>();
        var dataObj = data["data"] as JObject;
        if (dataObj == null) return result;

        foreach (var groupProp in dataObj.Properties())
        {
            foreach (var item in groupProp.Value)
            {
                var fields = (item["moql_field_list"] as JArray ?? [])
                    .ToDictionary(f => f["key"]?.ToString() ?? "", f => f["value"]);

                var statusList = fields.TryGetValue("work_item_status", out var sv)
                    ? sv?["key_label_value_list"] as JArray ?? []
                    : new JArray();
                var status = statusList.Count > 0 ? statusList[0]["label"]?.ToString() ?? "" : "";

                if (SkipStatuses.Contains(status)) continue;

                var id   = fields.TryGetValue("work_item_id", out var iv)
                    ? iv?["long_value"]?.ToString() ?? "" : "";
                var name = fields.TryGetValue("name", out var nv)
                    ? nv?["string_value"]?.ToString() ?? "" : "";

                // story 取需求内容字段，issue 取 description
                string desc;
                if (storyContentKey != null && fields.TryGetValue(storyContentKey, out var scv))
                    desc = scv?["string_value"]?.ToString() ?? "";
                else
                    desc = fields.TryGetValue("description", out var dv)
                        ? dv?["string_value"]?.ToString() ?? "" : "";

                if (!string.IsNullOrEmpty(id))
                    result.Add(new WorkItem(id, name, desc, status));
            }
        }
        return result;
    }

    // ── 评论拉取 ─────────────────────────────────────────────────────────────

    public static async Task<List<(string CommentId, string Content)>> GetCommentsAsync(string workItemId)
    {
        var data = await FeishuMcpClient.ListCommentsAsync(workItemId);
        var comments = data["comments"] as JArray ?? data["list"] as JArray ?? [];
        return comments
            .Select(c => (
                Id:      c["comment_id"]?.ToString() ?? "",
                Content: c["content"]?.ToString()    ?? ""))
            .Where(c => !string.IsNullOrEmpty(c.Id))
            .ToList();
    }

    /// <summary>
    /// 检查工作项是否已有 AI 分析标记（正文字段 + 评论，任一命中即跳过）。
    /// story 查 field_d9a4cd，issue 查 description；MQL 对 description 的 string_value 有时为空，
    /// 所以两个类型都实时拉一次字段值确保准确。
    /// </summary>
    public static async Task<bool> HasExistingAiCommentAsync(WorkItem item, string marker)
    {
        // 快速路径：MQL 已拉到正文内容（story 可靠，issue 有时为空）
        if (item.Desc.Contains(marker)) return true;

        // 实时拉正文字段
        var fieldKey = marker == CommentBuilder.StoryMarker
            ? FeishuMcpClient.StoryContentFieldKey
            : FeishuMcpClient.IssueDescFieldKey;
        var fieldVal = await FeishuMcpClient.GetCurrentFieldValueAsync(item.Id, fieldKey);
        if (fieldVal.Contains(marker)) return true;

        // 也检查评论（兼容历史上写过评论的单子）
        try
        {
            var comments = await GetCommentsAsync(item.Id);
            if (comments.Any(c => c.Content.Contains(marker))) return true;
        }
        catch { }

        return false;
    }

    /// <summary>同步版本，仅用 MQL 已有数据快速判断（知识回读用，不需要网络精确）。</summary>
    public static bool HasExistingAiDescription(WorkItem item, string marker)
        => item.Desc.Contains(marker);

    // ── 活动ID提取（缺陷标题用）─────────────────────────────────────────────

    private static readonly Regex RxActivityIdPure   = new(@"[【\[](\d{4,8})[】\]]", RegexOptions.Compiled);
    private static readonly Regex RxActivityIdSuffix = new(@"[【\[][^】\]]*?(\d{4,8})[】\]]", RegexOptions.Compiled);

    public static List<string> ExtractActivityIds(string title)
    {
        var ids = RxActivityIdPure.Matches(title).Select(m => m.Groups[1].Value).Distinct().ToList();
        if (ids.Count == 0)
            ids = RxActivityIdSuffix.Matches(title).Select(m => m.Groups[1].Value).Distinct().ToList();
        return ids;
    }
}
