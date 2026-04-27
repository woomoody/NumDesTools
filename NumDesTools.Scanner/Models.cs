using Newtonsoft.Json;

namespace NumDesTools.Scanner;

// ── ActivityTableRules.json 模型 ──────────────────────────────────────────────

public class ActivityTableRules
{
    [JsonProperty("tables")]
    public List<TableDef> Tables { get; set; } = [];

    [JsonProperty("typeSubTableRules")]
    public Dictionary<string, SubTableRule> TypeSubTableRules { get; set; } = [];

    [JsonProperty("typeTableMap")]
    public Dictionary<string, string> TypeTableMap { get; set; } = [];

    [JsonProperty("typeMultiSubTableRules")]
    public Dictionary<string, List<string>> TypeMultiSubTableRules { get; set; } = [];
}

public class TableDef
{
    [JsonProperty("name")]       public string Name      { get; set; } = "";
    [JsonProperty("luaKey")]     public string LuaKey    { get; set; } = "";
    [JsonProperty("excelFile")]  public string ExcelFile { get; set; } = "";
    [JsonProperty("desc")]       public string Desc      { get; set; } = "";
    [JsonProperty("keyField")]   public string KeyField  { get; set; } = "id";
    [JsonProperty("fields")]     public List<FieldDef> Fields { get; set; } = [];
}

public class FieldDef
{
    [JsonProperty("name")]       public string Name       { get; set; } = "";
    [JsonProperty("required")]   public bool   Required   { get; set; }
    [JsonProperty("type")]       public string Type       { get; set; } = "";
    [JsonProperty("refTable")]   public string RefTable   { get; set; } = "";
    [JsonProperty("refIsArray")] public bool   RefIsArray { get; set; }
    [JsonProperty("desc")]       public string Desc       { get; set; } = "";
}

public class SubTableRule
{
    [JsonProperty("table")]       public string? Table       { get; set; }
    [JsonProperty("lookupField")] public string  LookupField { get; set; } = "activityID";
    [JsonProperty("fields")]      public List<FieldDef> Fields { get; set; } = [];
}

// ── 工作项（需求 / 缺陷）────────────────────────────────────────────────────

public record WorkItem(string Id, string Name, string Desc, string Status);

// ── 识别到的配置表信息 ────────────────────────────────────────────────────────

public record TableMatch(
    string Excel,
    string Desc,
    List<string> RequiredFields,
    string LookupField,
    string? TypeNum   // null 表示主表
);

// ── 待写入飞书的评论条目 ──────────────────────────────────────────────────────

public class PendingItem
{
    [JsonProperty("id")]           public string Id          { get; set; } = "";
    [JsonProperty("name")]         public string Name        { get; set; } = "";
    [JsonProperty("tables")]       public List<string> Tables { get; set; } = [];
    [JsonProperty("comment")]      public string Comment     { get; set; } = "";
    [JsonProperty("item_type")]    public string ItemType    { get; set; } = "story"; // story | issue
    [JsonProperty("skip_comment")]   public bool   SkipComment   { get; set; }
    [JsonProperty("skip_reason")]    public string SkipReason   { get; set; } = "";
    [JsonProperty("original_desc")]  public string OriginalDesc { get; set; } = ""; // 人工原有正文，用于追加写入
}

// ── 知识库 ────────────────────────────────────────────────────────────────────

public class KnowledgeBase
{
    // key: "type_2", "type_126", "type_unknown"
    [JsonExtensionData]
    public Dictionary<string, Newtonsoft.Json.Linq.JToken> Types { get; set; } = [];

    public List<KnowledgeEntry> GetEntries(string typeKey)
    {
        if (!Types.TryGetValue(typeKey, out var token)) return [];
        try { return token.ToObject<List<KnowledgeEntry>>() ?? []; }
        catch { return []; }
    }

    public void SetEntries(string typeKey, List<KnowledgeEntry> entries)
        => Types[typeKey] = Newtonsoft.Json.Linq.JToken.FromObject(entries);
}

public class KnowledgeEntry
{
    [JsonProperty("source_id")]    public string SourceId   { get; set; } = "";
    [JsonProperty("source_name")]  public string SourceName { get; set; } = "";
    [JsonProperty("human_notes")]  public List<string> HumanNotes { get; set; } = [];
    [JsonProperty("learned_at")]   public string LearnedAt  { get; set; } = "";
    [JsonProperty("updated_at")]   public string UpdatedAt  { get; set; } = "";
}

// ── 已回读状态 ────────────────────────────────────────────────────────────────

// key: work_item_id → 已处理的 comment_id 集合
public class ReviewedState : Dictionary<string, List<string>> { }
