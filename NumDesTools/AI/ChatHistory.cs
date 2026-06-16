using System.Threading.Tasks;
using Microsoft.Data.Sqlite;

namespace NumDesTools.AI;

public record ChatSession(string SessionId, DateTime LastTime, string Preview, string? Title = null);

public class ChatMessage
{
    public string Role { get; set; }
    public string Message { get; set; }
    public bool IsUser { get; set; }
    public DateTime Timestamp { get; set; }
    public bool IsStreaming { get; set; }
    public string SessionId { get; set; }
    public bool IsAgent { get; set; }
}

public class ChatHistoryManager
{
    private readonly string _connectionString;

    // 单元测试可通过设置此环境变量覆盖默认数据库路径
    internal const string TestDbEnvVar = "NUMDES_CHATHISTORY_TEST_DB";

    public ChatHistoryManager()
        : this(
            Environment.GetEnvironmentVariable(TestDbEnvVar) is { Length: > 0 } testDb
                ? testDb
                : $"Data Source={Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ChatHistory.db")}"
        ) { }

    public ChatHistoryManager(string connectionString)
    {
        _connectionString = connectionString;
        using var connection = new SqliteConnection(_connectionString);
        connection.Open();

        var createTitles = connection.CreateCommand();
        createTitles.CommandText =
            @"CREATE TABLE IF NOT EXISTS SessionTitles (
                SessionId TEXT NOT NULL,
                IsAgent INTEGER NOT NULL DEFAULT 0,
                Title TEXT NOT NULL,
                PRIMARY KEY (SessionId, IsAgent)
            )";
        createTitles.ExecuteNonQuery();

        var create = connection.CreateCommand();
        create.CommandText =
            @"CREATE TABLE IF NOT EXISTS ChatHistory (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Role TEXT NOT NULL,
                Message TEXT NOT NULL,
                IsUser INTEGER NOT NULL,
                Timestamp DATETIME NOT NULL,
                SessionId TEXT NOT NULL DEFAULT '',
                IsAgent INTEGER NOT NULL DEFAULT 0
            )";
        create.ExecuteNonQuery();

        // 为旧库补列，列已存在时 ALTER TABLE 会抛异常，直接吞掉
        foreach (
            var ddl in new[]
            {
                "ALTER TABLE ChatHistory ADD COLUMN SessionId TEXT NOT NULL DEFAULT ''",
                "ALTER TABLE ChatHistory ADD COLUMN IsAgent INTEGER NOT NULL DEFAULT 0",
            }
        )
        {
            try
            {
                var alter = connection.CreateCommand();
                alter.CommandText = ddl;
                alter.ExecuteNonQuery();
            }
            catch { }
        }
    }

    public async Task SaveChatMessageAsync(ChatMessage message)
    {
        using var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync();
        var cmd = connection.CreateCommand();
        cmd.CommandText =
            @"INSERT INTO ChatHistory (Role, Message, IsUser, Timestamp, SessionId, IsAgent)
              VALUES (@Role, @Message, @IsUser, @Timestamp, @SessionId, @IsAgent)";
        cmd.Parameters.AddWithValue("@Role", message.Role);
        cmd.Parameters.AddWithValue("@Message", message.Message);
        cmd.Parameters.AddWithValue("@IsUser", message.IsUser ? 1 : 0);
        cmd.Parameters.AddWithValue("@Timestamp", message.Timestamp);
        cmd.Parameters.AddWithValue("@SessionId", message.SessionId ?? "");
        cmd.Parameters.AddWithValue("@IsAgent", message.IsAgent ? 1 : 0);
        await cmd.ExecuteNonQueryAsync();
    }

    /// <summary>加载历史消息。isAgent=true 只返回 Agent 记录，false 只返回 Chat 记录。</summary>
    public List<ChatMessage> LoadChatHistory(
        int limit = 50,
        string sessionId = "",
        bool isAgent = false
    )
    {
        var chatHistory = new List<ChatMessage>();
        using var connection = new SqliteConnection(_connectionString);
        connection.Open();
        var cmd = connection.CreateCommand();

        var agentFilter = isAgent ? "IsAgent = 1" : "IsAgent = 0";
        if (!string.IsNullOrEmpty(sessionId))
        {
            cmd.CommandText =
                $"SELECT Role, Message, IsUser, Timestamp, SessionId, IsAgent FROM ChatHistory WHERE SessionId = @sid AND {agentFilter} ORDER BY Timestamp DESC"
                + (limit > 0 ? $" LIMIT {limit}" : "");
            cmd.Parameters.AddWithValue("@sid", sessionId);
        }
        else
        {
            cmd.CommandText =
                $"SELECT Role, Message, IsUser, Timestamp, SessionId, IsAgent FROM ChatHistory WHERE Timestamp > '0002-01-01' AND {agentFilter} ORDER BY Timestamp DESC"
                + (limit > 0 ? $" LIMIT {limit}" : "");
        }

        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            chatHistory.Add(
                new ChatMessage
                {
                    Role = reader.GetString(0),
                    Message = reader.GetString(1),
                    IsUser = reader.GetInt32(2) == 1,
                    Timestamp = reader.GetDateTime(3),
                    SessionId = reader.IsDBNull(4) ? "" : reader.GetString(4),
                    IsAgent = reader.GetInt32(5) == 1,
                }
            );
        }
        chatHistory.Reverse();
        return chatHistory;
    }

    public int GetHistoryCount(bool isAgent = false)
    {
        using var connection = new SqliteConnection(_connectionString);
        connection.Open();
        var cmd = connection.CreateCommand();
        cmd.CommandText =
            $"SELECT COUNT(*) FROM ChatHistory WHERE Timestamp > '0002-01-01' AND IsAgent = {(isAgent ? 1 : 0)}";
        return Convert.ToInt32(cmd.ExecuteScalar());
    }

    /// <summary>列出所有不重复的 SessionId（按最近消息时间降序）。</summary>
    public List<string> ListSessions()
    {
        var sessions = new List<string>();
        using var connection = new SqliteConnection(_connectionString);
        connection.Open();
        var cmd = connection.CreateCommand();
        cmd.CommandText =
            "SELECT SessionId FROM ChatHistory WHERE SessionId != '' GROUP BY SessionId ORDER BY MAX(Timestamp) DESC";
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
            sessions.Add(reader.GetString(0));
        return sessions;
    }

    public void SaveSessionTitle(string sessionId, string title, bool isAgent = false)
    {
        using var conn = new SqliteConnection(_connectionString);
        conn.Open();
        var cmd = conn.CreateCommand();
        cmd.CommandText =
            @"INSERT INTO SessionTitles (SessionId, IsAgent, Title)
              VALUES (@sid, @isAgent, @title)
              ON CONFLICT(SessionId, IsAgent) DO UPDATE SET Title = excluded.Title";
        cmd.Parameters.AddWithValue("@sid", sessionId);
        cmd.Parameters.AddWithValue("@isAgent", isAgent ? 1 : 0);
        cmd.Parameters.AddWithValue("@title", title);
        cmd.ExecuteNonQuery();
    }

    public void DeleteAllHistory(bool isAgent = false)
    {
        using var conn = new SqliteConnection(_connectionString);
        conn.Open();
        var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM ChatHistory WHERE IsAgent = @isAgent";
        cmd.Parameters.AddWithValue("@isAgent", isAgent ? 1 : 0);
        cmd.ExecuteNonQuery();
    }

    public void DeleteSession(string sessionId)
    {
        using var conn = new SqliteConnection(_connectionString);
        conn.Open();
        var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM ChatHistory WHERE SessionId = @sid";
        cmd.Parameters.AddWithValue("@sid", sessionId);
        cmd.ExecuteNonQuery();
    }

    /// <summary>从另一个 ChatHistory.db 导入会话，每条会话分配新 SessionId 避免冲突。返回导入的会话数。</summary>
    public int ImportSessionsFromDb(string sourcePath)
    {
        var srcConn = $"Data Source={sourcePath}";
        using var src = new SqliteConnection(srcConn);
        src.Open();

        // 确保源库有必要的列（旧版本可能没有 IsAgent）
        var sessionMap = new Dictionary<string, string>();
        var srcCmd = src.CreateCommand();
        srcCmd.CommandText =
            "SELECT Role, Message, IsUser, Timestamp, SessionId, COALESCE(IsAgent,0) FROM ChatHistory WHERE SessionId != ''";
        using var reader = srcCmd.ExecuteReader();

        using var dest = new SqliteConnection(_connectionString);
        dest.Open();

        while (reader.Read())
        {
            var oldSid = reader.GetString(4);
            if (!sessionMap.ContainsKey(oldSid))
                sessionMap[oldSid] = Guid.NewGuid().ToString("N")[..12];

            var ins = dest.CreateCommand();
            ins.CommandText =
                @"INSERT INTO ChatHistory (Role, Message, IsUser, Timestamp, SessionId, IsAgent)
                  VALUES (@Role, @Message, @IsUser, @Timestamp, @SessionId, @IsAgent)";
            ins.Parameters.AddWithValue("@Role", reader.GetString(0));
            ins.Parameters.AddWithValue("@Message", reader.GetString(1));
            ins.Parameters.AddWithValue("@IsUser", reader.GetInt32(2));
            ins.Parameters.AddWithValue("@Timestamp", reader.GetString(3));
            ins.Parameters.AddWithValue("@SessionId", sessionMap[oldSid]);
            ins.Parameters.AddWithValue("@IsAgent", reader.GetInt32(5));
            ins.ExecuteNonQuery();
        }
        return sessionMap.Count;
    }

    public List<ChatSession> ListSessionsWithPreview(bool isAgent = false)
    {
        var result = new List<ChatSession>();
        using var conn = new SqliteConnection(_connectionString);
        conn.Open();
        var cmd = conn.CreateCommand();
        cmd.CommandText =
            @"
        SELECT ch.SessionId,
               MAX(ch.Timestamp) as LastTime,
               (SELECT SUBSTR(sub.Message, 1, 40)
                FROM ChatHistory sub
                WHERE sub.SessionId = ch.SessionId
                  AND sub.IsUser = 1
                  AND sub.IsAgent = @isAgent
                ORDER BY sub.Timestamp
                LIMIT 1) as Preview,
               (SELECT st.Title FROM SessionTitles st
                WHERE st.SessionId = ch.SessionId AND st.IsAgent = @isAgent
                LIMIT 1) as Title
        FROM ChatHistory ch
        WHERE ch.SessionId != '' AND ch.IsAgent = @isAgent
        GROUP BY ch.SessionId
        ORDER BY MAX(ch.Timestamp) DESC
        LIMIT 50";
        cmd.Parameters.AddWithValue("@isAgent", isAgent ? 1 : 0);
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            var sid = reader.GetString(0);
            var lastTime = reader.GetDateTime(1);
            var preview = reader.IsDBNull(2) ? "(空对话)" : reader.GetString(2);
            var title = reader.IsDBNull(3) ? null : reader.GetString(3);
            result.Add(new ChatSession(sid, lastTime, preview, title));
        }
        return result;
    }
}
