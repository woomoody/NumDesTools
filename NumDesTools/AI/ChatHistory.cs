using System.Threading.Tasks;
using Microsoft.Data.Sqlite;

namespace NumDesTools.AI;

public class ChatMessage
{
    public string Role { get; set; }
    public string Message { get; set; }
    public bool IsUser { get; set; }
    public DateTime Timestamp { get; set; }
    public bool IsStreaming { get; set; }
    public string SessionId { get; set; }
}

public class ChatHistoryManager
{
    private readonly string _connectionString =
        $"Data Source={Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ChatHistory.db")}";

    public ChatHistoryManager()
    {
        using var connection = new SqliteConnection(_connectionString);
        connection.Open();
        var cmd = connection.CreateCommand();
        cmd.CommandText =
            @"
                CREATE TABLE IF NOT EXISTS ChatHistory (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Role TEXT NOT NULL,
                    Message TEXT NOT NULL,
                    IsUser INTEGER NOT NULL,
                    Timestamp DATETIME NOT NULL,
                    SessionId TEXT NOT NULL DEFAULT ''
                );
                -- 历史数据无 SessionId 列时自动补列（SQLite 支持 ADD COLUMN）
                ALTER TABLE ChatHistory ADD COLUMN SessionId TEXT NOT NULL DEFAULT '' ON CONFLICT IGNORE;";
        try
        {
            cmd.ExecuteNonQuery();
        }
        catch
        {
            // ADD COLUMN 在列已存在时会抛，吞掉即可
        }
    }

    public async Task SaveChatMessageAsync(ChatMessage message)
    {
        using var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync();
        var cmd = connection.CreateCommand();
        cmd.CommandText =
            @"
                INSERT INTO ChatHistory (Role, Message, IsUser, Timestamp, SessionId)
                VALUES (@Role, @Message, @IsUser, @Timestamp, @SessionId)";
        cmd.Parameters.AddWithValue("@Role", message.Role);
        cmd.Parameters.AddWithValue("@Message", message.Message);
        cmd.Parameters.AddWithValue("@IsUser", message.IsUser ? 1 : 0);
        cmd.Parameters.AddWithValue("@Timestamp", message.Timestamp);
        cmd.Parameters.AddWithValue("@SessionId", message.SessionId ?? "");
        await cmd.ExecuteNonQueryAsync();
    }

    /// <summary>加载历史消息，sessionId 为空时跨会话加载最近 N 条。</summary>
    public List<ChatMessage> LoadChatHistory(int limit = 50, string sessionId = "")
    {
        var chatHistory = new List<ChatMessage>();
        using var connection = new SqliteConnection(_connectionString);
        connection.Open();
        var cmd = connection.CreateCommand();

        if (!string.IsNullOrEmpty(sessionId))
        {
            cmd.CommandText =
                limit > 0
                    ? $"SELECT Role, Message, IsUser, Timestamp, SessionId FROM ChatHistory WHERE SessionId = @sid ORDER BY Timestamp DESC LIMIT {limit}"
                    : "SELECT Role, Message, IsUser, Timestamp, SessionId FROM ChatHistory WHERE SessionId = @sid ORDER BY Timestamp DESC";
            cmd.Parameters.AddWithValue("@sid", sessionId);
        }
        else
        {
            cmd.CommandText =
                limit > 0
                    ? $"SELECT Role, Message, IsUser, Timestamp, SessionId FROM ChatHistory WHERE Timestamp > '0002-01-01' ORDER BY Timestamp DESC LIMIT {limit}"
                    : "SELECT Role, Message, IsUser, Timestamp, SessionId FROM ChatHistory WHERE Timestamp > '0002-01-01' ORDER BY Timestamp DESC";
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
                }
            );
        }
        chatHistory.Reverse();
        return chatHistory;
    }

    public int GetHistoryCount()
    {
        using var connection = new SqliteConnection(_connectionString);
        connection.Open();
        var cmd = connection.CreateCommand();
        cmd.CommandText = "SELECT COUNT(*) FROM ChatHistory WHERE Timestamp > '0002-01-01'";
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
}
