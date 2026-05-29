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
}

public class ChatHistoryManager
{
    private readonly string _connectionString =
        $"Data Source={Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ChatHistory.db")}";

    public ChatHistoryManager()
    {
        using var connection = new SqliteConnection(_connectionString);
        connection.Open();
        var command = connection.CreateCommand();
        command.CommandText =
            @"
                CREATE TABLE IF NOT EXISTS ChatHistory (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Role TEXT NOT NULL,
                    Message TEXT NOT NULL,
                    IsUser INTEGER NOT NULL,
                    Timestamp DATETIME NOT NULL
                )";
        command.ExecuteNonQuery();
    }

    public async Task SaveChatMessageAsync(ChatMessage message)
    {
        using var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync();
        var command = connection.CreateCommand();
        command.CommandText =
            @"
                INSERT INTO ChatHistory (Role, Message, IsUser, Timestamp)
                VALUES (@Role, @Message, @IsUser, @Timestamp)";
        command.Parameters.AddWithValue("@Role", message.Role);
        command.Parameters.AddWithValue("@Message", message.Message);
        command.Parameters.AddWithValue("@IsUser", message.IsUser ? 1 : 0);
        command.Parameters.AddWithValue("@Timestamp", message.Timestamp);
        await command.ExecuteNonQueryAsync();
    }

    public List<ChatMessage> LoadChatHistory(int limit = 50)
    {
        var chatHistory = new List<ChatMessage>();
        using var connection = new SqliteConnection(_connectionString);
        connection.Open();
        var command = connection.CreateCommand();
        command.CommandText =
            limit > 0
                ? $"SELECT Role, Message, IsUser, Timestamp FROM ChatHistory WHERE Timestamp > '0002-01-01' ORDER BY Timestamp DESC LIMIT {limit}"
                : "SELECT Role, Message, IsUser, Timestamp FROM ChatHistory WHERE Timestamp > '0002-01-01' ORDER BY Timestamp DESC";

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            chatHistory.Add(
                new ChatMessage
                {
                    Role = reader.GetString(0),
                    Message = reader.GetString(1),
                    IsUser = reader.GetInt32(2) == 1,
                    Timestamp = reader.GetDateTime(3),
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
        var command = connection.CreateCommand();
        command.CommandText = "SELECT COUNT(*) FROM ChatHistory WHERE Timestamp > '0002-01-01'";
        return Convert.ToInt32(command.ExecuteScalar());
    }
}
