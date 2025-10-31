using Microsoft.Data.Sqlite;
using MiniExcelLibs;

namespace NumDesTools.Advance
{
    internal class ExcelDataToDb
    {
        #region Excel数据DB化
        public void ConvertWithSchemaInference(string rootPath, string dbPath)
        {
            var filesCollection = new SelfExcelFileCollector(rootPath);
            var files = filesCollection.GetAllExcelFilesPath();

            if (File.Exists(dbPath))
                File.Delete(dbPath);

            using var connection = new SqliteConnection($"Data Source={dbPath}");
            connection.Open();

            foreach (var file in files)
            {
                if (!File.Exists(file))
                {
                    Debug.WriteLine($"文件不存在，跳过: {file}");
                    continue;
                }

                Debug.WriteLine($"处理文件: {file}");
                ProcessSingleExcelFile(connection, file);
            }
        }

        private static void ProcessSingleExcelFile(SqliteConnection connection, string file)
        {
            try
            {
                var sheets = MiniExcel.GetSheetNames(file);
                var fileName = Path.GetFileNameWithoutExtension(file);

                foreach (var sheetName in sheets)
                {
                    // 为每个文件的工作表创建唯一表名
                    var tableName = $"{fileName}_{sheetName}";

                    var rows = MiniExcel.Query(file, sheetName: sheetName).ToList();

                    if (rows.Count > 0)
                    {
                        var columnTypes = InferColumnTypes(rows);
                        CreateTableWithTypes(connection, tableName, columnTypes);
                        InsertDataWithTypes(connection, tableName, rows, columnTypes);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"处理文件 {file} 失败: {ex.Message}");
            }
        }

        // 补全的 InsertDataWithTypes 方法
        private static void InsertDataWithTypes(
            SqliteConnection connection,
            string tableName,
            List<dynamic> rows,
            Dictionary<string, Type> columnTypes,
            int batchSize = 1000
        )
        {
            if (rows.Count == 0)
                return;

            var columns = columnTypes.Keys.ToList();
            var columnNames = string.Join(", ", columns.Select(c => $"[{c}]"));
            var parameterNames = string.Join(", ", columns.Select(c => $"@{c}"));

            var insertSql = $"INSERT INTO [{tableName}] ({columnNames}) VALUES ({parameterNames})";

            for (int batchStart = 0; batchStart < rows.Count; batchStart += batchSize)
            {
                var batchRows = rows.Skip(batchStart).Take(batchSize).ToList();

                using var transaction = connection.BeginTransaction();
                try
                {
                    using var command = new SqliteCommand(insertSql, connection, transaction);

                    // 预定义参数
                    foreach (var column in columns)
                    {
                        command.Parameters.Add($"@{column}", MapTypeToSqlite(columnTypes[column]));
                    }

                    foreach (var row in batchRows)
                    {
                        var rowDict = (IDictionary<string, object>)row;

                        foreach (var column in columns)
                        {
                            var value = rowDict.ContainsKey(column)
                                ? rowDict[column]
                                : DBNull.Value;
                            command.Parameters[$"@{column}"].Value = ConvertValue(
                                value,
                                columnTypes[column]
                            );
                        }

                        command.ExecuteNonQuery();
                    }

                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    throw new Exception(
                        $"插入批次数据失败 (行 {batchStart}-{batchStart + batchSize}): {ex.Message}",
                        ex
                    );
                }
            }
        }

        // 值类型转换辅助方法
        private static object ConvertValue(object value, Type targetType)
        {
            if (value == null || value is DBNull)
                return DBNull.Value;

            try
            {
                if (targetType == typeof(int))
                    return Convert.ToInt32(value);
                else if (targetType == typeof(double))
                    return Convert.ToDouble(value);
                else if (targetType == typeof(DateTime))
                    return Convert.ToDateTime(value);
                else
                    return value.ToString();
            }
            catch
            {
                // 转换失败时返回原始字符串
                return value.ToString();
            }
        }

        private static Dictionary<string, Type> InferColumnTypes(List<dynamic> rows)
        {
            var types = new Dictionary<string, Type>();

            if (rows.Count == 0)
                return types;

            var firstRow = (IDictionary<string, object>)rows[0];

            foreach (var column in firstRow.Keys)
            {
                Type inferredType = typeof(string); // 默认字符串类型

                // 采样前10行推断类型
                var sampleValues = rows.Take(10)
                    .Select(r => ((IDictionary<string, object>)r)[column])
                    .Where(v => v != null)
                    .ToList();

                if (sampleValues.Count > 0)
                {
                    if (sampleValues.All(v => int.TryParse(v.ToString(), out _)))
                        inferredType = typeof(int);
                    else if (sampleValues.All(v => double.TryParse(v.ToString(), out _)))
                        inferredType = typeof(double);
                    else if (sampleValues.All(v => DateTime.TryParse(v.ToString(), out _)))
                        inferredType = typeof(DateTime);
                }

                types[column] = inferredType;
            }

            return types;
        }

        private static void CreateTableWithTypes(
            SqliteConnection connection,
            string tableName,
            Dictionary<string, Type> columnTypes
        )
        {
            var columns = columnTypes.Select(kvp =>
                $"[{kvp.Key}] {MapTypeToSqliteString(kvp.Value)}"
            ); // 使用字符串版本

            var createTableSql = $"CREATE TABLE [{tableName}] ({string.Join(", ", columns)})";

            using var command = new SqliteCommand(createTableSql, connection);
            command.ExecuteNonQuery();
        }

        private static SqliteType MapTypeToSqlite(Type type)
        {
            return type.Name switch
            {
                "Int32" => SqliteType.Integer,
                "Double" => SqliteType.Real,
                "DateTime" => SqliteType.Text, // SQLite 没有专门的 DateTime 类型
                _ => SqliteType.Text
            };
        }

        // 重载方法用于SQL语句中的类型映射
        private static string MapTypeToSqliteString(Type type)
        {
            return type.Name switch
            {
                "Int32" => "INTEGER",
                "Double" => "REAL",
                "DateTime" => "TEXT",
                _ => "TEXT"
            };
        }

        #endregion

        #region 查询DB返回Excel路径索引

        // 1. 搜索所有表格中的特定值
        public List<SearchResult> SearchAllTables(
            string searchValue,
            string _dbPath,
            bool exactMatch = false
        )
        {
            var results = new List<SearchResult>();

            using var connection = new SqliteConnection($"Data Source={_dbPath}");
            connection.Open();

            // 获取所有表名
            var tableNames = GetTableNames(connection);

            foreach (var tableName in tableNames)
            {
                var tableResults = SearchInTable(connection, tableName, searchValue, exactMatch);
                results.AddRange(tableResults);
            }

            return results;
        }

        // 2. 获取数据库中的所有表名
        private List<string> GetTableNames(SqliteConnection connection)
        {
            var tableNames = new List<string>();

            var command = new SqliteCommand(
                "SELECT name FROM sqlite_master WHERE type='table'",
                connection
            );

            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                tableNames.Add(reader.GetString(0));
            }

            return tableNames;
        }

        // 3. 在单个表中搜索
        private List<SearchResult> SearchInTable(
            SqliteConnection connection,
            string tableName,
            string searchValue,
            bool exactMatch
        )
        {
            var results = new List<SearchResult>();

            try
            {
                // 获取表结构
                var columns = GetTableColumns(connection, tableName);
                if (columns.Count == 0)
                    return results;

                // 构建查询条件
                var conditions = new List<string>();
                foreach (var column in columns)
                {
                    if (exactMatch)
                        conditions.Add($"[{column}] = @SearchValue");
                    else
                        conditions.Add($"[{column}] LIKE '%' || @SearchValue || '%'");
                }

                string whereClause = string.Join(" OR ", conditions);
                string query = $"SELECT * FROM [{tableName}] WHERE {whereClause}";

                var command = new SqliteCommand(query, connection);
                command.Parameters.AddWithValue("@SearchValue", searchValue);

                using var reader = command.ExecuteReader();
                int rowNumber = 1;

                while (reader.Read())
                {
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        if (!reader.IsDBNull(i))
                        {
                            string value = reader.GetString(i);
                            bool isMatch = exactMatch
                                ? value.Equals(searchValue, StringComparison.OrdinalIgnoreCase)
                                : value.IndexOf(searchValue, StringComparison.OrdinalIgnoreCase)
                                    >= 0;

                            if (isMatch)
                            {
                                results.Add(
                                    new SearchResult
                                    {
                                        TableName = tableName,
                                        ColumnName = reader.GetName(i),
                                        RowNumber = rowNumber,
                                        Value = value,
                                        FileName = ExtractFileNameFromTable(tableName)
                                    }
                                );
                            }
                        }
                    }
                    rowNumber++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"搜索表 {tableName} 时出错: {ex.Message}");
            }

            return results;
        }

        // 4. 获取表的列信息
        private List<string> GetTableColumns(SqliteConnection connection, string tableName)
        {
            var columns = new List<string>();

            var command = new SqliteCommand($"PRAGMA table_info([{tableName}])", connection);

            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                columns.Add(reader.GetString(1)); // name列
            }

            return columns;
        }

        // 5. 从表名中提取原始文件名
        private string ExtractFileNameFromTable(string tableName)
        {
            // 根据你的表名格式 "文件名_工作表名" 来提取
            var parts = tableName.Split('_');
            return parts.Length > 0 ? parts[0] : tableName;
        }

        // 6. 高级搜索：支持多条件
        public List<SearchResult> AdvancedSearch(
            Dictionary<string, string> conditions,
            string _dbPath
        )
        {
            var results = new List<SearchResult>();

            using var connection = new SqliteConnection($"Data Source={_dbPath}");
            connection.Open();

            var tableNames = GetTableNames(connection);

            foreach (var tableName in tableNames)
            {
                var tableResults = SearchInTableAdvanced(connection, tableName, conditions);
                results.AddRange(tableResults);
            }

            return results;
        }

        private List<SearchResult> SearchInTableAdvanced(
            SqliteConnection connection,
            string tableName,
            Dictionary<string, string> conditions
        )
        {
            var results = new List<SearchResult>();

            try
            {
                var columns = GetTableColumns(connection, tableName);
                var whereConditions = new List<string>();
                var parameters = new List<SqliteParameter>();

                int paramIndex = 0;
                foreach (var condition in conditions)
                {
                    if (columns.Contains(condition.Key))
                    {
                        whereConditions.Add($"[{condition.Key}] LIKE '%' || @p{paramIndex} || '%'");
                        parameters.Add(new SqliteParameter($"@p{paramIndex}", condition.Value));
                        paramIndex++;
                    }
                }

                if (whereConditions.Count == 0)
                    return results;

                string whereClause = string.Join(" AND ", whereConditions);
                string query = $"SELECT * FROM [{tableName}] WHERE {whereClause}";

                var command = new SqliteCommand(query, connection);
                command.Parameters.AddRange(parameters.ToArray());

                using var reader = command.ExecuteReader();
                int rowNumber = 1;

                while (reader.Read())
                {
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        if (!reader.IsDBNull(i))
                        {
                            results.Add(
                                new SearchResult
                                {
                                    TableName = tableName,
                                    ColumnName = reader.GetName(i),
                                    RowNumber = rowNumber,
                                    Value = reader.GetString(i),
                                    FileName = ExtractFileNameFromTable(tableName)
                                }
                            );
                        }
                    }
                    rowNumber++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"高级搜索表 {tableName} 时出错: {ex.Message}");
            }

            return results;
        }

        // 搜索结果类
        public class SearchResult
        {
            public string TableName { get; set; } // 表名
            public string ColumnName { get; set; } // 列名
            public int RowNumber { get; set; } // 行号
            public string Value { get; set; } // 匹配的值
            public string FileName { get; set; } // 原始文件名

            public override string ToString()
            {
                return $"文件: {FileName}, 表: {TableName}, 列: {ColumnName}, 行: {RowNumber}, 值: {Value}";
            }
        }

        // 统计信息类
        public class SearchStatistics
        {
            public int TotalMatches { get; set; }
            public int TablesSearched { get; set; }
            public Dictionary<string, int> MatchesPerTable { get; set; } = new();
            public TimeSpan SearchTime { get; set; }
        }

        #endregion
    }
}
