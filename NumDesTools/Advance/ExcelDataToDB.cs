using Microsoft.Data.Sqlite;
using MiniExcelLibs;

namespace NumDesTools.Advance
{
    internal class ExcelDataToDb
    {
        #region 单个文件更新方法

        /// <summary>
        /// 更新单个Excel文件到数据库（支持增量更新）
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <param name="dbPath">数据库路径</param>
        /// <param name="updateMode">更新模式：覆盖或追加</param>
        public void UpdateSingleFile(string filePath, string dbPath, UpdateMode updateMode = UpdateMode.Overwrite)
        {
            if (!File.Exists(filePath))
            {
                Debug.Print($"文件不存在: {filePath}");
                return;
            }

            if (filePath.Contains("#"))
            {
                Debug.Print($"[#]文件，跳过: {filePath}");
                return;
            }

            // 确保数据库存在
            EnsureDatabaseExists(dbPath);

            using var connection = new SqliteConnection($"Data Source={dbPath}");
            connection.Open();

            // 确保元数据表存在
            CreateMetadataTable(connection);

            Debug.Print($"更新文件: {filePath} (模式: {updateMode})");
            UpdateSingleExcelFile(connection, filePath, updateMode);

            // 验证行号映射
            VerifyRowMappings(dbPath);
        }

        /// <summary>
        /// 更新模式枚举
        /// </summary>
        public enum UpdateMode
        {
            /// <summary>
            /// 覆盖模式：删除原有数据后插入新数据
            /// </summary>
            Overwrite,

            /// <summary>
            /// 追加模式：保留原有数据，只添加新数据
            /// </summary>
            Append
        }

        /// <summary>
        /// 确保数据库文件存在
        /// </summary>
        private void EnsureDatabaseExists(string dbPath)
        {
            if (!File.Exists(dbPath))
            {
                // 创建空数据库文件
                using var connection = new SqliteConnection($"Data Source={dbPath}");
                connection.Open();
                CreateMetadataTable(connection);
                Debug.Print($"创建新数据库: {dbPath}");
            }
        }

        /// <summary>
        /// 处理单个Excel文件的更新
        /// </summary>
        private void UpdateSingleExcelFile(SqliteConnection connection, string filePath, UpdateMode updateMode)
        {
            try
            {
                var sheets = MiniExcel.GetSheetNames(filePath);
                var fileNameOnly = Path.GetFileNameWithoutExtension(filePath);
                var fileFullPath = Path.GetFullPath(filePath);

                foreach (var sheetName in sheets)
                {
                    if (sheetName.Contains("#"))
                    {
                        Debug.Print($"[#]工作表，跳过: {sheetName} in file {filePath}");
                        continue;
                    }

                    var tableName = $"{SanitizeTableName(fileNameOnly)}_{SanitizeTableName(sheetName)}";

                    // 检查表是否已存在
                    bool tableExists = CheckTableExists(connection, tableName);

                    if (tableExists)
                    {
                        Debug.Print($"表已存在，执行更新: {tableName}");
                        UpdateExistingTable(connection, tableName, filePath, sheetName, updateMode);
                    }
                    else
                    {
                        Debug.Print($"表不存在，创建新表: {tableName}");
                        ProcessSingleExcelFile(connection, filePath); // 使用原有的创建逻辑
                    }

                    // 更新文件路径映射
                    StoreFilePathMapping(connection, tableName, fileFullPath, sheetName, fileNameOnly);
                }
            }
            catch (Exception ex)
            {
                Debug.Print($"更新文件 {filePath} 失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 检查表是否存在
        /// </summary>
        private bool CheckTableExists(SqliteConnection connection, string tableName)
        {
            try
            {
                var sql = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@tableName";
                using var command = new SqliteCommand(sql, connection);
                command.Parameters.AddWithValue("@tableName", tableName);
                var result = Convert.ToInt32(command.ExecuteScalar());
                return result > 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 更新已存在的表
        /// </summary>
        private void UpdateExistingTable(SqliteConnection connection, string tableName,
            string filePath, string sheetName, UpdateMode updateMode)
        {
            try
            {
                var rows = MiniExcel.Query(filePath, sheetName: sheetName).ToList();

                if (rows.Count == 0)
                {
                    Debug.Print($"Excel文件为空，跳过更新: {filePath}");
                    return;
                }

                // 推断列类型（考虑表结构可能已变化）
                var columnTypes = InferColumnTypes(rows);

                // 检查并更新表结构
                UpdateTableStructure(connection, tableName, columnTypes);

                if (updateMode == UpdateMode.Overwrite)
                {
                    // 覆盖模式：清空表后插入
                    ClearTableData(connection, tableName);
                    InsertDataWithTypes(connection, tableName, rows, columnTypes);
                }
                else
                {
                    // 追加模式：直接插入新数据
                    InsertDataWithTypes(connection, tableName, rows, columnTypes);
                }

                Debug.Print($"表 {tableName} 更新完成，行数: {rows.Count}");
            }
            catch (Exception ex)
            {
                Debug.Print($"更新表 {tableName} 失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 更新表结构（添加新列）
        /// </summary>
        private void UpdateTableStructure(SqliteConnection connection, string tableName,
            Dictionary<string, Type> newColumnTypes)
        {
            try
            {
                var existingColumns = GetTableColumns(connection, tableName);
                var newColumns = newColumnTypes.Keys.Except(existingColumns, StringComparer.OrdinalIgnoreCase);

                foreach (var column in newColumns)
                {
                    var alterSql = $"ALTER TABLE [{tableName}] ADD COLUMN [{column}] {MapTypeToSqliteString(newColumnTypes[column])}";
                    using var command = new SqliteCommand(alterSql, connection);
                    command.ExecuteNonQuery();
                    Debug.Print($"表 {tableName} 添加新列: {column}");
                }
            }
            catch (Exception ex)
            {
                Debug.Print($"更新表结构失败 {tableName}: {ex.Message}");
            }
        }

        /// <summary>
        /// 清空表数据（覆盖模式使用）
        /// </summary>
        private void ClearTableData(SqliteConnection connection, string tableName)
        {
            try
            {
                // 清空数据表
                var deleteDataSql = $"DELETE FROM [{tableName}]";
                using var command1 = new SqliteCommand(deleteDataSql, connection);
                command1.ExecuteNonQuery();

                // 清空行号映射
                var deleteMappingSql = "DELETE FROM _row_mapping WHERE table_name = @tableName";
                using var command2 = new SqliteCommand(deleteMappingSql, connection);
                command2.Parameters.AddWithValue("@tableName", tableName);
                command2.ExecuteNonQuery();

                Debug.Print($"清空表数据: {tableName}");
            }
            catch (Exception ex)
            {
                Debug.Print($"清空表数据失败 {tableName}: {ex.Message}");
            }
        }

        #endregion

        #region 批量更新方法

        /// <summary>
        /// 批量更新多个Excel文件到数据库
        /// </summary>
        /// <param name="filePaths">Excel文件路径集合</param>
        /// <param name="dbPath">数据库路径</param>
        /// <param name="updateMode">更新模式</param>
        public void UpdateMultipleFiles(IEnumerable<string> filePaths, string dbPath, UpdateMode updateMode = UpdateMode.Overwrite)
        {
            EnsureDatabaseExists(dbPath);

            using var connection = new SqliteConnection($"Data Source={dbPath}");
            connection.Open();
            CreateMetadataTable(connection);

            int successCount = 0;
            int totalCount = 0;

            foreach (var filePath in filePaths)
            {
                totalCount++;
                try
                {
                    UpdateSingleExcelFile(connection, filePath, updateMode);
                    successCount++;
                    Debug.Print($"成功更新文件: {filePath}");
                }
                catch (Exception ex)
                {
                    Debug.Print($"更新文件失败 {filePath}: {ex.Message}");
                }
            }

            VerifyRowMappings(dbPath);
            Debug.Print($"批量更新完成: {successCount}/{totalCount} 个文件成功");
        }

        #endregion
        #region Excel数据DB化
        public void ConvertWithSchemaInference(string rootPath, string dbPath)
        {
            var filesCollection = new SelfExcelFileCollector(rootPath);
            var files = filesCollection.GetAllExcelFilesPath();

            if (File.Exists(dbPath))
                File.Delete(dbPath);

            using var connection = new SqliteConnection($"Data Source={dbPath}");
            connection.Open();

            // 第一步：创建元数据表和行号映射表
            CreateMetadataTable(connection);

            foreach (var file in files)
            {
                if (!File.Exists(file))
                {
                    Debug.Print($"文件不存在，跳过: {file}");
                    continue;
                }

                if (file.Contains("#"))
                {
                    Debug.Print($"[#]文件，跳过: {file}");
                    continue;
                }

                Debug.Print($"处理文件: {file}");
                NumDesAddIn.App.StatusBar = $"处理文件: {file}";

                ProcessSingleExcelFile(connection, file);
            }

            // 验证行号映射
            VerifyRowMappings(dbPath);
        }

        private void CreateMetadataTable(SqliteConnection connection)
        {
            var createMetaSql = @"
        CREATE TABLE IF NOT EXISTS _file_metadata (
            table_name TEXT PRIMARY KEY,
            file_full_path TEXT NOT NULL,
            sheet_name TEXT NOT NULL,
            file_name_only TEXT NOT NULL
        );
        
        CREATE TABLE IF NOT EXISTS _row_mapping (
            mapping_id INTEGER PRIMARY KEY AUTOINCREMENT,
            table_name TEXT NOT NULL,
            db_rowid INTEGER NOT NULL,
            excel_row INTEGER NOT NULL,
            UNIQUE(table_name, db_rowid)
        )";

            using var command = new SqliteCommand(createMetaSql, connection);
            command.ExecuteNonQuery();
        }

        private void ProcessSingleExcelFile(SqliteConnection connection, string file)
        {
            try
            {
                var sheets = MiniExcel.GetSheetNames(file);
                var fileNameOnly = Path.GetFileNameWithoutExtension(file);
                var fileFullPath = Path.GetFullPath(file);

                foreach (var sheetName in sheets)
                {
                    if (sheetName.Contains("#"))
                    {
                        Debug.Print($"[#]工作表，跳过: {sheetName} in file {file}");
                        continue;
                    }
                    // 表名格式：文件名_工作表名（不含特殊字符）
                    var tableName = $"{SanitizeTableName(fileNameOnly)}_{SanitizeTableName(sheetName)}";

                    // 存储文件路径映射
                    StoreFilePathMapping(connection, tableName, fileFullPath, sheetName, fileNameOnly);

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
                Debug.Print($"处理文件 {file} 失败: {ex.Message}");
            }
        }
        private string SanitizeTableName(string name)
        {
            // 移除或替换SQLite表名中的非法字符
            return name.Replace(" ", "_")
                      .Replace("-", "_")
                      .Replace(".", "_")
                      .Replace("(", "")
                      .Replace(")", "");
        }
        private void StoreFilePathMapping(SqliteConnection connection, string tableName,
       string fileFullPath, string sheetName, string fileNameOnly)
        {
            var insertMetaSql = @"
                INSERT OR REPLACE INTO _file_metadata 
                (table_name, file_full_path, sheet_name, file_name_only) 
                VALUES (@tableName, @filePath, @sheetName, @fileNameOnly)";

            using var command = new SqliteCommand(insertMetaSql, connection);
            command.Parameters.AddWithValue("@tableName", tableName);
            command.Parameters.AddWithValue("@filePath", fileFullPath);
            command.Parameters.AddWithValue("@sheetName", sheetName);
            command.Parameters.AddWithValue("@fileNameOnly", fileNameOnly);

            command.ExecuteNonQuery();
        }

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

                    // 获取当前最大行ID
                    long currentMaxId = GetMaxRowId(connection, tableName, transaction);
                    long nextRowId = currentMaxId + 1;

                    foreach (var (row, indexInBatch) in batchRows.Select((r, i) => (r, i)))
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

                        // 计算Excel行号
                        int excelRowNumber = batchStart + indexInBatch + 1;

                        // 记录行号映射
                        StoreRowMapping(connection, tableName, nextRowId, excelRowNumber, transaction);

                        nextRowId++;
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

        // 获取表的最大行ID
        private static long GetMaxRowId(SqliteConnection connection, string tableName, SqliteTransaction transaction = null)
        {
            try
            {
                var sql = $"SELECT COALESCE(MAX(rowid), 0) FROM [{tableName}]";
                using var command = transaction != null
                    ? new SqliteCommand(sql, connection, transaction)
                    : new SqliteCommand(sql, connection);

                return Convert.ToInt64(command.ExecuteScalar());
            }
            catch
            {
                return 0;
            }
        }

        // 存储行号映射
        private static void StoreRowMapping(SqliteConnection connection, string tableName,
            long dbRowId, int excelRow, SqliteTransaction transaction = null)
        {
            try
            {
                var sql = @"
            INSERT INTO _row_mapping (table_name, db_rowid, excel_row)
            VALUES (@tableName, @dbRowId, @excelRow)";

                using var command = transaction != null
                    ? new SqliteCommand(sql, connection, transaction)
                    : new SqliteCommand(sql, connection);

                command.Parameters.AddWithValue("@tableName", tableName);
                command.Parameters.AddWithValue("@dbRowId", dbRowId);
                command.Parameters.AddWithValue("@excelRow", excelRow);

                command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Debug.Print($"存储行号映射失败: {ex.Message}");
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
                _ => "TEXT"
            };
        }

        // 获取Excel行号（备用方法）

        // 验证行号映射
        public void VerifyRowMappings(string dbPath)
        {
            using var connection = new SqliteConnection($"Data Source={dbPath}");
            connection.Open();

            var tableNames = GetTableNames(connection);

            foreach (var tableName in tableNames)
            {
                if (tableName == "_file_metadata" || tableName == "_row_mapping") continue;

                // 检查数据行数
                var dataCount = GetRowCount(connection, tableName);

                // 检查映射行数
                var mappingCount = GetMappingCount(connection, tableName);

                Debug.Print($"表 {tableName}: 数据行={dataCount}, 映射行={mappingCount}, 状态={(dataCount == mappingCount ? "正常" : "异常")}");
            }
        }

        private int GetRowCount(SqliteConnection connection, string tableName)
        {
            var sql = $"SELECT COUNT(*) FROM [{tableName}]";
            using var command = new SqliteCommand(sql, connection);
            return Convert.ToInt32(command.ExecuteScalar());
        }

        private int GetMappingCount(SqliteConnection connection, string tableName)
        {
            var sql = "SELECT COUNT(*) FROM _row_mapping WHERE table_name = @tableName";
            using var command = new SqliteCommand(sql, connection);
            command.Parameters.AddWithValue("@tableName", tableName);
            return Convert.ToInt32(command.ExecuteScalar());
        }

        #endregion

        #region 查询DB返回Excel路径索引

        public class SearchOptions
        {
            public string SearchValue { get; set; }
            public string[] TargetColumns { get; set; } // 指定要搜索的列
            public bool ExactMatch { get; set; }
            public bool CaseSensitive { get; set; }

            public SearchOptions(string searchValue)
            {
                SearchValue = searchValue;
                ExactMatch = !searchValue.Contains("*");
                SearchValue = searchValue.Replace("*", "");
                TargetColumns = Array.Empty<string>(); // 默认空数组表示搜索所有列
                CaseSensitive = false;
            }
        }

        public List<SearchResult> SearchAllTables(string searchValue, string dbPath)
        {
            var options = new SearchOptions(searchValue);
            return SearchAllTables(options, dbPath);
        }

        public List<SearchResult> SearchAllTables(SearchOptions options, string dbPath)
        {
            var results = new List<SearchResult>();

            using var connection = new SqliteConnection($"Data Source={dbPath}");
            connection.Open();

            var tableNames = GetTableNames(connection);

            foreach (var tableName in tableNames)
            {
                if (tableName == "_file_metadata") continue;

                var tableResults = SearchInTable(connection, tableName, options);
                results.AddRange(tableResults);
            }

            return results;
        }

        public List<SearchResult> SearchInColumns(string dbPath, string searchValue, params string[] columnNames)
        {
            var options = new SearchOptions(searchValue)
            {
                TargetColumns = columnNames
            };

            return SearchAllTables(options, dbPath);
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

        private List<SearchResult> SearchInTable(SqliteConnection connection, string tableName, SearchOptions options)
        {
            var results = new List<SearchResult>();

            try
            {
                // 从元数据表获取文件路径和工作表名
                var fileInfo = GetFileInfoFromMetadata(connection, tableName);
                if (fileInfo == null) return results;

                // 获取表的所有列
                var allColumns = GetTableColumns(connection, tableName);
                if (allColumns.Count == 0) return results;

                // 确定要搜索的列
                var columnsToSearch = options.TargetColumns.Length > 0
                    ? allColumns.Intersect(options.TargetColumns, StringComparer.OrdinalIgnoreCase).ToList()
                    : allColumns;

                if (columnsToSearch.Count == 0)
                {
                    Debug.Print($"表 {tableName} 中没有找到指定的列: {string.Join(", ", options.TargetColumns)}");
                    return results;
                }

                // 构建查询条件
                var conditions = BuildSearchConditions(columnsToSearch, options);
                if (conditions.Count == 0) return results;

                string whereClause = string.Join(" OR ", conditions);

                // 修改查询，获取行号映射
                string query = $@"
            SELECT t.*, rm.excel_row 
            FROM [{tableName}] t
            LEFT JOIN _row_mapping rm ON t.rowid = rm.db_rowid AND rm.table_name = @tableName
            WHERE {whereClause}";

                using var command = new SqliteCommand(query, connection);
                command.Parameters.AddWithValue("@tableName", tableName);
                command.Parameters.AddWithValue("@SearchValue", options.SearchValue);

                using var reader = command.ExecuteReader();

                while (reader.Read())
                {
                    // 获取Excel行号（从映射表）
                    int excelRowNumber = 1;
                    if (!reader.IsDBNull(reader.FieldCount - 1))
                    {
                        excelRowNumber = reader.GetInt32(reader.FieldCount - 1);
                    }

                    // 检查数据列（排除最后一个excel_row字段）
                    for (int i = 0; i < reader.FieldCount - 1; i++)
                    {
                        string columnName = reader.GetName(i);

                        // 如果指定了列，只检查指定的列
                        if (options.TargetColumns.Length > 0 &&
                            !options.TargetColumns.Contains(columnName, StringComparer.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        if (!reader.IsDBNull(i))
                        {
                            string value = reader.GetString(i);
                            bool isMatch = IsValueMatch(value, options.SearchValue, options);

                            if (isMatch)
                            {
                                results.Add(new SearchResult
                                {
                                    TableName = fileInfo.SheetName,
                                    ColumnName = columnName,
                                    RowNumber = excelRowNumber, // 使用正确的Excel行号
                                    Value = value,
                                    FileName = fileInfo.FileFullPath,
                                    FileNameOnly = fileInfo.FileNameOnly
                                });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.Print($"搜索表 {tableName} 时出错: {ex.Message}");
            }

            return results;
        }

        private List<string> BuildSearchConditions(List<string> columns, SearchOptions options)
        {
            var conditions = new List<string>();

            foreach (var column in columns)
            {
                if (options.ExactMatch)
                {
                    conditions.Add(options.CaseSensitive
                        ? $"[{column}] = @SearchValue"
                        : $"LOWER([{column}]) = LOWER(@SearchValue)");
                }
                else
                {
                    conditions.Add(options.CaseSensitive
                        ? $"[{column}] LIKE '%' || @SearchValue || '%'"
                        : $"LOWER([{column}]) LIKE '%' || LOWER(@SearchValue) || '%'");
                }
            }

            return conditions;
        }

        private bool IsValueMatch(string cellValue, string searchValue, SearchOptions options)
        {
            if (string.IsNullOrEmpty(cellValue)) return false;

            if (options.ExactMatch)
            {
                return options.CaseSensitive
                    ? cellValue.Equals(searchValue)
                    : cellValue.Equals(searchValue, StringComparison.OrdinalIgnoreCase);
            }
            else
            {
                return options.CaseSensitive
                    ? cellValue.Contains(searchValue)
                    : cellValue.Contains(searchValue, StringComparison.OrdinalIgnoreCase);
            }
        }


        private FileInfoResult GetFileInfoFromMetadata(SqliteConnection connection, string tableName)
        {
            try
            {
                var query = "SELECT file_full_path, sheet_name, file_name_only FROM _file_metadata WHERE table_name = @tableName";

                using var command = new SqliteCommand(query, connection);
                command.Parameters.AddWithValue("@tableName", tableName);

                using var reader = command.ExecuteReader();
                if (reader.Read())
                {
                    return new FileInfoResult
                    {
                        FileFullPath = reader.GetString(0),
                        SheetName = reader.GetString(1),
                        FileNameOnly = reader.GetString(2)
                    };
                }
            }
            catch (Exception ex)
            {
                Debug.Print($"获取表 {tableName} 的元数据失败: {ex.Message}");
            }

            return null;
        }

        private List<string> GetTableColumns(SqliteConnection connection, string tableName)
        {
            var columns = new List<string>();

            try
            {
                var command = new SqliteCommand($"PRAGMA table_info([{tableName}])", connection);

                using var reader = command.ExecuteReader();
                while (reader.Read())
                {
                    columns.Add(reader.GetString(1)); // name列
                }
            }
            catch (Exception ex)
            {
                Debug.Print($"获取表 {tableName} 的列信息失败: {ex.Message}");
            }

            return columns;
        }

        // 多列精确搜索
        public List<SearchResult> SearchExactInColumns(string dbPath, string searchValue, params string[] columnNames)
        {
            var options = new SearchOptions(searchValue)
            {
                TargetColumns = columnNames,
                ExactMatch = true
            };

            return SearchAllTables(options, dbPath);
        }

        // 大小写敏感搜索
        public List<SearchResult> SearchCaseSensitive(string dbPath, string searchValue, params string[] columnNames)
        {
            var options = new SearchOptions(searchValue)
            {
                TargetColumns = columnNames,
                CaseSensitive = true
            };

            return SearchAllTables(options, dbPath);
        }

        // 组合搜索：多列+精确+大小写敏感
        public List<SearchResult> SearchAdvanced(string dbPath, string searchValue,
            bool exactMatch, bool caseSensitive, params string[] columnNames)
        {
            var options = new SearchOptions(searchValue)
            {
                TargetColumns = columnNames,
                ExactMatch = exactMatch,
                CaseSensitive = caseSensitive
            };

            return SearchAllTables(options, dbPath);
        }

        private string ExtractFileNameFromTable(string tableName)
        {
            // 根据你的表名格式 "文件名_工作表名" 来提取
            var parts = tableName.Split('_');
            return parts.Length > 0 ? parts[0] : tableName;
        }


        public List<SearchResult> AdvancedSearch(
            Dictionary<string, string> conditions,
            string dbPath
        )
        {
            var results = new List<SearchResult>();

            using var connection = new SqliteConnection($"Data Source={dbPath}");
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

        // 文件信息结果类
        private class FileInfoResult
        {
            public string FileFullPath { get; set; }
            public string SheetName { get; set; }
            public string FileNameOnly { get; set; }
        }

        public class SearchResult
        {
            public string TableName { get; set; } // 纯工作表名
            public string ColumnName { get; set; } // 列名
            public int RowNumber { get; set; } // Excel行号（正确的位置）
            public string Value { get; set; } // 匹配的值
            public string FileName { get; set; } // 完整文件路径
            public string FileNameOnly { get; set; } // 仅文件名（不含路径）

            public override string ToString()
            {
                return $"文件: {FileName}, 工作表: {TableName}, 列: {ColumnName}, 行: {RowNumber}, 值: {Value}";
            }

            // 添加格式化方法
            public string ToShortString()
            {
                return $"{FileNameOnly} -> {TableName}[{ColumnName}{RowNumber}] = {Value}";
            }

            // 添加Excel位置信息
            public string ExcelLocation => $"{TableName}!{ColumnName}{RowNumber}";
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
