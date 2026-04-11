using System.Data;
using Microsoft.Data.Sqlite;
using Dapper;
using System.Threading;
using System.Threading.Tasks;
using System.Reflection;

namespace ExcelSQLiteWeb.Services;

/// <summary>
/// SQLite数据库管理器
/// </summary>
public class SqliteManager : IDisposable
{
    private SqliteConnection? _connection;
    private readonly string _connectionString;
    private bool _disposed;

    /// <summary>
    /// 连接字符串（用于创建新的只读连接/后台任务等）
    /// </summary>
    public string ConnectionString => _connectionString;

    /// <summary>
    /// 当前连接（注意：不是线程安全的，仅用于主线程/需要备份时使用）
    /// </summary>
    public SqliteConnection? Connection => _connection;

    /// <summary>
    /// 是否已连接
    /// </summary>
    public bool IsConnected => _connection?.State == ConnectionState.Open;

    /// <summary>
    /// 内存数据库连接字符串
    /// </summary>
    public SqliteManager()
    {
        _connectionString = "Data Source=:memory:";
    }

    /// <summary>
    /// 文件数据库连接字符串
    /// </summary>
    public SqliteManager(string dbPath)
    {
        _connectionString = $"Data Source={dbPath}";
    }

    /// <summary>
    /// SQLite 标识符引用：使用 [] 包裹，内部 ] 转义为 ]]
    /// 重要：不要“清理/改写”表名/列名，否则会导致“建表名”和“插入名”不一致。
    /// </summary>
    public static string QuoteIdent(string name)
    {
        var n = (name ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(n)) n = "_";
        // 支持 schema.table（例如：base.Customer）：分别引用每一段
        if (n.Contains('.') && !n.Contains('[') && !n.Contains(']'))
        {
            var parts = n.Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            if (parts.Length >= 2)
            {
                return string.Join(".", parts.Select(p =>
                {
                    var x = (p ?? "_").Trim();
                    if (string.IsNullOrWhiteSpace(x)) x = "_";
                    x = x.Replace("]", "]]");
                    return $"[{x}]";
                }));
            }
        }
        n = n.Replace("]", "]]");
        return $"[{n}]";
    }

    private static string NormalizeIdent(string name)
    {
        var n = (name ?? string.Empty).Trim();
        return string.IsNullOrWhiteSpace(n) ? "_" : n;
    }

    /// <summary>
    /// 打开连接
    /// </summary>
    public void Open()
    {
        if (_connection == null)
        {
            _connection = new SqliteConnection(_connectionString);
            _connection.Open();
        }
        else if (_connection.State != ConnectionState.Open)
        {
            _connection.Open();
        }
    }

    /// <summary>
    /// 关闭连接
    /// </summary>
    public void Close()
    {
        _connection?.Close();
    }

    /// <summary>
    /// 关闭并重建连接（用于“强制恢复/重建连接”）。
    /// 注意：若是内存库（:memory:），重建会丢失数据；当前项目默认使用文件库。
    /// </summary>
    public void Reopen()
    {
        try { _connection?.Close(); } catch { }
        try { _connection?.Dispose(); } catch { }
        _connection = null;
        Open();
    }

    /// <summary>
    /// 创建表
    /// </summary>
    public void CreateTable(string tableName, List<(string ColumnName, string DataType)> columns)
        => CreateTable(tableName, columns, dropIfExists: true);

    /// <summary>
    /// 创建表（可控制是否删除已存在的表）
    /// - dropIfExists=true：覆盖导入（默认行为）
    /// - dropIfExists=false：如果表不存在则创建；若已存在则保持不变（用于追加导入）
    /// </summary>
    public void CreateTable(string tableName, List<(string ColumnName, string DataType)> columns, bool dropIfExists)
    {
        EnsureConnected();

        // 保留原表名（支持中文/符号），仅做最小归一化
        tableName = NormalizeIdent(tableName);

        if (dropIfExists)
        {
            // 删除已存在的表
            Execute($"DROP TABLE IF EXISTS {QuoteIdent(tableName)}");
        }
        else
        {
            // 追加导入：表已存在则直接复用
            if (TableExists(tableName)) return;
        }

        // 构建列定义
        var columnDefs = columns.Select(c =>
        {
            string colName = NormalizeIdent(c.ColumnName);
            string sqlType = MapToSqliteType(c.DataType);
            return $"{QuoteIdent(colName)} {sqlType}";
        });

        string createSql = $"CREATE TABLE {QuoteIdent(tableName)} ({string.Join(", ", columnDefs)})";
        Execute(createSql);
    }

    /// <summary>
    /// 表是否存在（table 或 view）
    /// </summary>
    public bool TableExists(string tableName)
    {
        EnsureConnected();
        tableName = NormalizeIdent(tableName);
        using var cmd = _connection!.CreateCommand();
        cmd.CommandText = "SELECT 1 FROM sqlite_master WHERE (type='table' OR type='view') AND name=@name LIMIT 1;";
        cmd.Parameters.AddWithValue("@name", tableName);
        var x = cmd.ExecuteScalar();
        return x != null && x != DBNull.Value;
    }

    /// <summary>
    /// 获取表字段名清单（按 PRAGMA table_info）
    /// </summary>
    public List<string> GetTableColumns(string tableName)
    {
        EnsureConnected();
        tableName = NormalizeIdent(tableName);
        var list = new List<string>();
        using var cmd = _connection!.CreateCommand();
        // 支持 schema.table：PRAGMA schema.table_info(table)
        if (tableName.Contains('.') && !tableName.Contains('[') && !tableName.Contains(']'))
        {
            var parts = tableName.Split('.', 2, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            var schema = parts.Length == 2 ? parts[0] : "main";
            var tn = parts.Length == 2 ? parts[1] : tableName;
            cmd.CommandText = $"PRAGMA {QuoteIdent(schema)}.table_info({QuoteIdent(tn)});";
        }
        else
        {
            cmd.CommandText = $"PRAGMA table_info({QuoteIdent(tableName)});";
        }
        using var rd = cmd.ExecuteReader();
        while (rd.Read())
        {
            // PRAGMA table_info: cid, name, type, notnull, dflt_value, pk
            var n = rd.IsDBNull(1) ? "" : (rd.GetString(1) ?? "");
            if (!string.IsNullOrWhiteSpace(n)) list.Add(n);
        }
        return list;
    }

    /// <summary>
    /// 插入数据
    /// </summary>
    public void InsertData(string tableName, List<Dictionary<string, object>> rows)
    {
        if (rows.Count == 0) return;

        EnsureConnected();
        tableName = NormalizeIdent(tableName);

        using var transaction = _connection!.BeginTransaction();
        try
        {
            var columnNames = rows[0].Keys.ToList();
            var columns = string.Join(", ", columnNames.Select(c => QuoteIdent(NormalizeIdent(c))));
            var parameters = string.Join(", ", columnNames.Select((_, i) => $"@p{i}"));

            string insertSql = $"INSERT INTO {QuoteIdent(tableName)} ({columns}) VALUES ({parameters})";

            foreach (var row in rows)
            {
                var paramDict = new DynamicParameters();
                for (int i = 0; i < columnNames.Count; i++)
                {
                    var value = row[columnNames[i]];
                    paramDict.Add($"@p{i}", value == DBNull.Value ? null : value);
                }

                _connection.Execute(insertSql, paramDict, transaction);
            }

            transaction.Commit();
        }
        catch
        {
            transaction.Rollback();
            throw;
        }
    }

    /// <summary>
    /// 批量插入数据（使用事务优化）
    /// </summary>
    public void BulkInsert(string tableName, List<Dictionary<string, object>> rows, int batchSize = 1000)
    {
        if (rows.Count == 0) return;

        EnsureConnected();
        tableName = NormalizeIdent(tableName);

        // 分批处理
        for (int i = 0; i < rows.Count; i += batchSize)
        {
            var batch = rows.Skip(i).Take(batchSize).ToList();
            InsertData(tableName, batch);
        }
    }

    /// <summary>
    /// 执行SQL查询
    /// </summary>
    public List<Dictionary<string, object>> Query(string sql, object? parameters = null, SqliteTransaction? txn = null)
    {
        EnsureConnected();
        // 参数绑定：若传入 IDictionary，则走 SqliteCommand（避免 Dapper 对字典参数兼容性差异）
        if (parameters is IDictionary<string, object?>)
        {
            var list = new List<Dictionary<string, object>>();
            using var command = new SqliteCommand(sql, _connection);
            if (txn != null) command.Transaction = txn;
            AddParameters(command, parameters);
            using var reader = command.ExecuteReader();
            var cols = new List<string>();
            for (int i = 0; i < reader.FieldCount; i++) cols.Add(reader.GetName(i));
            while (reader.Read())
            {
                var dict = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < cols.Count; i++)
                {
                    dict[cols[i]] = reader.IsDBNull(i) ? DBNull.Value : (reader.GetValue(i) ?? DBNull.Value);
                }
                list.Add(dict);
            }
            return list;
        }

        var result = _connection!.Query(sql, parameters, txn);
        return result.Select(row =>
        {
            var dict = new Dictionary<string, object>();
            var rowDict = (IDictionary<string, object>)row;
            foreach (var kvp in rowDict)
            {
                dict[kvp.Key] = kvp.Value ?? DBNull.Value;
            }
            return dict;
        }).ToList();
    }

    /// <summary>
    /// 执行SQL查询并返回DataTable
    /// </summary>
    public DataTable QueryToDataTable(string sql, object? parameters = null, SqliteTransaction? txn = null)
    {
        EnsureConnected();

        var dataTable = new DataTable();
        using var command = new SqliteCommand(sql, _connection);
        if (txn != null) command.Transaction = txn;

        if (parameters != null)
        {
            AddParameters(command, parameters);
        }

        using var reader = command.ExecuteReader();
        dataTable.Load(reader);
        return dataTable;
    }

    /// <summary>
    /// 执行SQL查询并返回DataTable（异步，可取消/可设置超时）
    /// </summary>
    public async Task<DataTable> QueryToDataTableAsync(
        string sql,
        object? parameters = null,
        SqliteTransaction? txn = null,
        int timeoutSeconds = 30,
        CancellationToken cancellationToken = default)
    {
        EnsureConnected();

        var dataTable = new DataTable();
        using var command = new SqliteCommand(sql, _connection);
        if (txn != null) command.Transaction = txn;
        if (timeoutSeconds > 0) command.CommandTimeout = timeoutSeconds;

        if (parameters != null)
        {
            AddParameters(command, parameters);
        }

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        dataTable.Load(reader);
        return dataTable;
    }

    /// <summary>
    /// 执行SQL查询并返回“列名 + 行字典列表”（异步，可取消/可设置超时）。
    /// 性能优于 DataTable.Load：减少中间对象与二次拷贝，适合 SQL 实验室分页预览。
    /// </summary>
    public async Task<(List<string> Columns, List<Dictionary<string, object?>> Rows)> QueryToRowsAsync(
        string sql,
        object? parameters = null,
        SqliteTransaction? txn = null,
        int timeoutSeconds = 30,
        CancellationToken cancellationToken = default)
    {
        EnsureConnected();

        using var command = new SqliteCommand(sql, _connection);
        if (txn != null) command.Transaction = txn;
        if (timeoutSeconds > 0) command.CommandTimeout = timeoutSeconds;
        if (parameters != null) AddParameters(command, parameters);

        var rows = new List<Dictionary<string, object?>>();
        var columns = new List<string>();

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        for (int i = 0; i < reader.FieldCount; i++)
        {
            columns.Add(reader.GetName(i));
        }

        while (await reader.ReadAsync(cancellationToken))
        {
            var dict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < columns.Count; i++)
            {
                dict[columns[i]] = reader.IsDBNull(i) ? null : reader.GetValue(i);
            }
            rows.Add(dict);
        }

        return (columns, rows);
    }

    /// <summary>
    /// 执行非查询SQL
    /// </summary>
    public int Execute(string sql, object? parameters = null, SqliteTransaction? txn = null)
    {
        EnsureConnected();
        // 参数绑定：若传入 IDictionary，则走 SqliteCommand（避免 Dapper 对字典参数兼容性差异）
        if (parameters is IDictionary<string, object?>)
        {
            using var command = new SqliteCommand(sql, _connection);
            if (txn != null) command.Transaction = txn;
            AddParameters(command, parameters);
            return command.ExecuteNonQuery();
        }
        return _connection!.Execute(sql, parameters, txn);
    }

    /// <summary>
    /// 执行非查询SQL（异步，可取消/可设置超时）
    /// </summary>
    public async Task<int> ExecuteAsync(
        string sql,
        object? parameters = null,
        SqliteTransaction? txn = null,
        int timeoutSeconds = 30,
        CancellationToken cancellationToken = default)
    {
        EnsureConnected();
        using var command = new SqliteCommand(sql, _connection);
        if (txn != null) command.Transaction = txn;
        if (timeoutSeconds > 0) command.CommandTimeout = timeoutSeconds;

        if (parameters != null)
        {
            AddParameters(command, parameters);
        }

        return await command.ExecuteNonQueryAsync(cancellationToken);
    }

    /// <summary>
    /// 获取标量值
    /// </summary>
    public T? ExecuteScalar<T>(string sql, object? parameters = null)
    {
        EnsureConnected();
        return _connection!.ExecuteScalar<T>(sql, parameters);
    }

    /// <summary>
    /// 获取表列表
    /// </summary>
    public List<string> GetTables()
    {
        // main + attached DB（例如 base）统一返回：
        // - main: TableName
        // - attached: alias.TableName
        var list = new List<string>();
        try
        {
            var dbs = Query("PRAGMA database_list;")
                .Select(r => Convert.ToString(r["name"]) ?? "")
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var db in dbs)
            {
                if (string.Equals(db, "temp", StringComparison.OrdinalIgnoreCase)) continue;
                var sql = $"SELECT name, type FROM {QuoteIdent(db)}.sqlite_master WHERE (type='table' OR type='view') AND name NOT LIKE 'sqlite_%' ORDER BY type, name;";
                var rows = Query(sql);
                foreach (var r in rows)
                {
                    var name = Convert.ToString(r["name"]) ?? "";
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    list.Add(string.Equals(db, "main", StringComparison.OrdinalIgnoreCase) ? name : $"{db}.{name}");
                }
            }
        }
        catch
        {
            // 回退：仅 main
            const string sql = "SELECT name FROM sqlite_master WHERE (type='table' OR type='view') AND name NOT LIKE 'sqlite_%'";
            list = Query(sql).Select(r => r["name"].ToString()!).ToList();
        }

        return list.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    public void AttachDatabase(string dbPath, string alias)
    {
        EnsureConnected();
        if (string.IsNullOrWhiteSpace(dbPath)) throw new ArgumentException("dbPath为空");
        if (string.IsNullOrWhiteSpace(alias)) throw new ArgumentException("alias为空");
        var path = dbPath.Replace("'", "''");
        Execute($"ATTACH DATABASE '{path}' AS {QuoteIdent(alias)};");
    }

    public void DetachDatabase(string alias)
    {
        EnsureConnected();
        if (string.IsNullOrWhiteSpace(alias)) return;
        Execute($"DETACH DATABASE {QuoteIdent(alias)};");
    }

    /// <summary>
    /// 获取表结构
    /// </summary>
    public List<(string ColumnName, string DataType)> GetTableSchema(string tableName)
    {
        EnsureConnected();
        tableName = NormalizeIdent(tableName);
        // 支持 schema.table：PRAGMA schema.table_info(table)
        string sql;
        if (tableName.Contains('.') && !tableName.Contains('[') && !tableName.Contains(']'))
        {
            var parts = tableName.Split('.', 2, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            var schema = parts.Length == 2 ? parts[0] : "main";
            var tn = parts.Length == 2 ? parts[1] : tableName;
            sql = $"PRAGMA {QuoteIdent(schema)}.table_info({QuoteIdent(tn)})";
        }
        else
        {
            sql = $"PRAGMA table_info({QuoteIdent(tableName)})";
        }
        var result = Query(sql);

        return result.Select(r => (
            r["name"].ToString()!,
            r["type"].ToString()!
        )).ToList();
    }

    /// <summary>
    /// 获取表的行数
    /// </summary>
    public int GetRowCount(string tableName)
    {
        EnsureConnected();
        tableName = NormalizeIdent(tableName);

        var sql = $"SELECT COUNT(*) FROM {QuoteIdent(tableName)}";
        return ExecuteScalar<int>(sql);
    }

    /// <summary>
    /// 删除表
    /// </summary>
    public void DropTable(string tableName)
    {
        EnsureConnected();
        tableName = NormalizeIdent(tableName);
        Execute($"DROP TABLE IF EXISTS {QuoteIdent(tableName)}");
    }

    /// <summary>
    /// 清空表
    /// </summary>
    public void TruncateTable(string tableName)
    {
        EnsureConnected();
        tableName = NormalizeIdent(tableName);
        Execute($"DELETE FROM {QuoteIdent(tableName)}");
    }

    /// <summary>
    /// 创建索引
    /// </summary>
    public void CreateIndex(string tableName, string columnName, string? indexName = null)
    {
        EnsureConnected();
        tableName = NormalizeIdent(tableName);
        columnName = NormalizeIdent(columnName);
        indexName ??= $"idx_{tableName}_{columnName}";
        indexName = NormalizeIdent(indexName);

        var sql = $"CREATE INDEX IF NOT EXISTS {QuoteIdent(indexName)} ON {QuoteIdent(tableName)} ({QuoteIdent(columnName)})";
        Execute(sql);
    }

    /// <summary>
    /// 获取字段的唯一值
    /// </summary>
    public List<string> GetUniqueValues(string tableName, string columnName, int limit = 1000)
    {
        EnsureConnected();
        tableName = NormalizeIdent(tableName);
        columnName = NormalizeIdent(columnName);

        var sql = $"SELECT DISTINCT {QuoteIdent(columnName)} FROM {QuoteIdent(tableName)} WHERE {QuoteIdent(columnName)} IS NOT NULL LIMIT {limit}";
        var result = Query(sql);
        return result.Select(r => r[columnName]?.ToString() ?? string.Empty)
                     .Where(v => !string.IsNullOrEmpty(v))
                     .ToList();
    }

    /// <summary>
    /// 获取数据库统计信息
    /// </summary>
    public (int TableCount, int TotalRows, long Size) GetDatabaseStats()
    {
        EnsureConnected();

        var tables = GetTables();
        int totalRows = tables.Sum(t => GetRowCount(t));

        return (tables.Count, totalRows, 0);
    }

    /// <summary>
    /// 将内存数据库保存到文件
    /// </summary>
    public void SaveToFile(string filePath)
    {
        EnsureConnected();

        // 使用VACUUM INTO保存内存数据库到文件
        var sql = $"VACUUM INTO '{filePath.Replace("'", "''")}'";
        Execute(sql);
    }

    /// <summary>
    /// 从文件加载数据库
    /// </summary>
    public void LoadFromFile(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("数据库文件不存在", filePath);

        EnsureConnected();

        // 附加文件数据库
        var attachSql = $"ATTACH DATABASE '{filePath.Replace("'", "''")}' AS source";
        Execute(attachSql);

        // 复制所有表
        var tables = Query("SELECT name FROM source.sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'");
        foreach (var table in tables)
        {
            var tableName = table["name"].ToString()!;
            Execute($"CREATE TABLE {QuoteIdent(tableName)} AS SELECT * FROM source.{QuoteIdent(tableName)}");
        }

        // 分离文件数据库
        Execute("DETACH DATABASE source");
    }

    /// <summary>
    /// 确保已连接
    /// </summary>
    private void EnsureConnected()
    {
        if (_connection == null || _connection.State != ConnectionState.Open)
        {
            Open();
        }
    }

    /// <summary>
    /// 尝试中断当前正在执行的 SQL（用于“取消执行”硬中止）。
    /// 注意：不同版本的 Microsoft.Data.Sqlite 可能实现不同，这里用反射调用 Interrupt() 以保持兼容性。
    /// </summary>
    public void Interrupt()
    {
        try
        {
            EnsureConnected();
            var conn = _connection;
            if (conn == null) return;
            var m = conn.GetType().GetMethod("Interrupt", BindingFlags.Instance | BindingFlags.Public);
            if (m != null)
            {
                m.Invoke(conn, null);
            }
        }
        catch
        {
            // ignore
        }
    }

    /// <summary>
    /// 添加参数
    /// </summary>
    private void AddParameters(SqliteCommand command, object parameters)
    {
        if (parameters is DynamicParameters dapperParams)
        {
            // Dapper参数已在查询中处理
            return;
        }

        if (parameters is IDictionary<string, object?> dict)
        {
            foreach (var kv in dict)
            {
                var key = (kv.Key ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(key)) continue;
                var val = kv.Value ?? DBNull.Value;

                // 支持 :p / @p / $p 三种命名
                void AddOne(string name)
                {
                    if (string.IsNullOrWhiteSpace(name)) return;
                    if (command.Parameters.Contains(name)) return;
                    command.Parameters.AddWithValue(name, val);
                }

                var bare = key.TrimStart(':', '@', '$');
                AddOne("@" + bare);
                AddOne(":" + bare);
                AddOne("$" + bare);
            }
            return;
        }

        var properties = parameters.GetType().GetProperties();
        foreach (var prop in properties)
        {
            var value = prop.GetValue(parameters) ?? DBNull.Value;
            command.Parameters.AddWithValue($"@{prop.Name}", value);
        }
    }

    /// <summary>
    /// 映射数据类型到SQLite类型
    /// </summary>
    private string MapToSqliteType(string dataType)
    {
        return dataType.ToLower() switch
        {
            "number" or "numeric" or "decimal" or "float" or "double" => "REAL",
            "int" or "integer" or "long" or "short" or "byte" => "INTEGER",
            "datetime" or "date" or "time" => "TEXT",
            "boolean" or "bool" => "INTEGER",
            "blob" or "binary" => "BLOB",
            _ => "TEXT"
        };
    }

    // 注意：不要再做“删除非法字符”的 SanitizeIdentifier，否则会导致：
    // 1) CreateTable 使用清理后的表名建表
    // 2) Insert / Query 使用原表名访问
    // => 最终出现 “no such table” 的 SqliteException

    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        if (!_disposed)
        {
            _connection?.Dispose();
            _disposed = true;
        }
        GC.SuppressFinalize(this);
    }
}
