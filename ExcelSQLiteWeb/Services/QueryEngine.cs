using ExcelSQLiteWeb.Models;
using System.Data;
using System.Diagnostics;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelSQLiteWeb.Services;

/// <summary>
/// 查询引擎 - 执行SQL查询
/// </summary>
public class QueryEngine
{
    private readonly SqliteManager _sqliteManager;

    public QueryEngine(SqliteManager sqliteManager)
    {
        _sqliteManager = sqliteManager;
    }

    /// <summary>
    /// 执行SQL查询
    /// </summary>
    public QueryResult ExecuteQuery(string sql)
    {
        return ExecuteQuery(sql, txn: null, parameters: null);
    }

    /// <summary>
    /// 执行SQL查询（可选事务：用于读取未提交变更）
    /// </summary>
    public QueryResult ExecuteQuery(string sql, Microsoft.Data.Sqlite.SqliteTransaction? txn, object? parameters)
    {
        var stopwatch = Stopwatch.StartNew();
        var result = new QueryResult { Sql = sql };

        try
        {
            var trimmed = (sql ?? string.Empty).TrimStart();
            bool isSelectLike =
                trimmed.StartsWith("select", StringComparison.OrdinalIgnoreCase)
                || trimmed.StartsWith("with", StringComparison.OrdinalIgnoreCase)
                || trimmed.StartsWith("pragma", StringComparison.OrdinalIgnoreCase);

            if (isSelectLike)
            {
                // 使用 DataTable：即使 0 行也能拿到列名（Dapper Query 在 0 行时拿不到 Columns）
                var dt = _sqliteManager.QueryToDataTable(sql, parameters: parameters, txn: txn);
                result.Columns = dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();
                foreach (DataRow row in dt.Rows)
                {
                    var dict = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                    foreach (DataColumn col in dt.Columns)
                    {
                        dict[col.ColumnName] = row[col] ?? DBNull.Value;
                    }
                    result.Rows.Add(dict);
                }
                result.TotalRows = dt.Rows.Count;
            }
            else
            {
                // 非查询：返回影响行数（TotalRows 复用为 affected rows）
                int affected = _sqliteManager.Execute(sql, parameters: parameters, txn: txn);
                result.Columns = new List<string>();
                result.Rows = new List<Dictionary<string, object>>();
                result.TotalRows = affected;
            }

            result.QueryTime = Math.Round(stopwatch.Elapsed.TotalSeconds, 2);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"查询执行失败: {ex.Message}", ex);
        }

        stopwatch.Stop();
        return result;
    }

    /// <summary>
    /// 执行SQL查询（异步，可取消/可设置超时）
    /// </summary>
    public async Task<QueryResult> ExecuteQueryAsync(
        string sql,
        Microsoft.Data.Sqlite.SqliteTransaction? txn = null,
        object? parameters = null,
        int timeoutSeconds = 30,
        CancellationToken cancellationToken = default)
    {
        var stopwatch = Stopwatch.StartNew();
        var result = new QueryResult { Sql = sql };

        try
        {
            var trimmed = (sql ?? string.Empty).TrimStart();
            bool isSelectLike =
                trimmed.StartsWith("select", StringComparison.OrdinalIgnoreCase)
                || trimmed.StartsWith("with", StringComparison.OrdinalIgnoreCase)
                || trimmed.StartsWith("pragma", StringComparison.OrdinalIgnoreCase)
                || trimmed.StartsWith("explain", StringComparison.OrdinalIgnoreCase);

            if (isSelectLike)
            {
                // 性能：避免 DataTable.Load 的额外开销（尤其在分页预览场景）
                var (cols, rows) = await _sqliteManager.QueryToRowsAsync(
                    sql,
                    parameters: parameters,
                    txn: txn,
                    timeoutSeconds: timeoutSeconds,
                    cancellationToken: cancellationToken);
                result.Columns = cols;
                foreach (var r in rows)
                {
                    // QueryResult.Rows 是 List<Dictionary<string, object>>，这里做一次安全转换
                    var dict = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                    foreach (var kv in r)
                    {
                        dict[kv.Key] = kv.Value ?? DBNull.Value;
                    }
                    result.Rows.Add(dict);
                }
                result.TotalRows = rows.Count;
            }
            else
            {
                int affected = await _sqliteManager.ExecuteAsync(sql, parameters: parameters, txn: txn, timeoutSeconds: timeoutSeconds, cancellationToken: cancellationToken);
                result.Columns = new List<string>();
                result.Rows = new List<Dictionary<string, object>>();
                result.TotalRows = affected;
            }

            result.QueryTime = Math.Round(stopwatch.Elapsed.TotalSeconds, 2);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"查询执行失败: {ex.Message}", ex);
        }
        finally
        {
            stopwatch.Stop();
        }

        return result;
    }

    /// <summary>
    /// 执行单表查询
    /// </summary>
    public QueryResult ExecuteSingleTableQuery(
        string tableName,
        List<string> columns,
        List<QueryCondition> conditions,
        string? orderBy = null,
        int? limit = null)
    {
        var sql = BuildSingleTableQuerySql(tableName, columns, conditions, orderBy, limit);
        return ExecuteQuery(sql);
    }

    /// <summary>
    /// 执行多表关联查询
    /// </summary>
    public QueryResult ExecuteJoinQuery(
        string mainTable,
        string joinTable,
        string mainField,
        string joinField,
        List<string> columns,
        string joinType = "INNER")
    {
        var sql = BuildJoinQuerySql(mainTable, joinTable, mainField, joinField, columns, joinType);
        return ExecuteQuery(sql);
    }

    /// <summary>
    /// 执行全局搜索
    /// </summary>
    public List<GlobalSearchResult> ExecuteGlobalSearch(
        string keyword,
        List<string> tables,
        bool searchFieldNames = false,
        bool searchDataContent = true,
        bool exactMatch = false,
        string searchOption = "contains")
    {
        var results = new List<GlobalSearchResult>();

        foreach (var table in tables)
        {
            var schema = _sqliteManager.GetTableSchema(table);
            var columnNames = schema.Select(s => s.ColumnName).ToList();

            // 搜索字段名
            if (searchFieldNames)
            {
                foreach (var col in columnNames)
                {
                    if (MatchKeyword(col, keyword, exactMatch, searchOption))
                    {
                        results.Add(new GlobalSearchResult
                        {
                            SheetName = table,
                            FieldName = col,
                            MatchContent = $"[字段名匹配] {col}",
                            RowNumber = 0,
                            ColumnNumber = 0
                        });
                    }
                }
            }

            // 搜索数据内容
            if (searchDataContent)
            {
                var searchColumns = columnNames
                    .Where(c =>
                    {
                        var low = (c ?? string.Empty).ToLowerInvariant();
                        // 排除明显的 ID/编号列（减少噪声与扫描成本）
                        return !(low.Contains("id") || low.Contains("编号"));
                    })
                    .ToList();

                if (searchColumns.Count > 0)
                {
                    var whereClause = BuildSearchWhereClause(searchColumns, keyword, exactMatch, searchOption);
                    var sql = $"SELECT * FROM {SqliteManager.QuoteIdent(table)} WHERE {whereClause} LIMIT 100";

                    try
                    {
                        var data = _sqliteManager.Query(sql);
                        int rowNum = 1;
                        foreach (var row in data)
                        {
                            foreach (var col in searchColumns)
                            {
                                var value = row[col]?.ToString() ?? string.Empty;
                                if (MatchKeyword(value, keyword, exactMatch, searchOption))
                                {
                                    results.Add(new GlobalSearchResult
                                    {
                                        SheetName = table,
                                        FieldName = col,
                                        MatchContent = value.Length > 100 ? value.Substring(0, 100) + "..." : value,
                                        RowNumber = rowNum,
                                        ColumnNumber = columnNames.IndexOf(col) + 1
                                    });
                                }
                            }
                            rowNum++;
                        }
                    }
                    catch
                    {
                        // 忽略搜索错误
                    }
                }
            }
        }

        return results;
    }

    /// <summary>
    /// 获取表的所有数据
    /// </summary>
    public QueryResult GetAllData(string tableName, int limit = 1000)
    {
        var sql = $"SELECT * FROM {SqliteManager.QuoteIdent(tableName)} LIMIT {limit}";
        return ExecuteQuery(sql);
    }

    /// <summary>
    /// 获取表的字段列表
    /// </summary>
    public List<string> GetTableColumns(string tableName)
    {
        var schema = _sqliteManager.GetTableSchema(tableName);
        return schema.Select(s => s.ColumnName).ToList();
    }

    /// <summary>
    /// 获取字段的唯一值
    /// </summary>
    public List<string> GetDistinctValues(string tableName, string columnName, int limit = 100)
    {
        return _sqliteManager.GetUniqueValues(tableName, columnName, limit);
    }

    /// <summary>
    /// 构建单表查询SQL
    /// </summary>
    private string BuildSingleTableQuerySql(
        string tableName,
        List<string> columns,
        List<QueryCondition> conditions,
        string? orderBy = null,
        int? limit = null)
    {
        var sql = new StringBuilder();

        // SELECT 子句
        var columnList = columns.Count > 0
            ? string.Join(", ", columns.Select(c => $"{SqliteManager.QuoteIdent(c)}"))
            : "*";
        sql.Append($"SELECT {columnList} FROM {SqliteManager.QuoteIdent(tableName)}");

        // WHERE 子句
        if (conditions.Count > 0)
        {
            sql.Append(" WHERE ");
            for (int i = 0; i < conditions.Count; i++)
            {
                var condition = conditions[i];
                if (i > 0)
                {
                    sql.Append($" {condition.LogicOperator} ");
                }
                sql.Append(BuildConditionSql(condition));
            }
        }

        // ORDER BY 子句
        if (!string.IsNullOrWhiteSpace(orderBy))
        {
            sql.Append($" ORDER BY {SqliteManager.QuoteIdent(orderBy)}");
        }

        // LIMIT 子句
        if (limit.HasValue && limit.Value > 0)
        {
            sql.Append($" LIMIT {limit.Value}");
        }

        return sql.ToString();
    }

    /// <summary>
    /// 构建关联查询SQL
    /// </summary>
    private string BuildJoinQuerySql(
        string mainTable,
        string joinTable,
        string mainField,
        string joinField,
        List<string> columns,
        string joinType = "INNER")
    {
        var columnList = columns.Count > 0
            ? string.Join(", ", columns.Select(c => $"{SqliteManager.QuoteIdent(c)}"))
            : "*";

        var sql = $"SELECT {columnList} FROM {SqliteManager.QuoteIdent(mainTable)} " +
                  $"{joinType} JOIN {SqliteManager.QuoteIdent(joinTable)} " +
                  $"ON {SqliteManager.QuoteIdent(mainTable)}.{SqliteManager.QuoteIdent(mainField)} = " +
                  $"{SqliteManager.QuoteIdent(joinTable)}.{SqliteManager.QuoteIdent(joinField)}";

        return sql;
    }

    /// <summary>
    /// 构建条件SQL
    /// </summary>
    private string BuildConditionSql(QueryCondition condition)
    {
        var field = condition.FieldName ?? "";
        var op = condition.Operator.ToUpper();
        var value = condition.Value.Replace("'", "''");

        return op switch
        {
            "=" or "!=" or "<>" or ">" or "<" or ">=" or "<=" =>
                $"{SqliteManager.QuoteIdent(field)} {op} '{value}'",
            "LIKE" =>
                $"{SqliteManager.QuoteIdent(field)} LIKE '%{value}%'",
            "NOT LIKE" =>
                $"{SqliteManager.QuoteIdent(field)} NOT LIKE '%{value}%'",
            "IN" =>
                $"{SqliteManager.QuoteIdent(field)} IN ({value})",
            "NOT IN" =>
                $"{SqliteManager.QuoteIdent(field)} NOT IN ({value})",
            "IS NULL" =>
                $"{SqliteManager.QuoteIdent(field)} IS NULL",
            "IS NOT NULL" =>
                $"{SqliteManager.QuoteIdent(field)} IS NOT NULL",
            "BETWEEN" =>
                $"{SqliteManager.QuoteIdent(field)} BETWEEN {value}",
            _ => $"{SqliteManager.QuoteIdent(field)} = '{value}'"
        };
    }

    /// <summary>
    /// 构建搜索WHERE子句
    /// </summary>
    private string BuildSearchWhereClause(List<string> columns, string keyword, bool exactMatch, string searchOption)
    {
        var conditions = new List<string>();
        var escapedKeyword = keyword.Replace("'", "''");
        var opt = (searchOption ?? "contains")
            .Trim()
            .Replace("_", "-")
            .ToLowerInvariant();
        // 兼容多种前端取值：starts-with / startswith / begins-with / equals / exact / contains
        if (opt == "starts-with" || opt == "begin" || opt == "begins-with") opt = "startswith";
        if (opt == "ends-with") opt = "endswith";
        if (opt == "equals") opt = "exact";
        if (opt == "contains") opt = "contains";

        foreach (var col in columns)
        {
            var condition = opt switch
            {
                "startswith" => $"{SqliteManager.QuoteIdent(col)} LIKE '{escapedKeyword}%'",
                "endswith" => $"{SqliteManager.QuoteIdent(col)} LIKE '%{escapedKeyword}'",
                "exact" => $"{SqliteManager.QuoteIdent(col)} = '{escapedKeyword}'",
                _ => $"{SqliteManager.QuoteIdent(col)} LIKE '%{escapedKeyword}%'"
            };
            conditions.Add(condition);
        }

        return string.Join(" OR ", conditions);
    }

    /// <summary>
    /// 匹配关键字
    /// </summary>
    private bool MatchKeyword(string text, string keyword, bool exactMatch, string searchOption)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;

        var textLower = text.ToLower();
        var keywordLower = keyword.ToLower();

        if (exactMatch)
        {
            return textLower == keywordLower;
        }

        var opt = (searchOption ?? "contains")
            .Trim()
            .Replace("_", "-")
            .ToLowerInvariant();
        if (opt == "starts-with" || opt == "begin" || opt == "begins-with") opt = "startswith";
        if (opt == "ends-with") opt = "endswith";
        if (opt == "equals") opt = "exact";
        if (opt == "contains") opt = "contains";

        return opt switch
        {
            "startswith" => textLower.StartsWith(keywordLower),
            "endswith" => textLower.EndsWith(keywordLower),
            "exact" => textLower == keywordLower,
            _ => textLower.Contains(keywordLower)
        };
    }

}
