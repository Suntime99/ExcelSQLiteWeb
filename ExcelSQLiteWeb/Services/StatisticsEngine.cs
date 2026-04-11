using ExcelSQLiteWeb.Models;
using System.Diagnostics;
using System.Text;

namespace ExcelSQLiteWeb.Services;

/// <summary>
/// 统计引擎 - 执行数据统计分析
/// </summary>
public class StatisticsEngine
{
    private readonly SqliteManager _sqliteManager;

    public StatisticsEngine(SqliteManager sqliteManager)
    {
        _sqliteManager = sqliteManager;
    }

    /// <summary>
    /// 执行单表统计
    /// </summary>
    public StatisticsResult ExecuteSingleTableStatistics(
        string tableName,
        List<string> groupByFields,
        List<StatisticsMetric> metrics)
    {
        var stopwatch = Stopwatch.StartNew();
        var sql = BuildStatisticsSql(tableName, groupByFields, metrics);

        var result = new StatisticsResult
        {
            Sql = sql,
            GroupByFields = groupByFields,
            Metrics = metrics.Select(m => $"{m.AggregateFunction}({m.FieldName})").ToList()
        };

        try
        {
            var data = _sqliteManager.Query(sql);

            if (data.Count > 0)
            {
                result.Columns = data[0].Keys.ToList();
                result.Rows = data;
                result.TotalRows = data.Count;
            }

            result.StatisticsTime = Math.Round(stopwatch.Elapsed.TotalSeconds, 2);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"统计执行失败: {ex.Message}", ex);
        }

        stopwatch.Stop();
        return result;
    }

    /// <summary>
    /// 执行多表关联统计
    /// </summary>
    public StatisticsResult ExecuteMultiTableStatistics(
        string mainTable,
        string joinTable,
        string mainField,
        string joinField,
        List<string> groupByFields,
        List<StatisticsMetric> metrics,
        string joinType = "INNER")
    {
        var stopwatch = Stopwatch.StartNew();
        var sql = BuildMultiTableStatisticsSql(mainTable, joinTable, mainField, joinField,
            groupByFields, metrics, joinType);

        var result = new StatisticsResult
        {
            Sql = sql,
            GroupByFields = groupByFields,
            Metrics = metrics.Select(m => $"{m.AggregateFunction}({m.FieldName})").ToList()
        };

        try
        {
            var data = _sqliteManager.Query(sql);

            if (data.Count > 0)
            {
                result.Columns = data[0].Keys.ToList();
                result.Rows = data;
                result.TotalRows = data.Count;
            }

            result.StatisticsTime = Math.Round(stopwatch.Elapsed.TotalSeconds, 2);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"统计执行失败: {ex.Message}", ex);
        }

        stopwatch.Stop();
        return result;
    }

    /// <summary>
    /// 获取字段的基本统计信息
    /// </summary>
    public Dictionary<string, object> GetFieldStatistics(string tableName, string fieldName)
    {
        var sql = $@"SELECT
            COUNT(*) as TotalCount,
            COUNT(DISTINCT [{SanitizeIdentifier(fieldName)}]) as UniqueCount,
            COUNT(CASE WHEN [{SanitizeIdentifier(fieldName)}] IS NULL THEN 1 END) as NullCount,
            MIN([{SanitizeIdentifier(fieldName)}]) as MinValue,
            MAX([{SanitizeIdentifier(fieldName)}]) as MaxValue
        FROM [{SanitizeIdentifier(tableName)}]";

        var result = _sqliteManager.Query(sql);
        return result.FirstOrDefault() ?? new Dictionary<string, object>();
    }

    /// <summary>
    /// 获取数值字段的统计信息
    /// </summary>
    public Dictionary<string, object> GetNumericFieldStatistics(string tableName, string fieldName)
    {
        var sql = $@"SELECT
            COUNT(*) as TotalCount,
            COUNT(DISTINCT [{SanitizeIdentifier(fieldName)}]) as UniqueCount,
            COUNT(CASE WHEN [{SanitizeIdentifier(fieldName)}] IS NULL THEN 1 END) as NullCount,
            MIN([{SanitizeIdentifier(fieldName)}]) as MinValue,
            MAX([{SanitizeIdentifier(fieldName)}]) as MaxValue,
            AVG(CAST([{SanitizeIdentifier(fieldName)}] AS REAL)) as AvgValue,
            SUM(CAST([{SanitizeIdentifier(fieldName)}] AS REAL)) as SumValue
        FROM [{SanitizeIdentifier(tableName)}]";

        var result = _sqliteManager.Query(sql);
        return result.FirstOrDefault() ?? new Dictionary<string, object>();
    }

    /// <summary>
    /// 获取分组统计（简单版）
    /// </summary>
    public StatisticsResult GetGroupStatistics(string tableName, string groupByField, string aggregateField, string aggregateFunction)
    {
        var metrics = new List<StatisticsMetric>
        {
            new() { FieldName = aggregateField, AggregateFunction = aggregateFunction }
        };

        return ExecuteSingleTableStatistics(tableName, new List<string> { groupByField }, metrics);
    }

    /// <summary>
    /// 构建统计SQL
    /// </summary>
    private string BuildStatisticsSql(
        string tableName,
        List<string> groupByFields,
        List<StatisticsMetric> metrics)
    {
        var sql = new StringBuilder();

        // SELECT 子句
        var selectParts = new List<string>();

        // 添加分组字段
        foreach (var field in groupByFields)
        {
            selectParts.Add($"[{SanitizeIdentifier(field)}]");
        }

        // 添加统计指标
        foreach (var metric in metrics)
        {
            var field = SanitizeIdentifier(metric.FieldName);
            var func = metric.AggregateFunction.ToUpper();
            var alias = string.IsNullOrWhiteSpace(metric.Alias)
                ? $"{func}_{field}"
                : SanitizeIdentifier(metric.Alias);

            selectParts.Add($"{func}([{field}]) AS [{alias}]");
        }

        sql.Append($"SELECT {string.Join(", ", selectParts)} ");
        sql.Append($"FROM [{SanitizeIdentifier(tableName)}]");

        // GROUP BY 子句
        if (groupByFields.Count > 0)
        {
            var groupByClause = string.Join(", ", groupByFields.Select(f => $"[{SanitizeIdentifier(f)}]"));
            sql.Append($" GROUP BY {groupByClause}");
        }

        // 添加排序（按第一个统计指标降序）
        if (metrics.Count > 0)
        {
            var firstMetric = metrics[0];
            var alias = string.IsNullOrWhiteSpace(firstMetric.Alias)
                ? $"{firstMetric.AggregateFunction.ToUpper()}_{SanitizeIdentifier(firstMetric.FieldName)}"
                : SanitizeIdentifier(firstMetric.Alias);
            sql.Append($" ORDER BY [{alias}] DESC");
        }

        return sql.ToString();
    }

    /// <summary>
    /// 构建多表统计SQL
    /// </summary>
    private string BuildMultiTableStatisticsSql(
        string mainTable,
        string joinTable,
        string mainField,
        string joinField,
        List<string> groupByFields,
        List<StatisticsMetric> metrics,
        string joinType = "INNER")
    {
        var sql = new StringBuilder();

        // SELECT 子句
        var selectParts = new List<string>();

        // 添加分组字段
        foreach (var field in groupByFields)
        {
            selectParts.Add($"[{SanitizeIdentifier(field)}]");
        }

        // 添加统计指标
        foreach (var metric in metrics)
        {
            var field = SanitizeIdentifier(metric.FieldName);
            var func = metric.AggregateFunction.ToUpper();
            var alias = string.IsNullOrWhiteSpace(metric.Alias)
                ? $"{func}_{field}"
                : SanitizeIdentifier(metric.Alias);

            selectParts.Add($"{func}([{field}]) AS [{alias}]");
        }

        sql.Append($"SELECT {string.Join(", ", selectParts)} ");
        sql.Append($"FROM [{SanitizeIdentifier(mainTable)}] ");
        sql.Append($"{joinType} JOIN [{SanitizeIdentifier(joinTable)}] ");
        sql.Append($"ON [{SanitizeIdentifier(mainTable)}].[{SanitizeIdentifier(mainField)}] = ");
        sql.Append($"[{SanitizeIdentifier(joinTable)}].[{SanitizeIdentifier(joinField)}]");

        // GROUP BY 子句
        if (groupByFields.Count > 0)
        {
            var groupByClause = string.Join(", ", groupByFields.Select(f => $"[{SanitizeIdentifier(f)}]"));
            sql.Append($" GROUP BY {groupByClause}");
        }

        // 添加排序
        if (metrics.Count > 0)
        {
            var firstMetric = metrics[0];
            var alias = string.IsNullOrWhiteSpace(firstMetric.Alias)
                ? $"{firstMetric.AggregateFunction.ToUpper()}_{SanitizeIdentifier(firstMetric.FieldName)}"
                : SanitizeIdentifier(firstMetric.Alias);
            sql.Append($" ORDER BY [{alias}] DESC");
        }

        return sql.ToString();
    }

    /// <summary>
    /// 清理标识符
    /// </summary>
    private string SanitizeIdentifier(string identifier)
    {
        if (string.IsNullOrWhiteSpace(identifier))
            return "_";

        var sanitized = new string(identifier.Where(c => char.IsLetterOrDigit(c) || c == '_').ToArray());

        if (sanitized.Length > 0 && char.IsDigit(sanitized[0]))
            sanitized = "_" + sanitized;

        return string.IsNullOrEmpty(sanitized) ? "_" : sanitized;
    }
}
