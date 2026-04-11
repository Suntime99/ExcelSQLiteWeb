using ExcelSQLiteWeb.Models;
using System.Text;

namespace ExcelSQLiteWeb.Services;

/// <summary>
/// SQL脚本生成器
/// </summary>
public class SqlScriptGenerator
{
    /// <summary>
    /// 生成查询脚本
    /// </summary>
    public string GenerateQueryScript(
        string tableName,
        List<string> columns,
        List<QueryCondition> conditions,
        string? orderBy = null,
        int? limit = null)
    {
        var sql = new StringBuilder();

        var columnList = columns.Count > 0
            ? string.Join(", ", columns.Select(c => $"[{SanitizeIdentifier(c)}]"))
            : "*";
        sql.AppendLine($"SELECT {columnList}");
        sql.AppendLine($"FROM [{SanitizeIdentifier(tableName)}]");

        if (conditions.Count > 0)
        {
            sql.AppendLine("WHERE ");
            for (int i = 0; i < conditions.Count; i++)
            {
                var condition = conditions[i];
                if (i > 0)
                {
                    sql.AppendLine($"  {condition.LogicOperator} ");
                }
                sql.AppendLine($"  {BuildConditionSql(condition)}");
            }
        }

        if (!string.IsNullOrWhiteSpace(orderBy))
        {
            sql.AppendLine($"ORDER BY [{SanitizeIdentifier(orderBy)}]");
        }

        if (limit.HasValue && limit.Value > 0)
        {
            sql.AppendLine($"LIMIT {limit.Value}");
        }

        return sql.ToString();
    }

    /// <summary>
    /// 生成统计脚本
    /// </summary>
    public string GenerateStatisticsScript(
        string tableName,
        List<string> groupByFields,
        List<StatisticsMetric> metrics)
    {
        var sql = new StringBuilder();
        var selectParts = new List<string>();

        foreach (var field in groupByFields)
        {
            selectParts.Add($"[{SanitizeIdentifier(field)}]");
        }

        foreach (var metric in metrics)
        {
            var field = SanitizeIdentifier(metric.FieldName);
            var func = metric.AggregateFunction.ToUpper();
            var alias = string.IsNullOrWhiteSpace(metric.Alias)
                ? $"{func}_{field}"
                : SanitizeIdentifier(metric.Alias);

            selectParts.Add($"{func}([{field}]) AS [{alias}]");
        }

        sql.AppendLine($"SELECT {string.Join(", ", selectParts)}");
        sql.AppendLine($"FROM [{SanitizeIdentifier(tableName)}]");

        if (groupByFields.Count > 0)
        {
            var groupByClause = string.Join(", ", groupByFields.Select(f => $"[{SanitizeIdentifier(f)}]"));
            sql.AppendLine($"GROUP BY {groupByClause}");
        }

        if (metrics.Count > 0)
        {
            var firstMetric = metrics[0];
            var alias = string.IsNullOrWhiteSpace(firstMetric.Alias)
                ? $"{firstMetric.AggregateFunction.ToUpper()}_{SanitizeIdentifier(firstMetric.FieldName)}"
                : SanitizeIdentifier(firstMetric.Alias);
            sql.AppendLine($"ORDER BY [{alias}] DESC");
        }

        return sql.ToString();
    }

    /// <summary>
    /// 生成关联查询脚本
    /// </summary>
    public string GenerateJoinQueryScript(
        string mainTable,
        string joinTable,
        string mainField,
        string joinField,
        List<string> columns,
        string joinType = "INNER")
    {
        var columnList = columns.Count > 0
            ? string.Join(", ", columns.Select(c => $"[{SanitizeIdentifier(c)}]"))
            : "*";

        var sql = new StringBuilder();
        sql.AppendLine($"SELECT {columnList}");
        sql.AppendLine($"FROM [{SanitizeIdentifier(mainTable)}]");
        sql.AppendLine($"{joinType} JOIN [{SanitizeIdentifier(joinTable)}]");
        sql.AppendLine($"  ON [{SanitizeIdentifier(mainTable)}].[{SanitizeIdentifier(mainField)}] = ");
        sql.AppendLine($"     [{SanitizeIdentifier(joinTable)}].[{SanitizeIdentifier(joinField)}]");

        return sql.ToString();
    }

    /// <summary>
    /// 生成CREATE TABLE脚本
    /// </summary>
    public string GenerateCreateTableScript(string tableName, List<FieldInfo> fields)
    {
        var sql = new StringBuilder();
        sql.AppendLine($"CREATE TABLE [{SanitizeIdentifier(tableName)}] (");

        var columnDefs = new List<string>();
        foreach (var field in fields)
        {
            var colName = SanitizeIdentifier(field.Name);
            var colType = MapToSqliteType(field.DataType);
            columnDefs.Add($"  [{colName}] {colType}");
        }

        sql.AppendLine(string.Join(",\n", columnDefs));
        sql.AppendLine(");");

        return sql.ToString();
    }

    /// <summary>
    /// 生成INSERT脚本模板
    /// </summary>
    public string GenerateInsertTemplateScript(string tableName, List<string> columns)
    {
        var colList = string.Join(", ", columns.Select(c => $"[{SanitizeIdentifier(c)}]"));
        var paramList = string.Join(", ", columns.Select((_, i) => $"@p{i}"));

        return $"INSERT INTO [{SanitizeIdentifier(tableName)}] ({colList}) VALUES ({paramList});";
    }

    private string BuildConditionSql(QueryCondition condition)
    {
        var field = SanitizeIdentifier(condition.FieldName);
        var op = condition.Operator.ToUpper();
        var value = condition.Value.Replace("'", "''");

        return op switch
        {
            "=" or "!=" or "<>" or ">" or "<" or ">=" or "<=" =>
                $"[{field}] {op} '{value}'",
            "LIKE" => $"[{field}] LIKE '%{value}%'",
            "NOT LIKE" => $"[{field}] NOT LIKE '%{value}%'",
            "IS NULL" => $"[{field}] IS NULL",
            "IS NOT NULL" => $"[{field}] IS NOT NULL",
            _ => $"[{field}] = '{value}'"
        };
    }

    private string MapToSqliteType(string dataType)
    {
        return dataType.ToLower() switch
        {
            "number" or "numeric" or "decimal" or "float" or "double" => "REAL",
            "int" or "integer" or "long" or "short" or "byte" => "INTEGER",
            "datetime" or "date" or "time" => "TEXT",
            "boolean" or "bool" => "INTEGER",
            _ => "TEXT"
        };
    }

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

/// <summary>
/// SQL脚本配置
/// </summary>
public class ScriptConfig
{
    public string Description { get; set; } = string.Empty;
    public List<QueryCondition> Conditions { get; set; } = new();
    public List<StatisticsMetric> Metrics { get; set; } = new();
}
