namespace ExcelSQLiteWeb.Models;

/// <summary>
/// 文件分析结果
/// </summary>
public class FileAnalysisResult
{
    /// <summary>
    /// 文件名
    /// </summary>
    public string FileName { get; set; } = string.Empty;

    /// <summary>
    /// 文件路径
    /// </summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>
    /// 文件大小（字节）
    /// </summary>
    public long FileSize { get; set; }

    /// <summary>
    /// 工作表数量
    /// </summary>
    public int WorksheetCount { get; set; }

    /// <summary>
    /// 总行数
    /// </summary>
    public int TotalRowCount { get; set; }

    /// <summary>
    /// 总列数
    /// </summary>
    public int TotalColumnCount { get; set; }

    /// <summary>
    /// 分析耗时（秒）
    /// </summary>
    public double AnalyzeTime { get; set; }

    /// <summary>
    /// 数据质量评分
    /// </summary>
    public double DataQualityScore { get; set; }

    /// <summary>
    /// 工作表信息列表
    /// </summary>
    public List<WorksheetInfo> Worksheets { get; set; } = new();
}

/// <summary>
/// 工作表信息
/// </summary>
public class WorksheetInfo
{
    /// <summary>
    /// 工作表名称
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// 行数（包含标题行）
    /// </summary>
    public int RowCount { get; set; }

    /// <summary>
    /// 列数
    /// </summary>
    public int ColumnCount { get; set; }

    /// <summary>
    /// 数据行数（不包含标题行）
    /// </summary>
    public int DataRowCount => RowCount > 0 ? RowCount - 1 : 0;

    /// <summary>
    /// 工作表大小（估算）
    /// </summary>
    public string Size { get; set; } = string.Empty;

    /// <summary>
    /// 数据完整度百分比
    /// </summary>
    public double Completeness { get; set; }

    /// <summary>
    /// 字段信息列表
    /// </summary>
    public List<FieldInfo> Fields { get; set; } = new();

    /// <summary>
    /// 是否已导入SQLite
    /// </summary>
    public bool IsImported { get; set; }

    /// <summary>
    /// SQLite表名
    /// </summary>
    public string? TableName { get; set; }
}

/// <summary>
/// 字段信息
/// </summary>
public class FieldInfo
{
    /// <summary>
    /// 字段名称
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// 列索引（从1开始）
    /// </summary>
    public int ColumnIndex { get; set; }

    /// <summary>
    /// 数据类型
    /// </summary>
    public string DataType { get; set; } = "Text";

    /// <summary>
    /// 是否可为空
    /// </summary>
    public bool IsNullable { get; set; } = true;

    /// <summary>
    /// 空值数量
    /// </summary>
    public int NullCount { get; set; }

    /// <summary>
    /// 唯一值数量
    /// </summary>
    public int UniqueValueCount { get; set; }

    /// <summary>
    /// 示例值
    /// </summary>
    public List<string> SampleValues { get; set; } = new();

    /// <summary>
    /// 最大长度
    /// </summary>
    public int MaxLength { get; set; }

    /// <summary>
    /// 是否为关键字段
    /// </summary>
    public bool IsKeyField { get; set; }
}

/// <summary>
/// 查询条件
/// </summary>
public class QueryCondition
{
    /// <summary>
    /// 字段名
    /// </summary>
    public string FieldName { get; set; } = string.Empty;

    /// <summary>
    /// 操作符
    /// </summary>
    public string Operator { get; set; } = "=";

    /// <summary>
    /// 值
    /// </summary>
    public string Value { get; set; } = string.Empty;

    /// <summary>
    /// 逻辑连接符（AND/OR）
    /// </summary>
    public string LogicOperator { get; set; } = "AND";
}

/// <summary>
/// 统计配置
/// </summary>
public class StatisticsConfig
{
    /// <summary>
    /// 分组字段
    /// </summary>
    public List<string> GroupByFields { get; set; } = new();

    /// <summary>
    /// 统计指标
    /// </summary>
    public List<StatisticsMetric> Metrics { get; set; } = new();
}

/// <summary>
/// 统计指标
/// </summary>
public class StatisticsMetric
{
    /// <summary>
    /// 字段名
    /// </summary>
    public string FieldName { get; set; } = string.Empty;

    /// <summary>
    /// 聚合函数（COUNT, SUM, AVG, MAX, MIN）
    /// </summary>
    public string AggregateFunction { get; set; } = "COUNT";

    /// <summary>
    /// 别名
    /// </summary>
    public string Alias { get; set; } = string.Empty;
}

/// <summary>
/// 分拆配置
/// </summary>
public class SplitConfig
{
    /// <summary>
    /// 分拆字段
    /// </summary>
    public string SplitField { get; set; } = string.Empty;

    /// <summary>
    /// 输出目录
    /// </summary>
    public string OutputDirectory { get; set; } = string.Empty;

    /// <summary>
    /// 文件名前缀
    /// </summary>
    public string FileNamePrefix { get; set; } = string.Empty;

    /// <summary>
    /// 是否保留原始格式
    /// </summary>
    public bool KeepOriginalFormat { get; set; } = true;
}

/// <summary>
/// 全局搜索结果
/// </summary>
public class GlobalSearchResult
{
    /// <summary>
    /// 工作表名称
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// 字段名
    /// </summary>
    public string FieldName { get; set; } = string.Empty;

    /// <summary>
    /// 匹配内容
    /// </summary>
    public string MatchContent { get; set; } = string.Empty;

    /// <summary>
    /// 行号
    /// </summary>
    public int RowNumber { get; set; }

    /// <summary>
    /// 列号
    /// </summary>
    public int ColumnNumber { get; set; }
}

/// <summary>
/// 查询结果
/// </summary>
public class QueryResult
{
    /// <summary>
    /// 列名
    /// </summary>
    public List<string> Columns { get; set; } = new();

    /// <summary>
    /// 数据行
    /// </summary>
    public List<Dictionary<string, object>> Rows { get; set; } = new();

    /// <summary>
    /// 总行数
    /// </summary>
    public int TotalRows { get; set; }

    /// <summary>
    /// 查询耗时（秒）
    /// </summary>
    public double QueryTime { get; set; }

    /// <summary>
    /// SQL语句
    /// </summary>
    public string Sql { get; set; } = string.Empty;
}

/// <summary>
/// 统计结果
/// </summary>
public class StatisticsResult
{
    /// <summary>
    /// 列名
    /// </summary>
    public List<string> Columns { get; set; } = new();

    /// <summary>
    /// 数据行
    /// </summary>
    public List<Dictionary<string, object>> Rows { get; set; } = new();

    /// <summary>
    /// 总行数
    /// </summary>
    public int TotalRows { get; set; }

    /// <summary>
    /// 统计耗时（秒）
    /// </summary>
    public double StatisticsTime { get; set; }

    /// <summary>
    /// SQL语句
    /// </summary>
    public string Sql { get; set; } = string.Empty;

    /// <summary>
    /// 分组字段
    /// </summary>
    public List<string> GroupByFields { get; set; } = new();

    /// <summary>
    /// 统计指标
    /// </summary>
    public List<string> Metrics { get; set; } = new();
}

/// <summary>
/// 分拆结果
/// </summary>
public class SplitResult
{
    /// <summary>
    /// 分拆文件列表
    /// </summary>
    public List<SplitFileInfo> Files { get; set; } = new();

    /// <summary>
    /// 分拆耗时（秒）
    /// </summary>
    public double SplitTime { get; set; }

    /// <summary>
    /// 总行数
    /// </summary>
    public int TotalRows { get; set; }
}

/// <summary>
/// 分拆文件信息
/// </summary>
public class SplitFileInfo
{
    /// <summary>
    /// 文件名
    /// </summary>
    public string FileName { get; set; } = string.Empty;

    /// <summary>
    /// 文件路径
    /// </summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>
    /// 行数
    /// </summary>
    public int RowCount { get; set; }

    /// <summary>
    /// 分拆字段值
    /// </summary>
    public string SplitValue { get; set; } = string.Empty;
}
