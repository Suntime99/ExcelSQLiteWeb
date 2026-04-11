namespace ExcelSQLiteWeb.Models;

/// <summary>
/// 导入进度（用于 WebView 进度条）
/// </summary>
public sealed class ImportProgress
{
    public string Stage { get; set; } = "";
    public int Percentage { get; set; }
    public int CurrentRow { get; set; }
    public int TotalRows { get; set; }
}

/// <summary>
/// 智能转换统计（与前端约定使用 lowerCamel 字段名）
/// </summary>
public sealed class ConversionStat
{
    public string columnName { get; set; } = "";
    public string targetType { get; set; } = "Text";
    public long nonEmptyCount { get; set; }
    public long okCount { get; set; }
    public long failCount { get; set; }
    public List<string> failSamples { get; set; } = new();
}

/// <summary>
/// 导入结果
/// </summary>
public sealed class ImportResult
{
    public string FilePath { get; set; } = "";
    public string WorksheetName { get; set; } = "";
    public string TableName { get; set; } = "";
    public string ImportMode { get; set; } = "text";

    public bool Success { get; set; }
    public string Message { get; set; } = "";

    public int RowCount { get; set; }
    public int ColumnCount { get; set; }
    public double ImportTime { get; set; }

    public List<ConversionStat>? ConversionStats { get; set; }
}

