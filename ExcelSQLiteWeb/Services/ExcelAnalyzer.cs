using ExcelSQLiteWeb.Models;
using OfficeOpenXml;
using System.Diagnostics;

namespace ExcelSQLiteWeb.Services;

/// <summary>
/// Excel文件分析器
/// </summary>
public class ExcelAnalyzer
{
    public ExcelAnalyzer()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    /// <summary>
    /// 分析Excel文件
    /// </summary>
    public FileAnalysisResult Analyze(string filePath)
    {
        var stopwatch = Stopwatch.StartNew();
        var fileInfo = new FileInfo(filePath);

        using var package = new ExcelPackage(fileInfo);
        var workbook = package.Workbook;

        var result = new FileAnalysisResult
        {
            FileName = fileInfo.Name,
            FilePath = filePath,
            FileSize = fileInfo.Length,
            WorksheetCount = workbook.Worksheets.Count,
            Worksheets = new List<WorksheetInfo>()
        };

        int totalRows = 0;
        int totalColumns = 0;
        double totalCompleteness = 0;

        foreach (var worksheet in workbook.Worksheets)
        {
            var worksheetInfo = AnalyzeWorksheet(worksheet);
            result.Worksheets.Add(worksheetInfo);

            totalRows += worksheetInfo.RowCount;
            totalColumns += worksheetInfo.ColumnCount;
            totalCompleteness += worksheetInfo.Completeness;
        }

        result.TotalRowCount = totalRows;
        result.TotalColumnCount = totalColumns;
        result.DataQualityScore = result.Worksheets.Count > 0 ? totalCompleteness / result.Worksheets.Count : 0;

        stopwatch.Stop();
        result.AnalyzeTime = Math.Round(stopwatch.Elapsed.TotalSeconds, 2);

        return result;
    }

    /// <summary>
    /// 分析单个工作表
    /// </summary>
    private WorksheetInfo AnalyzeWorksheet(ExcelWorksheet worksheet)
    {
        var info = new WorksheetInfo
        {
            Name = worksheet.Name,
            RowCount = worksheet.Dimension?.Rows ?? 0,
            ColumnCount = worksheet.Dimension?.Columns ?? 0,
            Fields = new List<FieldInfo>()
        };

        if (info.RowCount == 0 || info.ColumnCount == 0)
        {
            info.Size = "0 KB";
            info.Completeness = 0;
            return info;
        }

        // 估算大小
        long estimatedSize = (long)(info.RowCount * info.ColumnCount * 8L);
        info.Size = FormatFileSize(estimatedSize);

        // 分析每个字段
        int totalNullCount = 0;
        int totalCells = 0;

        for (int col = 1; col <= info.ColumnCount; col++)
        {
            var fieldInfo = AnalyzeField(worksheet, col);
            info.Fields.Add(fieldInfo);
            totalNullCount += fieldInfo.NullCount;
            totalCells += info.DataRowCount;
        }

        // 计算完整度
        info.Completeness = totalCells > 0
            ? Math.Round((1 - (double)totalNullCount / totalCells) * 100, 1)
            : 100;

        return info;
    }

    /// <summary>
    /// 分析单个字段
    /// </summary>
    private FieldInfo AnalyzeField(ExcelWorksheet worksheet, int columnIndex)
    {
        var fieldInfo = new FieldInfo
        {
            ColumnIndex = columnIndex,
            SampleValues = new List<string>()
        };

        int rowCount = worksheet.Dimension?.Rows ?? 0;
        if (rowCount == 0) return fieldInfo;

        // 获取字段名（第一行）
        fieldInfo.Name = worksheet.Cells[1, columnIndex].Text?.Trim() ?? $"Column{columnIndex}";

        // 收集数据类型和统计信息
        var uniqueValues = new HashSet<string>();
        int nullCount = 0;
        int maxLength = 0;
        bool hasNumeric = false;
        bool hasDate = false;
        bool hasBool = false;

        // 采样最多100个值
        int sampleSize = Math.Min(rowCount - 1, 100);
        int sampleInterval = Math.Max(1, (rowCount - 1) / sampleSize);

        for (int row = 2; row <= rowCount; row++)
        {
            var cell = worksheet.Cells[row, columnIndex];
            string value = cell.Text?.Trim() ?? string.Empty;

            if (string.IsNullOrEmpty(value))
            {
                nullCount++;
                continue;
            }

            uniqueValues.Add(value);
            maxLength = Math.Max(maxLength, value.Length);

            // 采样值
            if (fieldInfo.SampleValues.Count < 5 && (row - 2) % sampleInterval == 0)
            {
                fieldInfo.SampleValues.Add(value.Length > 50 ? value.Substring(0, 50) + "..." : value);
            }

            // 检测数据类型
            if (!hasNumeric && double.TryParse(value, out _))
                hasNumeric = true;
            if (!hasDate && DateTime.TryParse(value, out _))
                hasDate = true;
            if (!hasBool && bool.TryParse(value, out _))
                hasBool = true;
        }

        fieldInfo.NullCount = nullCount;
        fieldInfo.UniqueValueCount = uniqueValues.Count;
        fieldInfo.MaxLength = maxLength;

        // 确定数据类型
        fieldInfo.DataType = DetermineDataType(hasNumeric, hasDate, hasBool, uniqueValues);

        // 判断是否可能是关键字段（唯一值多且不为空）
        fieldInfo.IsKeyField = nullCount == 0 && uniqueValues.Count > (rowCount - 1) * 0.8;

        return fieldInfo;
    }

    /// <summary>
    /// 确定数据类型
    /// </summary>
    private string DetermineDataType(bool hasNumeric, bool hasDate, bool hasBool, HashSet<string> uniqueValues)
    {
        if (hasBool && uniqueValues.Count <= 2)
            return "Boolean";
        if (hasDate && uniqueValues.Count > 0)
            return "DateTime";
        if (hasNumeric && uniqueValues.Count > 0)
            return "Number";
        return "Text";
    }

    /// <summary>
    /// 格式化文件大小
    /// </summary>
    private string FormatFileSize(long bytes)
    {
        if (bytes < 1024) return $"{bytes} B";
        if (bytes < 1024 * 1024) return $"{bytes / 1024.0:F1} KB";
        if (bytes < 1024 * 1024 * 1024) return $"{bytes / (1024.0 * 1024.0):F1} MB";
        return $"{bytes / (1024.0 * 1024.0 * 1024.0):F1} GB";
    }

    /// <summary>
    /// 获取工作表数据（用于导入）
    /// </summary>
    public List<Dictionary<string, object>> GetWorksheetData(string filePath, string worksheetName, int maxRows = 0)
    {
        var result = new List<Dictionary<string, object>>();

        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets[worksheetName];
        if (worksheet == null) return result;

        int rowCount = worksheet.Dimension?.Rows ?? 0;
        int colCount = worksheet.Dimension?.Columns ?? 0;

        if (rowCount < 2) return result;

        // 获取列名
        var columnNames = new List<string>();
        for (int col = 1; col <= colCount; col++)
        {
            string colName = worksheet.Cells[1, col].Text?.Trim() ?? $"Column{col}";
            columnNames.Add(SanitizeColumnName(colName));
        }

        // 读取数据
        int endRow = maxRows > 0 ? Math.Min(rowCount, maxRows + 1) : rowCount;

        for (int row = 2; row <= endRow; row++)
        {
            var rowData = new Dictionary<string, object>();
            for (int col = 1; col <= colCount; col++)
            {
                var cell = worksheet.Cells[row, col];
                object value = GetCellValue(cell);
                rowData[columnNames[col - 1]] = value;
            }
            result.Add(rowData);
        }

        return result;
    }

    /// <summary>
    /// 获取单元格值
    /// </summary>
    private object GetCellValue(ExcelRange cell)
    {
        if (cell.Value == null) return DBNull.Value;

        return cell.Value switch
        {
            double d => d,
            int i => i,
            long l => l,
            decimal dec => dec,
            DateTime dt => dt,
            bool b => b,
            _ => cell.Text?.Trim() ?? string.Empty
        };
    }

    /// <summary>
    /// 清理列名（用于SQL）
    /// </summary>
    private string SanitizeColumnName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return "Column";

        // 移除非法字符
        var sanitized = new string(name.Where(c => char.IsLetterOrDigit(c) || c == '_').ToArray());

        // 确保以字母开头
        if (sanitized.Length > 0 && char.IsDigit(sanitized[0]))
            sanitized = "Col" + sanitized;

        return string.IsNullOrEmpty(sanitized) ? "Column" : sanitized;
    }

    /// <summary>
    /// 获取工作表列表
    /// </summary>
    public List<string> GetWorksheetNames(string filePath)
    {
        using var package = new ExcelPackage(new FileInfo(filePath));
        return package.Workbook.Worksheets.Select(w => w.Name).ToList();
    }

    /// <summary>
    /// 获取工作表字段列表
    /// </summary>
    public List<string> GetWorksheetFields(string filePath, string worksheetName)
    {
        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets[worksheetName];
        if (worksheet == null) return new List<string>();

        int colCount = worksheet.Dimension?.Columns ?? 0;
        var fields = new List<string>();

        for (int col = 1; col <= colCount; col++)
        {
            string fieldName = worksheet.Cells[1, col].Text?.Trim() ?? $"Column{col}";
            fields.Add(fieldName);
        }

        return fields;
    }
}
