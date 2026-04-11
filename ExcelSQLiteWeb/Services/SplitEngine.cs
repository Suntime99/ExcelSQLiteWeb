using ExcelSQLiteWeb.Models;
using OfficeOpenXml;
using System.Diagnostics;
using System.Text;

namespace ExcelSQLiteWeb.Services;

/// <summary>
/// 数据分拆引擎 - 单表/临时表结果集分拆导出
/// </summary>
public class SplitEngine
{
    private readonly SqliteManager _sqliteManager;
    private readonly ExcelAnalyzer _excelAnalyzer;

    public SplitEngine(SqliteManager sqliteManager, ExcelAnalyzer excelAnalyzer)
    {
        _sqliteManager = sqliteManager;
        _excelAnalyzer = excelAnalyzer;
    }

    /// <summary>
    /// 按字段值分拆数据
    /// outputOption:
    /// - files: 每个分拆值一个文件（xlsx/csv）
    /// - sheets: 一个工作簿多个工作表（xlsx）
    /// - csv: 每个分拆值一个csv
    /// </summary>
    public SplitResult SplitByField(
        string sourceTable,
        string splitField,
        string outputDirectory,
        string outputOption = "files",
        string? fileNamePrefix = null,
        bool includeNullGroup = true,
        bool overwriteExisting = false,
        string csvDelimiter = ",",
        Action<int, string>? progress = null)
    {
        var sw = Stopwatch.StartNew();
        var result = new SplitResult();

        try
        {
            EnsureOutputDirectory(outputDirectory);
            var mode = (outputOption ?? "files").Trim().ToLowerInvariant();
            if (mode != "files" && mode != "sheets" && mode != "csv") mode = "files";

            fileNamePrefix ??= sourceTable;

            // 1) 取唯一值（NULL 不包含在 DISTINCT 结果中，因此单独处理）
            var uniqueValues = _sqliteManager.GetUniqueValues(sourceTable, splitField, limit: 10000);

            // 2) 表结构
            var schema = _sqliteManager.GetTableSchema(sourceTable);
            var columnNames = schema.Select(s => s.ColumnName).ToList();

            var totalGroups = uniqueValues.Count + (includeNullGroup ? 1 : 0);
            var doneGroups = 0;
            void Tick(string stage)
            {
                doneGroups++;
                var pct = totalGroups <= 0 ? 0 : (int)Math.Round(doneGroups * 100.0 / totalGroups);
                progress?.Invoke(Math.Max(0, Math.Min(100, pct)), stage);
            }

            if (mode == "sheets")
            {
                // 一个 workbook，多 sheet
                var outFile = Path.Combine(outputDirectory, $"{SanitizeFileName(fileNamePrefix)}_{SanitizeFileName(splitField)}_split.xlsx");
                EnsureCanWrite(outFile, overwriteExisting);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using var package = new ExcelPackage();
                var usedSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var val in uniqueValues)
                {
                    var data = QueryByField(sourceTable, splitField, val);
                    if (data.Count == 0) continue;

                    var sheetName = MakeSafeSheetName(val, usedSheetNames);
                    var ws = package.Workbook.Worksheets.Add(sheetName);
                    ExportToWorksheet(data, columnNames, ws);

                    result.Files.Add(new SplitFileInfo
                    {
                        FileName = Path.GetFileName(outFile),
                        FilePath = outFile,
                        RowCount = data.Count,
                        SplitValue = val
                    });
                    result.TotalRows += data.Count;
                    Tick($"导出：{val}");
                }

                if (includeNullGroup)
                {
                    var dataNull = QueryNullGroup(sourceTable, splitField);
                    if (dataNull.Count > 0)
                    {
                        var sheetName = MakeSafeSheetName("NULL", usedSheetNames);
                        var ws = package.Workbook.Worksheets.Add(sheetName);
                        ExportToWorksheet(dataNull, columnNames, ws);
                        result.Files.Add(new SplitFileInfo
                        {
                            FileName = Path.GetFileName(outFile),
                            FilePath = outFile,
                            RowCount = dataNull.Count,
                            SplitValue = "(NULL)"
                        });
                        result.TotalRows += dataNull.Count;
                        Tick("导出：NULL");
                    }
                }

                if (package.Workbook.Worksheets.Count == 0)
                {
                    throw new InvalidOperationException("没有可分拆的数据（请检查分拆字段是否存在或是否有数据）");
                }

                package.SaveAs(new FileInfo(outFile));
            }
            else
            {
                // files / csv：每个分拆值一个文件
                foreach (var val in uniqueValues)
                {
                    var data = QueryByField(sourceTable, splitField, val);
                    if (data.Count == 0) continue;

                    var safeValue = SanitizeFileName(val);
                    var ext = mode == "csv" ? "csv" : "xlsx";
                    var fileName = $"{SanitizeFileName(fileNamePrefix)}_{safeValue}.{ext}";
                    var filePath = Path.Combine(outputDirectory, fileName);
                    EnsureCanWrite(filePath, overwriteExisting);

                    if (mode == "csv")
                        ExportToCsv(data, columnNames, filePath, csvDelimiter);
                    else
                        ExportToExcel(data, columnNames, filePath);

                    result.Files.Add(new SplitFileInfo
                    {
                        FileName = fileName,
                        FilePath = filePath,
                        RowCount = data.Count,
                        SplitValue = val
                    });
                    result.TotalRows += data.Count;
                    Tick($"导出：{val}");
                }

                if (includeNullGroup)
                {
                    var dataNull = QueryNullGroup(sourceTable, splitField);
                    if (dataNull.Count > 0)
                    {
                        var ext = mode == "csv" ? "csv" : "xlsx";
                        var fileName = $"{SanitizeFileName(fileNamePrefix)}_NULL.{ext}";
                        var filePath = Path.Combine(outputDirectory, fileName);
                        EnsureCanWrite(filePath, overwriteExisting);

                        if (mode == "csv")
                            ExportToCsv(dataNull, columnNames, filePath, csvDelimiter);
                        else
                            ExportToExcel(dataNull, columnNames, filePath);

                        result.Files.Add(new SplitFileInfo
                        {
                            FileName = fileName,
                            FilePath = filePath,
                            RowCount = dataNull.Count,
                            SplitValue = "(NULL)"
                        });
                        result.TotalRows += dataNull.Count;
                        Tick("导出：NULL");
                    }
                }

                if (result.Files.Count == 0)
                {
                    throw new InvalidOperationException("没有可分拆的数据（请检查分拆字段是否存在或是否有数据）");
                }
            }

            result.SplitTime = Math.Round(sw.Elapsed.TotalSeconds, 2);
            progress?.Invoke(100, "完成");
            return result;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"数据分拆失败: {ex.Message}", ex);
        }
        finally
        {
            sw.Stop();
        }
    }

    /// <summary>
    /// 按行数分拆数据
    /// outputOption:
    /// - files: 每份一个xlsx/csv文件
    /// - sheets: 一个xlsx多个sheet
    /// - csv: 每份一个csv
    /// </summary>
    public SplitResult SplitByRowCount(
        string sourceTable,
        int rowsPerFile,
        string outputDirectory,
        string outputOption = "files",
        string? fileNamePrefix = null,
        bool overwriteExisting = false,
        string csvDelimiter = ",",
        Action<int, string>? progress = null)
    {
        var sw = Stopwatch.StartNew();
        var result = new SplitResult();

        try
        {
            EnsureOutputDirectory(outputDirectory);
            if (rowsPerFile <= 0) rowsPerFile = 50000;

            var mode = (outputOption ?? "files").Trim().ToLowerInvariant();
            if (mode != "files" && mode != "sheets" && mode != "csv") mode = "files";

            fileNamePrefix ??= sourceTable;

            int totalRows = _sqliteManager.GetRowCount(sourceTable);
            var schema = _sqliteManager.GetTableSchema(sourceTable);
            var columnNames = schema.Select(s => s.ColumnName).ToList();
            var totalParts = (int)Math.Ceiling(totalRows / (double)rowsPerFile);
            if (totalParts <= 0) totalParts = 0;
            void Tick(int part, string stage)
            {
                var pct = totalParts <= 0 ? 0 : (int)Math.Round(part * 100.0 / totalParts);
                progress?.Invoke(Math.Max(0, Math.Min(100, pct)), stage);
            }

            if (mode == "sheets")
            {
                var outFile = Path.Combine(outputDirectory, $"{SanitizeFileName(fileNamePrefix)}_split_by_rows.xlsx");
                EnsureCanWrite(outFile, overwriteExisting);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using var package = new ExcelPackage();

                int part = 1;
                for (int offset = 0; offset < totalRows; offset += rowsPerFile)
                {
                    var data = QueryByOffset(sourceTable, rowsPerFile, offset);
                    if (data.Count == 0) break;

                    var ws = package.Workbook.Worksheets.Add($"Part{part:D3}");
                    ExportToWorksheet(data, columnNames, ws);

                    result.Files.Add(new SplitFileInfo
                    {
                        FileName = Path.GetFileName(outFile),
                        FilePath = outFile,
                        RowCount = data.Count,
                        SplitValue = $"Part{part:D3}"
                    });
                    result.TotalRows += data.Count;
                    Tick(part, $"导出：Part{part:D3}");
                    part++;
                }

                if (package.Workbook.Worksheets.Count == 0)
                    throw new InvalidOperationException("没有可分拆的数据（空表）");

                package.SaveAs(new FileInfo(outFile));
            }
            else
            {
                int part = 1;
                for (int offset = 0; offset < totalRows; offset += rowsPerFile)
                {
                    var data = QueryByOffset(sourceTable, rowsPerFile, offset);
                    if (data.Count == 0) break;

                    var ext = mode == "csv" ? "csv" : "xlsx";
                    var fileName = $"{SanitizeFileName(fileNamePrefix)}_Part{part:D3}.{ext}";
                    var filePath = Path.Combine(outputDirectory, fileName);
                    EnsureCanWrite(filePath, overwriteExisting);

                    if (mode == "csv")
                        ExportToCsv(data, columnNames, filePath, csvDelimiter);
                    else
                        ExportToExcel(data, columnNames, filePath);

                    result.Files.Add(new SplitFileInfo
                    {
                        FileName = fileName,
                        FilePath = filePath,
                        RowCount = data.Count,
                        SplitValue = $"Part{part:D3}"
                    });
                    result.TotalRows += data.Count;
                    Tick(part, $"导出：Part{part:D3}");
                    part++;
                }

                if (result.Files.Count == 0)
                    throw new InvalidOperationException("没有可分拆的数据（空表）");
            }

            result.SplitTime = Math.Round(sw.Elapsed.TotalSeconds, 2);
            progress?.Invoke(100, "完成");
            return result;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"数据分拆失败: {ex.Message}", ex);
        }
        finally
        {
            sw.Stop();
        }
    }

    // ====================== Query Helpers ======================

    private List<Dictionary<string, object>> QueryByField(string table, string field, string value)
    {
        // 使用参数化，避免引号/特殊字符破坏 SQL
        var sql = $"SELECT * FROM {SqliteManager.QuoteIdent(table)} WHERE {SqliteManager.QuoteIdent(field)} = @v";
        return _sqliteManager.Query(sql, new { v = value });
    }

    private List<Dictionary<string, object>> QueryNullGroup(string table, string field)
    {
        var sql = $"SELECT * FROM {SqliteManager.QuoteIdent(table)} WHERE {SqliteManager.QuoteIdent(field)} IS NULL";
        return _sqliteManager.Query(sql);
    }

    private List<Dictionary<string, object>> QueryByOffset(string table, int limit, int offset)
    {
        var sql = $"SELECT * FROM {SqliteManager.QuoteIdent(table)} LIMIT {limit} OFFSET {offset}";
        return _sqliteManager.Query(sql);
    }

    // ====================== Export Helpers ======================

    private void ExportToExcel(List<Dictionary<string, object>> data, List<string> columnNames, string filePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using var package = new ExcelPackage();
        var ws = package.Workbook.Worksheets.Add("Data");
        ExportToWorksheet(data, columnNames, ws);
        package.SaveAs(new FileInfo(filePath));
    }

    private void ExportToWorksheet(List<Dictionary<string, object>> data, List<string> columnNames, ExcelWorksheet worksheet)
    {
        // 表头
        for (int i = 0; i < columnNames.Count; i++)
        {
            worksheet.Cells[1, i + 1].Value = columnNames[i];
            worksheet.Cells[1, i + 1].Style.Font.Bold = true;
        }

        // 数据
        for (int r = 0; r < data.Count; r++)
        {
            for (int c = 0; c < columnNames.Count; c++)
            {
                var col = columnNames[c];
                object? v = data[r].TryGetValue(col, out var vv) ? vv : null;
                worksheet.Cells[r + 2, c + 1].Value = v == DBNull.Value ? null : v;
            }
        }

        try { worksheet.Cells.AutoFitColumns(); } catch { }
    }

    private void ExportToCsv(List<Dictionary<string, object>> data, List<string> columnNames, string filePath, string delimiter)
    {
        delimiter = string.IsNullOrEmpty(delimiter) ? "," : delimiter[..1];
        string Escape(string s)
        {
            if (s == null) return "";
            bool need = s.Contains('"') || s.Contains(delimiter) || s.Contains('\n') || s.Contains('\r');
            var x = s.Replace("\"", "\"\"");
            return need ? $"\"{x}\"" : x;
        }
        using var sw = new StreamWriter(filePath, append: false, encoding: new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        sw.WriteLine(string.Join(delimiter, columnNames.Select(Escape)));
        foreach (var row in data)
        {
            var fields = columnNames.Select(c =>
            {
                if (!row.TryGetValue(c, out var v) || v == null || v == DBNull.Value) return "";
                return Escape(Convert.ToString(v) ?? "");
            });
            sw.WriteLine(string.Join(delimiter, fields));
        }
    }

    // ====================== Naming Helpers ======================

    private static void EnsureOutputDirectory(string dir)
    {
        if (string.IsNullOrWhiteSpace(dir)) throw new InvalidOperationException("输出目录为空");
        if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
    }

    private static void EnsureCanWrite(string filePath, bool overwrite)
    {
        var fi = new FileInfo(filePath);
        if (fi.Directory != null && !fi.Directory.Exists) fi.Directory.Create();
        if (fi.Exists && !overwrite)
            throw new InvalidOperationException($"输出目录中已存在文件：{fi.Name}。请使用空目录或勾选“允许覆盖已存在文件”。");
    }

    private static string MakeSafeSheetName(string raw, HashSet<string> used)
    {
        var s = string.IsNullOrWhiteSpace(raw) ? "EMPTY" : raw.Trim();
        // Excel sheet name illegal chars: : \ / ? * [ ]
        foreach (var ch in new[] { ':', '\\', '/', '?', '*', '[', ']' })
            s = s.Replace(ch, '_');
        if (s.Length > 31) s = s[..31];
        if (string.IsNullOrWhiteSpace(s)) s = "Sheet";

        var baseName = s;
        int i = 1;
        while (used.Contains(s))
        {
            var suffix = $"_{i}";
            var max = 31 - suffix.Length;
            var head = baseName.Length > max ? baseName[..max] : baseName;
            s = head + suffix;
            i++;
        }
        used.Add(s);
        return s;
    }

    private static string SanitizeFileName(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName)) return "_";
        var invalid = Path.GetInvalidFileNameChars();
        var sanitized = new string(fileName.Where(c => !invalid.Contains(c)).ToArray());
        sanitized = sanitized.Replace(' ', '_');
        if (sanitized.Length > 80) sanitized = sanitized[..80];
        return string.IsNullOrWhiteSpace(sanitized) ? "_" : sanitized;
    }
}
