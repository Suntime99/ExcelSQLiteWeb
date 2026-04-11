using Microsoft.Data.Sqlite;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using ExcelDataReader;
using ExcelSQLiteWeb.Models;

namespace ExcelSQLiteWeb.Services;

/// <summary>
/// 数据导入器：
/// - 使用 ExcelDataReader 流式读取（比 EPPlus 逐单元格快很多）
/// - 支持两种导入模式：全列文本 / 智能转换+失败统计
/// - 支持“追加导入”（append=true）：目标表存在时不重建表，仅追加写入（支持源字段新增：自动加列）
/// </summary>
public class DataImporter
{
    private readonly ExcelAnalyzer _excelAnalyzer;
    private readonly SqliteManager _sqliteManager;

    public DataImporter(ExcelAnalyzer excelAnalyzer, SqliteManager sqliteManager)
    {
        _excelAnalyzer = excelAnalyzer;
        _sqliteManager = sqliteManager;
    }

    public Task<ImportResult> ImportWorksheetAsync(
        string filePath,
        string worksheetName,
        string? tableName = null,
        string importMode = "text",
        IProgress<ImportProgress>? progress = null,
        CancellationToken cancellationToken = default,
        bool append = false)
    {
        return Task.Run(() =>
            ImportWorksheetCore(filePath, worksheetName, tableName, importMode, progress, cancellationToken, append),
            cancellationToken);
    }

    private sealed class ConversionStats
    {
        private readonly List<string> _colNames;
        private readonly List<(string ColumnName, string DataType)> _columns;
        private readonly long[] _nonEmpty;
        private readonly long[] _ok;
        private readonly long[] _fail;
        private readonly List<string>[] _samples;

        public ConversionStats(List<string> colNames, List<(string ColumnName, string DataType)> columns)
        {
            _colNames = colNames;
            _columns = columns;
            int n = Math.Max(0, colNames.Count);
            _nonEmpty = new long[n];
            _ok = new long[n];
            _fail = new long[n];
            _samples = new List<string>[n];
            for (int i = 0; i < n; i++) _samples[i] = new List<string>();
        }

        public void NonEmpty(int i) { if (i >= 0 && i < _nonEmpty.Length) _nonEmpty[i]++; }
        public void Ok(int i) { if (i >= 0 && i < _ok.Length) _ok[i]++; }
        public void Fail(int i, object? raw)
        {
            if (i < 0 || i >= _fail.Length) return;
            _fail[i]++;
            var s = Convert.ToString(raw, CultureInfo.InvariantCulture) ?? "";
            if (string.IsNullOrWhiteSpace(s)) return;
            var list = _samples[i];
            if (list.Count < 3) list.Add(s.Length > 50 ? s.Substring(0, 50) : s);
        }

        public List<ConversionStat> ToResult()
        {
            var list = new List<ConversionStat>();
            for (int i = 0; i < _colNames.Count; i++)
            {
                list.Add(new ConversionStat
                {
                    columnName = _colNames[i],
                    targetType = _columns[i].DataType,
                    nonEmptyCount = _nonEmpty[i],
                    okCount = _ok[i],
                    failCount = _fail[i],
                    failSamples = _samples[i].ToList()
                });
            }
            return list;
        }
    }

    private ImportResult ImportWorksheetCore(
        string filePath,
        string worksheetName,
        string? tableName,
        string importMode,
        IProgress<ImportProgress>? progress,
        CancellationToken cancellationToken,
        bool append)
    {
        var sw = Stopwatch.StartNew();
        var result = new ImportResult
        {
            FilePath = filePath,
            WorksheetName = worksheetName,
            TableName = tableName ?? SanitizeTableName(worksheetName),
            ImportMode = importMode
        };

        try
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException("Excel文件不存在", filePath);

            progress?.Report(new ImportProgress { Stage = "打开Excel", Percentage = 2 });

            // 支持 .xls 编码
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            // 定位工作表
            if (!MoveToSheet(reader, worksheetName))
                throw new InvalidOperationException($"工作表不存在: {worksheetName}");

            progress?.Report(new ImportProgress { Stage = "读取表头", Percentage = 5 });

            // 读取表头行
            if (!reader.Read())
            {
                result.Success = true;
                result.Message = "工作表为空，无需导入";
                return result;
            }

            int colCount = reader.FieldCount;
            if (colCount <= 0)
            {
                result.Success = true;
                result.Message = "工作表为空，无需导入";
                return result;
            }

            var columnNames = ReadHeaderFromReader(reader);

            // 采样（智能转换需要）
            bool smart = string.Equals(importMode, "smart", StringComparison.OrdinalIgnoreCase);
            List<object?[]> sampleRows = new();
            List<(string ColumnName, string DataType)> columns;
            ConversionStats? stats = null;

            if (!smart)
            {
                columns = columnNames.Select(n => (n, "Text")).ToList();
            }
            else
            {
                progress?.Report(new ImportProgress { Stage = "推断类型(抽样)", Percentage = 10 });
                // 抽样最多 1000 行
                sampleRows = ReadSampleRows(reader, maxRows: 1000, cancellationToken);
                columns = InferColumnTypesBySampling(columnNames, sampleRows);
                stats = new ConversionStats(columnNames, columns);
            }

            result.ColumnCount = columns.Count;

            progress?.Report(new ImportProgress { Stage = "创建/校验表结构", Percentage = 15 });

            // 追加导入（按你确认的策略 C）：
            // - 允许“源字段新增”：自动 ALTER TABLE ADD COLUMN 后再追加
            // - 允许“目标表多字段”：追加时对缺失字段写 NULL（不影响）
            if (append && _sqliteManager.TableExists(result.TableName))
            {
                var existingCols = _sqliteManager.GetTableColumns(result.TableName);
                var exists = new HashSet<string>(existingCols.Select(c => c.Trim()), StringComparer.OrdinalIgnoreCase);
                foreach (var c in columns)
                {
                    var col = (c.ColumnName ?? "").Trim();
                    if (string.IsNullOrWhiteSpace(col)) continue;
                    if (exists.Contains(col)) continue;
                    // 新增列：按推断类型映射 SQLite 类型
                    var sqlType = MapToSqliteType(c.DataType);
                    _sqliteManager.Execute($"ALTER TABLE {SqliteManager.QuoteIdent(result.TableName)} ADD COLUMN {SqliteManager.QuoteIdent(col)} {sqlType};");
                    exists.Add(col);
                }
            }

            // 创建表：覆盖=drop+create；追加=仅当不存在时创建
            _sqliteManager.CreateTable(result.TableName, columns, dropIfExists: !append);

            // SQLite 导入优化（内存库可激进）
            _sqliteManager.Execute("PRAGMA temp_store = MEMORY;");
            _sqliteManager.Execute("PRAGMA synchronous = OFF;");
            _sqliteManager.Execute("PRAGMA journal_mode = OFF;");

            progress?.Report(new ImportProgress { Stage = "导入数据", Percentage = 20, CurrentRow = 0, TotalRows = 0 });

            // 导入：先插入 sampleRows（智能转换），再继续读剩余行
            int imported = 0;
            if (smart && sampleRows.Count > 0)
            {
                imported += BulkInsertRows(result.TableName, columnNames, columns, sampleRows, smart, stats!, progress, cancellationToken, startRowNumber: 2);
            }

            // sampleRows 读完后，reader 当前位置已在样本末尾的下一行：继续导入剩余行
            imported += BulkInsertRemaining(result.TableName, columnNames, columns, reader, smart, stats, progress, cancellationToken, startRowNumber: 2 + sampleRows.Count);

            result.RowCount = imported;
            if (stats != null) result.ConversionStats = stats.ToResult();

            progress?.Report(new ImportProgress { Stage = "完成", Percentage = 100, CurrentRow = imported, TotalRows = imported });
            result.Success = true;
            result.Message = smart
                ? $"成功导入 {imported} 行数据（智能转换）"
                : $"成功导入 {imported} 行数据（全列文本）";
        }
        catch (OperationCanceledException)
        {
            result.Success = false;
            result.Message = "导入已取消";
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.Message = $"导入失败: {ex.Message}";
        }
        finally
        {
            sw.Stop();
            result.ImportTime = Math.Round(sw.Elapsed.TotalSeconds, 2);
        }

        return result;
    }

    private static string MapToSqliteType(string dataType)
    {
        // 与 SqliteManager.MapToSqliteType 保持一致的最小映射
        if (string.IsNullOrWhiteSpace(dataType)) return "TEXT";
        var t = dataType.Trim().ToLowerInvariant();
        if (t.Contains("int")) return "INTEGER";
        if (t.Contains("number") || t.Contains("double") || t.Contains("float") || t.Contains("decimal")) return "REAL";
        if (t.Contains("bool")) return "INTEGER";
        if (t.Contains("date") || t.Contains("time")) return "TEXT"; // 存字符串，避免 SQLite 日期解析差异
        return "TEXT";
    }

    private static bool MoveToSheet(IExcelDataReader reader, string worksheetName)
    {
        // 当前 sheet
        if (string.Equals(reader.Name, worksheetName, StringComparison.OrdinalIgnoreCase))
            return true;

        // 依次 NextResult
        while (reader.NextResult())
        {
            if (string.Equals(reader.Name, worksheetName, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    private static List<string> ReadHeaderFromReader(IExcelDataReader reader)
    {
        int colCount = reader.FieldCount;
        var columnNames = new List<string>(colCount);
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (int i = 0; i < colCount; i++)
        {
            var raw = Convert.ToString(reader.GetValue(i), CultureInfo.InvariantCulture)?.Trim();
            var name = SanitizeColumnName(string.IsNullOrWhiteSpace(raw) ? $"Column{i + 1}" : raw!);
            var baseName = name;
            int suffix = 1;
            while (!used.Add(name))
            {
                suffix++;
                name = $"{baseName}_{suffix}";
            }
            columnNames.Add(name);
        }

        return columnNames;
    }

    private static List<object?[]> ReadSampleRows(IExcelDataReader reader, int maxRows, CancellationToken ct)
    {
        var rows = new List<object?[]>(capacity: Math.Min(maxRows, 1000));
        int colCount = reader.FieldCount;
        while (rows.Count < maxRows && reader.Read())
        {
            ct.ThrowIfCancellationRequested();
            var arr = new object?[colCount];
            for (int i = 0; i < colCount; i++)
                arr[i] = reader.GetValue(i);
            rows.Add(arr);
        }
        return rows;
    }

    private static List<(string ColumnName, string DataType)> InferColumnTypesBySampling(
        List<string> columnNames,
        List<object?[]> sampleRows)
    {
        int colCount = columnNames.Count;
        var intCount = new int[colCount];
        var doubleCount = new int[colCount];
        var dateCount = new int[colCount];
        var boolCount = new int[colCount];
        var textCount = new int[colCount];

        foreach (var row in sampleRows)
        {
            for (int i = 0; i < colCount; i++)
            {
                var v = row[i];
                if (v == null || v == DBNull.Value) continue;
                if (v is bool) { boolCount[i]++; continue; }
                if (v is DateTime) { dateCount[i]++; continue; }
                if (v is sbyte or byte or short or ushort or int or uint or long or ulong) { intCount[i]++; continue; }
                if (v is float or double or decimal) { doubleCount[i]++; continue; }

                var s = Convert.ToString(v, CultureInfo.InvariantCulture)?.Trim();
                if (string.IsNullOrWhiteSpace(s)) continue;
                if (bool.TryParse(s, out _)) { boolCount[i]++; continue; }
                if (DateTime.TryParse(s, out _)) { dateCount[i]++; continue; }
                if (long.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out _)) { intCount[i]++; continue; }
                if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out _)) { doubleCount[i]++; continue; }
                textCount[i]++;
            }
        }

        var columns = new List<(string ColumnName, string DataType)>(colCount);
        for (int i = 0; i < colCount; i++)
        {
            int total = intCount[i] + doubleCount[i] + dateCount[i] + boolCount[i] + textCount[i];
            if (total == 0)
            {
                columns.Add((columnNames[i], "Text"));
                continue;
            }
            // 简单决策：Text 优先兜底
            if (textCount[i] > 0) { columns.Add((columnNames[i], "Text")); continue; }
            if (dateCount[i] >= doubleCount[i] && dateCount[i] >= intCount[i] && dateCount[i] >= boolCount[i]) { columns.Add((columnNames[i], "DateTime")); continue; }
            if (doubleCount[i] >= intCount[i] && doubleCount[i] >= boolCount[i]) { columns.Add((columnNames[i], "Number")); continue; }
            if (intCount[i] >= boolCount[i]) { columns.Add((columnNames[i], "Integer")); continue; }
            columns.Add((columnNames[i], "Boolean"));
        }
        return columns;
    }

    private int BulkInsertRows(
        string tableName,
        List<string> columnNames,
        List<(string ColumnName, string DataType)> columns,
        List<object?[]> rows,
        bool smart,
        ConversionStats stats,
        IProgress<ImportProgress>? progress,
        CancellationToken ct,
        int startRowNumber)
    {
        if (rows.Count == 0) return 0;
        int colCount = columnNames.Count;
        int inserted = 0;

        using var conn = _sqliteManager.Connection;
        using var tx = conn.BeginTransaction();

        string colSql = string.Join(", ", columnNames.Select(n => SqliteManager.QuoteIdent(n)));
        string paramSql = string.Join(", ", columnNames.Select((_, i) => $"@p{i}"));
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"INSERT INTO {SqliteManager.QuoteIdent(tableName)} ({colSql}) VALUES ({paramSql});";
        cmd.Transaction = tx;
        for (int i = 0; i < colCount; i++) cmd.Parameters.Add(new SqliteParameter($"@p{i}", DBNull.Value));

        int batch = 0;
        for (int r = 0; r < rows.Count; r++)
        {
            ct.ThrowIfCancellationRequested();
            var arr = rows[r];
            for (int i = 0; i < colCount; i++)
            {
                var v = (arr.Length > i) ? arr[i] : null;
                if (!smart)
                {
                    cmd.Parameters[i].Value = v == null || v == DBNull.Value ? DBNull.Value : Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";
                }
                else
                {
                    if (v != null && v != DBNull.Value && !string.IsNullOrWhiteSpace(Convert.ToString(v, CultureInfo.InvariantCulture))) stats.NonEmpty(i);
                    var (ok, val) = ConvertSmart(v, columns[i].DataType);
                    if (ok) stats.Ok(i); else stats.Fail(i, v);
                    cmd.Parameters[i].Value = val ?? DBNull.Value;
                }
            }
            cmd.ExecuteNonQuery();
            inserted++;
            batch++;
            if (batch >= 5000)
            {
                batch = 0;
                progress?.Report(new ImportProgress { Stage = "导入数据", Percentage = 20, CurrentRow = startRowNumber + r, TotalRows = 0 });
            }
        }
        tx.Commit();

        return inserted;
    }

    private int BulkInsertRemaining(
        string tableName,
        List<string> columnNames,
        List<(string ColumnName, string DataType)> columns,
        IExcelDataReader reader,
        bool smart,
        ConversionStats? stats,
        IProgress<ImportProgress>? progress,
        CancellationToken ct,
        int startRowNumber)
    {
        int colCount = columnNames.Count;
        int inserted = 0;

        using var conn = _sqliteManager.Connection;
        using var tx = conn.BeginTransaction();

        string colSql = string.Join(", ", columnNames.Select(n => SqliteManager.QuoteIdent(n)));
        string paramSql = string.Join(", ", columnNames.Select((_, i) => $"@p{i}"));
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"INSERT INTO {SqliteManager.QuoteIdent(tableName)} ({colSql}) VALUES ({paramSql});";
        cmd.Transaction = tx;
        for (int i = 0; i < colCount; i++) cmd.Parameters.Add(new SqliteParameter($"@p{i}", DBNull.Value));

        int rowNum = startRowNumber;
        int batch = 0;
        while (reader.Read())
        {
            ct.ThrowIfCancellationRequested();
            for (int i = 0; i < colCount; i++)
            {
                var v = reader.GetValue(i);
                if (!smart)
                {
                    cmd.Parameters[i].Value = v == null || v == DBNull.Value ? DBNull.Value : Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";
                }
                else
                {
                    var (ok, val) = ConvertSmart(v, columns[i].DataType);
                    if (stats != null)
                    {
                        if (v != null && v != DBNull.Value && !string.IsNullOrWhiteSpace(Convert.ToString(v, CultureInfo.InvariantCulture))) stats.NonEmpty(i);
                        if (ok) stats.Ok(i); else stats.Fail(i, v);
                    }
                    cmd.Parameters[i].Value = val ?? DBNull.Value;
                }
            }
            cmd.ExecuteNonQuery();
            inserted++;
            batch++;
            rowNum++;

            if (batch >= 5000)
            {
                batch = 0;
                progress?.Report(new ImportProgress { Stage = "导入数据", Percentage = 20, CurrentRow = rowNum, TotalRows = 0 });
            }
        }
        tx.Commit();

        return inserted;
    }

    private static (bool ok, object? value) ConvertSmart(object? v, string targetType)
    {
        if (v == null || v == DBNull.Value) return (true, null);
        try
        {
            if (string.Equals(targetType, "Integer", StringComparison.OrdinalIgnoreCase))
            {
                if (v is sbyte or byte or short or ushort or int or uint or long or ulong) return (true, Convert.ToInt64(v, CultureInfo.InvariantCulture));
                var s = Convert.ToString(v, CultureInfo.InvariantCulture)?.Trim();
                if (long.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var x)) return (true, x);
                return (false, null);
            }
            if (string.Equals(targetType, "Number", StringComparison.OrdinalIgnoreCase))
            {
                if (v is float or double or decimal) return (true, Convert.ToDouble(v, CultureInfo.InvariantCulture));
                var s = Convert.ToString(v, CultureInfo.InvariantCulture)?.Trim();
                if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var x)) return (true, x);
                return (false, null);
            }
            if (string.Equals(targetType, "Boolean", StringComparison.OrdinalIgnoreCase))
            {
                if (v is bool b) return (true, b ? 1 : 0);
                var s = Convert.ToString(v, CultureInfo.InvariantCulture)?.Trim();
                if (bool.TryParse(s, out var x)) return (true, x ? 1 : 0);
                if (s == "1") return (true, 1);
                if (s == "0") return (true, 0);
                return (false, null);
            }
            if (string.Equals(targetType, "DateTime", StringComparison.OrdinalIgnoreCase))
            {
                if (v is DateTime dt) return (true, dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
                var s = Convert.ToString(v, CultureInfo.InvariantCulture)?.Trim();
                if (DateTime.TryParse(s, out var x)) return (true, x.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
                return (false, null);
            }
            // Text
            return (true, Convert.ToString(v, CultureInfo.InvariantCulture));
        }
        catch { return (false, null); }
    }

    private static string SanitizeColumnName(string name)
    {
        var s = (name ?? "").Trim();
        if (string.IsNullOrWhiteSpace(s)) return "Column";
        // 替换控制字符
        s = new string(s.Select(ch => char.IsControl(ch) ? '_' : ch).ToArray());
        // 去掉换行与制表
        s = s.Replace("\r", " ").Replace("\n", " ").Replace("\t", " ");
        // 最小规整
        return s;
    }

    private static string SanitizeTableName(string name)
        => string.IsNullOrWhiteSpace(name) ? "Main" : name.Trim();
}
