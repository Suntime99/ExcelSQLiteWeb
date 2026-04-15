using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelSQLiteWeb.Models;
using Microsoft.Data.Sqlite;

namespace ExcelSQLiteWeb.Services;

/// <summary>
/// CSV 格式专用的数据导入适配器
/// - 支持大文件流式读取
/// - 支持自动探测编码 (UTF-8, GBK 等)
/// - 支持基于采样的分隔符探测 (逗号, Tab, 分号, 竖线)
/// </summary>
public class CsvImporter : IDataImporter
{
    private readonly SqliteManager _sqliteManager;

    public CsvImporter(SqliteManager sqliteManager)
    {
        _sqliteManager = sqliteManager;
    }

    public bool CanHandle(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath)) return false;
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        return ext == ".csv" || ext == ".tsv" || ext == ".txt";
    }

    public string Kind => "csv";

    public IEnumerable<DatasetInfo> ListDatasets(string filePath)
    {
        if (!File.Exists(filePath)) yield break;
        // 对于 CSV 文件，只有一个数据集，即文件本身
        var datasetName = Path.GetFileNameWithoutExtension(filePath);
        yield return new DatasetInfo(datasetName, datasetName, null);
    }

    public DatasetSchema GetSchema(string filePath, string datasetId)
    {
        var schema = new DatasetSchema { DatasetId = datasetId };
        if (!File.Exists(filePath)) return schema;

        var encoding = DetectEncoding(filePath);
        var delimiter = DetectDelimiter(filePath, encoding);

        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = new StreamReader(stream, encoding);
        
        string? headerLine = reader.ReadLine();
        if (!string.IsNullOrWhiteSpace(headerLine))
        {
            var cols = ParseCsvLine(headerLine, delimiter);
            // 自动补齐空列名，防重复
            var finalCols = new List<string>();
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < cols.Count; i++)
            {
                var name = string.IsNullOrWhiteSpace(cols[i]) ? $"Column{i + 1}" : cols[i].Trim();
                int idx = 1;
                string finalName = name;
                while (set.Contains(finalName))
                {
                    finalName = $"{name}_{idx++}";
                }
                set.Add(finalName);
                finalCols.Add(finalName);
            }

            schema.Columns = finalCols.Select(n => new DatasetSchema.ColumnDef { Name = n, SqliteType = "TEXT" }).ToList();
        }
        
        return schema;
    }

#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
    public async IAsyncEnumerable<RowData> ReadRowsAsync(string filePath, string datasetId, ImportOptions options)
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
    {
        if (!File.Exists(filePath)) yield break;

        var encoding = DetectEncoding(filePath);
        var delimiter = DetectDelimiter(filePath, encoding);

        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = new StreamReader(stream, encoding);

        if (options.HasHeader)
        {
            reader.ReadLine(); // Skip header
        }

        string? line;
        while ((line = reader.ReadLine()) != null)
        {
            if (string.IsNullOrWhiteSpace(line)) continue;
            var cols = ParseCsvLine(line, delimiter);
            yield return new RowData { Values = cols.Cast<object?>().ToArray() };
        }
    }

    /// <summary>
    /// 简易编码探测（UTF-8 带/不带 BOM，GBK）
    /// </summary>
    private Encoding DetectEncoding(string filePath)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var gbk = Encoding.GetEncoding("GBK");
        
        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        if (fs.Length < 3) return gbk; // 极短文件默认 GBK 处理

        byte[] bom = new byte[4];
        fs.Read(bom, 0, 4);

        if (bom[0] == 0x2b && bom[1] == 0x2f && bom[2] == 0x76) return Encoding.UTF7;
        if (bom[0] == 0xef && bom[1] == 0xbb && bom[2] == 0xbf) return Encoding.UTF8;
        if (bom[0] == 0xff && bom[1] == 0xfe) return Encoding.Unicode; // UTF-16LE
        if (bom[0] == 0xfe && bom[1] == 0xff) return Encoding.BigEndianUnicode; // UTF-16BE
        if (bom[0] == 0 && bom[1] == 0 && bom[2] == 0xfe && bom[3] == 0xff) return Encoding.UTF32;

        // 没有 BOM，尝试简单启发式判断是否是 UTF-8
        fs.Position = 0;
        byte[] buffer = new byte[Math.Min(fs.Length, 4096)];
        int read = fs.Read(buffer, 0, buffer.Length);
        if (IsUtf8(buffer, read)) return Encoding.UTF8;

        return gbk;
    }

    private bool IsUtf8(byte[] buffer, int length)
    {
        int i = 0;
        bool hasNonAscii = false;
        while (i < length)
        {
            byte b = buffer[i];
            if (b <= 0x7F) { i++; continue; }
            hasNonAscii = true;
            if (b >= 0xC2 && b <= 0xDF) {
                if (i + 1 >= length || (buffer[i + 1] & 0xC0) != 0x80) return false;
                i += 2;
            } else if (b >= 0xE0 && b <= 0xEF) {
                if (i + 2 >= length || (buffer[i + 1] & 0xC0) != 0x80 || (buffer[i + 2] & 0xC0) != 0x80) return false;
                i += 3;
            } else if (b >= 0xF0 && b <= 0xF4) {
                if (i + 3 >= length || (buffer[i + 1] & 0xC0) != 0x80 || (buffer[i + 2] & 0xC0) != 0x80 || (buffer[i + 3] & 0xC0) != 0x80) return false;
                i += 4;
            } else {
                return false;
            }
        }
        return hasNonAscii;
    }

    /// <summary>
    /// 简易分隔符探测：采样前10行，统计不同分隔符的出现次数一致性
    /// </summary>
    private char DetectDelimiter(string filePath, Encoding encoding)
    {
        var delimiters = new[] { ',', '\t', ';', '|' };
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = new StreamReader(stream, encoding);

        var lines = new List<string>();
        for (int i = 0; i < 10; i++)
        {
            var line = reader.ReadLine();
            if (line != null) lines.Add(line);
        }
        if (!lines.Any()) return ',';

        foreach (var delim in delimiters)
        {
            var counts = lines.Select(l => l.Count(c => c == delim)).ToList();
            // 如果某分隔符在每行出现的次数相同且 > 0
            if (counts.All(c => c == counts.First()) && counts.First() > 0)
            {
                return delim;
            }
        }
        
        // 兜底方案，哪个出现次数最多选哪个
        var maxDelim = ',';
        int maxCount = -1;
        foreach (var delim in delimiters)
        {
            int c = lines.First().Count(x => x == delim);
            if (c > maxCount)
            {
                maxCount = c;
                maxDelim = delim;
            }
        }
        return maxDelim;
    }

    /// <summary>
    /// 简易 CSV 行解析（处理双引号转义）
    /// </summary>
    private List<string> ParseCsvLine(string line, char delimiter)
    {
        var result = new List<string>();
        bool inQuotes = false;
        var current = new StringBuilder();

        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (inQuotes)
            {
                if (c == '"')
                {
                    if (i + 1 < line.Length && line[i + 1] == '"')
                    {
                        current.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    current.Append(c);
                }
            }
            else
            {
                if (c == '"')
                {
                    inQuotes = true;
                }
                else if (c == delimiter)
                {
                    result.Add(current.ToString());
                    current.Clear();
                }
                else
                {
                    current.Append(c);
                }
            }
        }
        result.Add(current.ToString());
        return result;
    }

    public async Task<ImportResult> ImportWorksheetAsync(
        string filePath,
        string worksheetName,
        string? tableName = null,
        string importMode = "text",
        IProgress<ImportProgress>? progress = null,
        CancellationToken cancellationToken = default,
        bool append = false)
    {
        var sw = Stopwatch.StartNew();
        var datasetId = Path.GetFileNameWithoutExtension(filePath);
        tableName ??= SanitizeTableName(datasetId);

        var result = new ImportResult
        {
            FilePath = filePath,
            WorksheetName = datasetId,
            TableName = tableName,
            ImportMode = importMode
        };

        try
        {
            progress?.Report(new ImportProgress { Stage = "解析CSV结构", Percentage = 5 });
            var schema = GetSchema(filePath, datasetId);
            if (schema.Columns == null || schema.Columns.Count == 0)
                throw new Exception("无法解析CSV列或文件为空");

            // 创建表
            progress?.Report(new ImportProgress { Stage = "创建数据表", Percentage = 10 });
            var colDefs = string.Join(", ", schema.Columns.Select(c => $"{SqliteManager.QuoteIdent(c.Name)} {c.SqliteType}"));
            
            if (!append)
            {
                _sqliteManager.Execute($"DROP TABLE IF EXISTS {SqliteManager.QuoteIdent(tableName)}");
                _sqliteManager.Execute($"CREATE TABLE {SqliteManager.QuoteIdent(tableName)} ({colDefs});");
            }
            else
            {
                _sqliteManager.Execute($"CREATE TABLE IF NOT EXISTS {SqliteManager.QuoteIdent(tableName)} ({colDefs});");
                // 补齐可能缺失的列
                var existingCols = new HashSet<string>(_sqliteManager.GetTableColumns(tableName), StringComparer.OrdinalIgnoreCase);
                foreach (var c in schema.Columns)
                {
                    if (!existingCols.Contains(c.Name))
                    {
                        try { _sqliteManager.Execute($"ALTER TABLE {SqliteManager.QuoteIdent(tableName)} ADD COLUMN {SqliteManager.QuoteIdent(c.Name)} {c.SqliteType};"); } catch { }
                    }
                }
            }

            // 流式导入
            progress?.Report(new ImportProgress { Stage = "导入数据", Percentage = 15 });

            // 性能优化
            _sqliteManager.Execute("PRAGMA temp_store = MEMORY;");
            _sqliteManager.Execute("PRAGMA synchronous = OFF;");
            _sqliteManager.Execute("PRAGMA journal_mode = OFF;");

            var conn = _sqliteManager.Connection;
            if (conn == null) throw new InvalidOperationException("SQLite 连接未打开");

            int totalInserted = 0;
            int batchSize = 5000;
            var batch = new List<RowData>(batchSize);

            // 预构建 SQL
            var colNamesStr = string.Join(", ", schema.Columns.Select(c => SqliteManager.QuoteIdent(c.Name)));
            var paramNamesStr = string.Join(", ", schema.Columns.Select((_, i) => $"@p{i}"));
            var insertSql = $"INSERT INTO {SqliteManager.QuoteIdent(tableName)} ({colNamesStr}) VALUES ({paramNamesStr});";

            async Task FlushBatch()
            {
                if (batch.Count == 0) return;
                using var tx = conn.BeginTransaction();
                using var cmd = conn.CreateCommand();
                cmd.Transaction = tx;
                cmd.CommandText = insertSql;

                var pms = new SqliteParameter[schema.Columns.Count];
                for (int i = 0; i < schema.Columns.Count; i++)
                {
                    var p = cmd.CreateParameter();
                    p.ParameterName = $"@p{i}";
                    cmd.Parameters.Add(p);
                    pms[i] = p;
                }

                foreach (var row in batch)
                {
                    for (int i = 0; i < schema.Columns.Count; i++)
                    {
                        var val = i < row.Values.Length ? row.Values[i] : null;
                        pms[i].Value = val ?? DBNull.Value;
                    }
                    cmd.ExecuteNonQuery();
                }

                tx.Commit();
                totalInserted += batch.Count;
                batch.Clear();

                progress?.Report(new ImportProgress
                {
                    Stage = "导入数据",
                    Percentage = 15 + (int)Math.Min(80, totalInserted / 1000.0),
                    CurrentRow = totalInserted,
                    TotalRows = -1
                });
            }

            var rows = ReadRowsAsync(filePath, datasetId, new ImportOptions { HasHeader = true });
            await foreach (var row in rows.WithCancellation(cancellationToken))
            {
                batch.Add(row);
                if (batch.Count >= batchSize)
                {
                    await FlushBatch();
                }
            }
            await FlushBatch();

            _sqliteManager.Execute("PRAGMA synchronous = NORMAL;");
            _sqliteManager.Execute("PRAGMA journal_mode = WAL;");

            sw.Stop();
            result.Success = true;
            result.RowsInserted = totalInserted;
            result.Message = $"导入成功 (CSV)，耗时: {sw.ElapsedMilliseconds}ms";
            
            progress?.Report(new ImportProgress { Stage = "完成", Percentage = 100, CurrentRow = totalInserted, TotalRows = totalInserted });
        }
        catch (Exception ex)
        {
            sw.Stop();
            result.Success = false;
            result.Message = $"导入失败: {ex.Message}";
            progress?.Report(new ImportProgress { Stage = "失败", Percentage = 100 });
        }

        return result;
    }

    private string SanitizeTableName(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return "Table1";
        var invalidChars = Path.GetInvalidFileNameChars().Concat(new[] { ' ', '-', '.', ',' }).ToArray();
        var safe = new string(name.Select(c => invalidChars.Contains(c) ? '_' : c).ToArray());
        if (char.IsDigit(safe[0])) safe = "_" + safe;
        return safe;
    }
}
