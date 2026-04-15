using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelSQLiteWeb.Services
{
    public class BatchEngine
    {
        private readonly SqliteManager _sqliteManager;
        private readonly Action<int, string> _onProgress;

        public BatchEngine(SqliteManager sqliteManager, Action<int, string> onProgress)
        {
            _sqliteManager = sqliteManager ?? throw new ArgumentNullException(nameof(sqliteManager));
            _onProgress = onProgress ?? throw new ArgumentNullException(nameof(onProgress));
        }

        public async Task<BatchResult> ExecuteBatchAsync(System.Text.Json.JsonElement settings)
        {
            var result = new BatchResult();
            var startTime = DateTime.Now;

            try
            {
                // Parse settings
                string fileSelectType = settings.TryGetProperty("fileSelectType", out var fst) ? fst.GetString() ?? "folder" : "folder";
                string folderPath = settings.TryGetProperty("folderPath", out var fp) ? fp.GetString() ?? "" : "";
                string pattern = settings.TryGetProperty("pattern", out var pt) ? pt.GetString() ?? "*.*" : "*.*";
                string taskType = settings.TryGetProperty("taskType", out var tt) ? tt.GetString() ?? "" : "";
                string outputType = settings.TryGetProperty("outputType", out var ot) ? ot.GetString() ?? "new" : "new";
                string outputFolder = settings.TryGetProperty("outputFolder", out var of) ? of.GetString() ?? "" : "";

                bool useParallel = false;
                bool generateLog = false;
                if (settings.TryGetProperty("advancedOptions", out var adv))
                {
                    if (adv.TryGetProperty("parallel", out var p)) useParallel = p.GetBoolean();
                    if (adv.TryGetProperty("log", out var lg)) generateLog = lg.GetBoolean();
                }

                result.TaskType = taskType;

                // Discover files
                var targetFiles = new List<string>();
                if (fileSelectType == "folder")
                {
                    if (!string.IsNullOrWhiteSpace(folderPath) && Directory.Exists(folderPath))
                    {
                        var filters = new List<string>();
                        if (settings.TryGetProperty("filters", out var flts))
                        {
                            if (flts.TryGetProperty("xlsx", out var x1) && x1.GetBoolean()) filters.Add(".xlsx");
                            if (flts.TryGetProperty("xls", out var x2) && x2.GetBoolean()) filters.Add(".xls");
                            if (flts.TryGetProperty("csv", out var x3) && x3.GetBoolean()) filters.Add(".csv");
                        }
                        if (filters.Count == 0) filters.AddRange(new[] { ".xlsx", ".xls", ".csv" });

                        foreach (var f in Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly))
                        {
                            if (filters.Any(ext => f.EndsWith(ext, StringComparison.OrdinalIgnoreCase)))
                            {
                                targetFiles.Add(f);
                            }
                        }
                    }
                }
                else if (fileSelectType == "pattern")
                {
                    if (!string.IsNullOrWhiteSpace(folderPath) && Directory.Exists(folderPath))
                    {
                        var regex = new Regex("^" + Regex.Escape(pattern).Replace("\\*", ".*").Replace("\\?", ".") + "$", RegexOptions.IgnoreCase);
                        foreach (var f in Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly))
                        {
                            if (regex.IsMatch(Path.GetFileName(f)))
                            {
                                targetFiles.Add(f);
                            }
                        }
                    }
                }

                result.TotalFiles = targetFiles.Count;

                if (targetFiles.Count == 0)
                {
                    _onProgress(100, "没有找到符合条件的文件");
                    result.ProcessTime = 0;
                    return result;
                }

                if (outputType == "folder" && !string.IsNullOrWhiteSpace(outputFolder) && !Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                }
                result.OutputPath = outputType == "folder" ? outputFolder : "与源文件同目录";

                var logLines = new List<string>();
                if (generateLog) logLines.Add($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 开始批量任务: {taskType}, 总文件数: {targetFiles.Count}");

                int processed = 0;
                object syncLock = new object();

                var parallelOptions = new ParallelOptions { MaxDegreeOfParallelism = useParallel ? Environment.ProcessorCount : 1 };

                await Parallel.ForEachAsync(targetFiles, parallelOptions, async (file, ct) =>
                {
                    bool success = true;
                    string logMsg = $"成功";
                    try
                    {
                        // 模拟任务执行，后续可根据 taskType 接入具体的引擎逻辑
                        await Task.Delay(100, ct); 
                        
                        // TODO: 这里接入具体的清理、脱敏、分析逻辑
                    }
                    catch (Exception ex)
                    {
                        success = false;
                        logMsg = $"失败: {ex.Message}";
                    }

                    lock (syncLock)
                    {
                        processed++;
                        if (success) result.SuccessCount++;
                        else result.FailureCount++;

                        if (generateLog) logLines.Add($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 文件: {Path.GetFileName(file)} -> {logMsg}");

                        int pct = (int)((processed * 100.0) / targetFiles.Count);
                        _onProgress(pct, $"正在处理 ({processed}/{targetFiles.Count}): {Path.GetFileName(file)}");
                    }
                });

                if (generateLog)
                {
                    string logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs", "Batch");
                    Directory.CreateDirectory(logDir);
                    string logPath = Path.Combine(logDir, $"BatchLog_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                    File.WriteAllLines(logPath, logLines);
                    result.LogPath = logPath;
                }

                _onProgress(100, "批量处理完成");
            }
            catch (Exception ex)
            {
                _onProgress(100, $"发生错误: {ex.Message}");
            }
            finally
            {
                result.ProcessTime = (DateTime.Now - startTime).TotalSeconds;
            }

            return result;
        }
    }

    public class BatchResult
    {
        [System.Text.Json.Serialization.JsonPropertyName("taskType")]
        public string TaskType { get; set; } = "";
        [System.Text.Json.Serialization.JsonPropertyName("totalFiles")]
        public int TotalFiles { get; set; } = 0;
        [System.Text.Json.Serialization.JsonPropertyName("successCount")]
        public int SuccessCount { get; set; } = 0;
        [System.Text.Json.Serialization.JsonPropertyName("failureCount")]
        public int FailureCount { get; set; } = 0;
        [System.Text.Json.Serialization.JsonPropertyName("processTime")]
        public double ProcessTime { get; set; } = 0;
        [System.Text.Json.Serialization.JsonPropertyName("outputPath")]
        public string OutputPath { get; set; } = "";
        [System.Text.Json.Serialization.JsonPropertyName("logPath")]
        public string LogPath { get; set; } = "";
    }
}
