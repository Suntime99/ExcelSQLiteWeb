using ExcelSQLiteWeb.Models;
using ExcelSQLiteWeb.Services;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace ExcelSQLiteWeb;

public partial class Form1 : Form
{
    private ExcelAnalyzer? _excelAnalyzer;
    private SqliteManager? _sqliteManager;
    private DataImporter? _dataImporter;
    private QueryEngine? _queryEngine;
    private StatisticsEngine? _statisticsEngine;
    private SplitEngine? _splitEngine;
    private SqlScriptGenerator? _sqlScriptGenerator;

    private string? _currentFilePath;
    private FileAnalysisResult? _currentAnalysisResult;
    private string? _currentWorksheetName;
    // 最近一次导入（智能转换）产生的异常统计，用于“SQLite分析”展示数据质量
    private List<ConversionStat> _lastConversionStats = new();
    // Excel 模式 SQL：已导入到 SQLite 的工作表（表名=工作表名）
    private readonly HashSet<string> _excelSqliteImportedTables = new(StringComparer.OrdinalIgnoreCase);

    // 最近一次错误日志（用于“错误记录下载”）
    private string? _lastErrorLogPath;

    // 最近文件（持久化到本机设置文件）
    private const int MaxRecentFiles = 10;
    private readonly List<RecentFileRecord> _recentFiles = new();

    public Form1()
    {
        InitializeComponent();
        InitializeServices();
        LoadRecentFilesFromDisk();
        InitializeWebView2();
    }

    private sealed class RecentFileRecord
    {
        public string FullPath { get; set; } = string.Empty;
        public DateTime LastOpenTime { get; set; } = DateTime.Now;
    }

    private string GetRecentFilesStorePath()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelSQLiteWeb");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, "recent-files.json");
    }

    private void LoadRecentFilesFromDisk()
    {
        try
        {
            var path = GetRecentFilesStorePath();
            if (!File.Exists(path)) return;

            var json = File.ReadAllText(path, Encoding.UTF8);
            var list = System.Text.Json.JsonSerializer.Deserialize<List<RecentFileRecord>>(json)
                       ?? new List<RecentFileRecord>();

            _recentFiles.Clear();
            _recentFiles.AddRange(
                list.Where(x => !string.IsNullOrWhiteSpace(x.FullPath))
                    .OrderByDescending(x => x.LastOpenTime)
                    .Take(MaxRecentFiles));
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"LoadRecentFilesFromDisk failed: {ex.Message}");
        }
    }

    private void SaveRecentFilesToDisk()
    {
        try
        {
            var path = GetRecentFilesStorePath();
            var json = System.Text.Json.JsonSerializer.Serialize(_recentFiles);
            File.WriteAllText(path, json, Encoding.UTF8);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"SaveRecentFilesToDisk failed: {ex.Message}");
        }
    }

    private object BuildRecentFilesPayload()
    {
        // 前端只需要展示文件名+时间；打开时仍以 fileName 发回（宿主内部再匹配 FullPath）
        var recentFiles = _recentFiles
            .OrderByDescending(x => x.LastOpenTime)
            .Take(MaxRecentFiles)
            .Select(x => new
            {
                name = Path.GetFileName(x.FullPath),
                time = x.LastOpenTime.ToString("yyyy-MM-dd HH:mm")
            })
            .ToArray();

        return recentFiles;
    }

    private void TouchRecentFile(string fullPath)
    {
        if (string.IsNullOrWhiteSpace(fullPath)) return;

        var existing = _recentFiles.FirstOrDefault(x =>
            string.Equals(x.FullPath, fullPath, StringComparison.OrdinalIgnoreCase));

        if (existing != null)
        {
            existing.LastOpenTime = DateTime.Now;
        }
        else
        {
            _recentFiles.Add(new RecentFileRecord
            {
                FullPath = fullPath,
                LastOpenTime = DateTime.Now
            });
        }

        _recentFiles.Sort((a, b) => b.LastOpenTime.CompareTo(a.LastOpenTime));
        if (_recentFiles.Count > MaxRecentFiles)
            _recentFiles.RemoveRange(MaxRecentFiles, _recentFiles.Count - MaxRecentFiles);

        SaveRecentFilesToDisk();
    }

    private void NotifyRecentFilesUpdated()
    {
        SendMessageToWebView(new
        {
            action = "recentFilesUpdated",
            recentFiles = BuildRecentFilesPayload()
        });
    }

    private void InitializeServices()
    {
        _excelAnalyzer = new ExcelAnalyzer();
        _sqliteManager = new SqliteManager(); // 内存数据库
        _dataImporter = new DataImporter(_excelAnalyzer, _sqliteManager);
        _queryEngine = new QueryEngine(_sqliteManager);
        _statisticsEngine = new StatisticsEngine(_sqliteManager);
        _splitEngine = new SplitEngine(_sqliteManager, _excelAnalyzer);
        _sqlScriptGenerator = new SqlScriptGenerator();
    }

    private async void InitializeWebView2()
    {
        try
        {
            System.Diagnostics.Debug.WriteLine("开始初始化WebView2");
            await webView21.EnsureCoreWebView2Async(null);
            System.Diagnostics.Debug.WriteLine("WebView2初始化成功");

            // 启用 WebMessage
            webView21.CoreWebView2.Settings.IsWebMessageEnabled = true;
            System.Diagnostics.Debug.WriteLine("WebMessage已启用");

            // 页面加载完成后下发初始化数据（最近文件等）
            webView21.CoreWebView2.NavigationCompleted += (_, __) =>
            {
                SendMessageToWebView(new
                {
                    action = "initState",
                    recentFiles = BuildRecentFilesPayload()
                });
            };

            // 优先从磁盘读取（便于“就近覆盖”快速迭代）；若不存在再回退到嵌入式资源
            string htmlContent = "";
            string htmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "index.html");
            if (!File.Exists(htmlPath))
            {
                // 兼容调试目录：向上三级
                var htmlPath2 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "index.html");
                if (File.Exists(htmlPath2)) htmlPath = htmlPath2;
            }

            if (File.Exists(htmlPath))
            {
                htmlContent = File.ReadAllText(htmlPath, Encoding.UTF8);
                System.Diagnostics.Debug.WriteLine($"成功从磁盘读取HTML文件: {htmlPath}");
            }
            else
            {
                try
                {
                    htmlContent = GetEmbeddedResource("ExcelSQLiteWeb.index.html");
                    System.Diagnostics.Debug.WriteLine("成功从嵌入式资源加载HTML文件");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"HTML加载失败: {ex.Message}");
                    htmlContent = $"<html><body><h1>错误</h1><p>无法找到index.html文件</p><p>当前目录: {AppDomain.CurrentDomain.BaseDirectory}</p></body></html>";
                }
            }

            webView21.NavigateToString(htmlContent);
            System.Diagnostics.Debug.WriteLine("HTML内容加载完成");

            webView21.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;
            System.Diagnostics.Debug.WriteLine("WebMessageReceived事件注册成功");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"WebView2初始化异常: {ex.Message}");
        }
    }

    private string GetEmbeddedResource(string resourceName)
    {
        var assembly = System.Reflection.Assembly.GetExecutingAssembly();
        using var stream = assembly.GetManifestResourceStream(resourceName);
        if (stream == null)
        {
            throw new InvalidOperationException($"Resource {resourceName} not found");
        }
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    private void CoreWebView2_WebMessageReceived(object sender, Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs e)
    {
        // 用 JSON 接收对象消息（关键）
        var json = e.WebMessageAsJson;
        System.Diagnostics.Debug.WriteLine($"Received json: {json}");

        try
        {
            using var doc = System.Text.Json.JsonDocument.Parse(json);
            var data = doc.RootElement;

            // 有时 json 可能是一个字符串（比如 "\"{...}\""），这里做兼容
            if (data.ValueKind == System.Text.Json.JsonValueKind.String)
            {
                var s = data.GetString() ?? "";
                using var doc2 = System.Text.Json.JsonDocument.Parse(s);
                data = doc2.RootElement;
            }

            string action = data.GetProperty("action").GetString() ?? string.Empty;
            System.Diagnostics.Debug.WriteLine($"Processing action: {action}");

            switch (action)
                {
                    case "browseFile":
                        System.Diagnostics.Debug.WriteLine("Calling BrowseFile()");
                        BrowseFile();
                        break;
                    case "browseQueryFile":
                        System.Diagnostics.Debug.WriteLine("Calling BrowseQueryFile()");
                        BrowseQueryFile();
                        break;
                    case "browseFolder":
                        System.Diagnostics.Debug.WriteLine("Calling BrowseFolder()");
                        BrowseFolder();
                        break;
                    case "browseMainTableFile":
                        System.Diagnostics.Debug.WriteLine("Calling BrowseMainTableFile()");
                        BrowseMainTableFile();
                        break;
                    case "browseSubTableFile":
                        string subTableId = data.GetProperty("subTableId").GetString() ?? string.Empty;
                        System.Diagnostics.Debug.WriteLine($"Calling BrowseSubTableFile() with subTableId: {subTableId}");
                        BrowseSubTableFile(subTableId);
                        break;
                    case "browseAiFile":
                        System.Diagnostics.Debug.WriteLine("Calling BrowseFile() for AI file");
                        BrowseFile();
                        break;
                    case "browseDcOutputPath":
                        System.Diagnostics.Debug.WriteLine("Calling BrowseFolder() for DC output path");
                        BrowseFolder();
                        break;
                    case "browseCompareSourceFile":
                        System.Diagnostics.Debug.WriteLine("Calling BrowseFile() for compare source");
                        BrowseFile();
                        break;
                    case "browseCompareTargetFile":
                        System.Diagnostics.Debug.WriteLine("Calling BrowseFile() for compare target");
                        BrowseFile();
                        break;
                    case "browseBatchFolder":
                        System.Diagnostics.Debug.WriteLine("Calling BrowseFolder() for batch folder");
                        BrowseFolder();
                        break;
                    case "browseBatchFiles":
                        System.Diagnostics.Debug.WriteLine("Calling BrowseFile() for batch files");
                        BrowseFile();
                        break;
                    case "browseBatchPatternFolder":
                        System.Diagnostics.Debug.WriteLine("Calling BrowseFolder() for batch pattern");
                        BrowseFolder();
                        break;
                    case "browseBatchOutputFolder":
                        System.Diagnostics.Debug.WriteLine("Calling BrowseFolder() for batch output");
                        BrowseFolder();
                        break;
                case "downloadErrorLog":
                    DownloadErrorLog();
                    break;
                case "startAnalysis":
                    string filePath = data.GetProperty("filePath").GetString() ?? string.Empty;
                    StartAnalysis(filePath);
                    break;
                case "startAnalysisSqlite":
                    {
                        string tableName = (data.TryGetProperty("tableName", out var tn) ? tn.GetString() : null) ?? "Main";
                        StartAnalysisSqlite(tableName);
                        break;
                    }
                case "analyzeWorksheet":
                    {
                        string sheetName = data.GetProperty("sheetName").GetString() ?? string.Empty;
                        string sheetFilePath =
                            (data.TryGetProperty("filePath", out var fp) ? fp.GetString() : null)
                            ?? _currentFilePath
                            ?? string.Empty;
                        AnalyzeWorksheet(sheetName, sheetFilePath);
                        break;
                    }
                case "analyzeSqliteWorksheet":
                    {
                        string tableName = (data.TryGetProperty("tableName", out var tn) ? tn.GetString() : null) ?? "Main";
                        AnalyzeSqliteWorksheet(tableName);
                        break;
                    }
                case "startMetadataScan":
                    StartMetadataScan(data);
                    break;
                case "executeQuery":
                    ExecuteQuery(data);
                    break;
                case "executeGlobalSearch":
                    ExecuteGlobalSearch(data);
                    break;
                case "executeSingleTableStats":
                    ExecuteSingleTableStats(data);
                    break;
                case "executeMultiTableStats":
                    ExecuteMultiTableStats(data);
                    break;
                case "executeSplit":
                    ExecuteSplit(data);
                    break;
                case "importWorksheet":
                    ImportWorksheet(data);
                    break;
                case "openRecentFile":
                    string recentFileName = data.GetProperty("fileName").GetString() ?? string.Empty;
                    OpenRecentFile(recentFileName);
                    break;
                case "applyTemplate":
                    string templateName = data.GetProperty("templateName").GetString() ?? string.Empty;
                    ApplyTemplate(templateName);
                    break;
                case "openTool":
                    string toolName = data.GetProperty("toolName").GetString() ?? string.Empty;
                    OpenTool(toolName);
                    break;
                case "showLearningCases":
                    ShowLearningCases();
                    break;
                case "showHelp":
                    ShowHelp();
                    break;
                case "showAbout":
                    ShowAbout();
                    break;
                case "getWorksheetFields":
                    GetWorksheetFields(data);
                    break;
                case "getSqliteTableFields":
                    GetSqliteTableFields(data);
                    break;
                case "getActiveFields":
                    GetActiveFields(data);
                    break;
                case "getTableList":
                    GetTableList();
                    break;
                case "executeSqlEditor":
                    ExecuteSqlEditor(data);
                    break;
                case "executeMultiTableQuery":
                    System.Diagnostics.Debug.WriteLine("Executing multi-table query");
                    ExecuteQuery(data);
                    break;
                case "exportSqlResult":
                    System.Diagnostics.Debug.WriteLine("Exporting SQL result");
                    // 这里可以添加导出 SQL 结果的逻辑
                    break;
                case "exportQueryToFile":
                    ExportQueryToFile(data);
                    break;
                case "exportChart":
                    System.Diagnostics.Debug.WriteLine("Exporting chart");
                    // 这里可以添加导出图表的逻辑
                    break;
                case "refreshData":
                    System.Diagnostics.Debug.WriteLine("Refreshing data");
                    // 这里可以添加刷新数据的逻辑
                    break;
                case "openCleansedFile":
                    System.Diagnostics.Debug.WriteLine("Opening cleansed file");
                    // 这里可以添加打开清洗后文件的逻辑
                    break;
                case "viewCleansingReport":
                    System.Diagnostics.Debug.WriteLine("Viewing cleansing report");
                    // 这里可以添加查看清洗报告的逻辑
                    break;
                case "openBatchLog":
                    System.Diagnostics.Debug.WriteLine("Opening batch log");
                    // 这里可以添加打开批处理日志的逻辑
                    break;
                case "openBatchOutputFolder":
                    System.Diagnostics.Debug.WriteLine("Opening batch output folder");
                    // 这里可以添加打开批处理输出文件夹的逻辑
                    break;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error processing message: {ex.Message}");
            SendMessageToWebView(new { action = "error", message = ex.Message });
        }
    }

    private void BrowseFile()
    {
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
            openFileDialog.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                _currentFilePath = filePath;
                _excelSqliteImportedTables.Clear();
                TouchRecentFile(filePath);
                NotifyRecentFilesUpdated();
                SendMessageToWebView(new { action = "fileSelected", filePath = filePath.Replace('\\', '/') });
                LoadWorksheetList(filePath);
            }
        }
    }

    private void BrowseQueryFile()
    {
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
            openFileDialog.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                _currentFilePath = filePath;
                _excelSqliteImportedTables.Clear();
                TouchRecentFile(filePath);
                NotifyRecentFilesUpdated();
                SendMessageToWebView(new { action = "queryFileSelected", filePath = filePath.Replace('\\', '/') });
                LoadWorksheetList(filePath);
            }
        }
    }

    private void BrowseFolder()
    {
        using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
        {
            folderDialog.Description = "选择包含Excel文件的文件夹";
            folderDialog.ShowNewFolderButton = false;

            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                string folderPath = folderDialog.SelectedPath;
                SendMessageToWebView(new { action = "folderSelected", folderPath = folderPath.Replace('\\', '/') });
            }
        }
    }

    private void BrowseMainTableFile()
    {
        System.Diagnostics.Debug.WriteLine("BrowseMainTableFile() called");
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
            openFileDialog.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;

            System.Diagnostics.Debug.WriteLine("Showing OpenFileDialog...");
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                _currentFilePath = filePath;
                _excelSqliteImportedTables.Clear();
                System.Diagnostics.Debug.WriteLine($"File selected: {filePath}");
                TouchRecentFile(filePath);
                NotifyRecentFilesUpdated();
                SendMessageToWebView(new { action = "mainTableFileSelected", filePath = filePath.Replace('\\', '/') });
                System.Diagnostics.Debug.WriteLine("Message sent to webview: mainTableFileSelected");
                LoadWorksheetList(filePath);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("OpenFileDialog canceled");
            }
        }
    }

    private void BrowseSubTableFile(string subTableId)
    {
        System.Diagnostics.Debug.WriteLine($"BrowseSubTableFile() called with subTableId: {subTableId}");
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
            openFileDialog.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;

            System.Diagnostics.Debug.WriteLine("Showing OpenFileDialog...");
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                System.Diagnostics.Debug.WriteLine($"File selected: {filePath}");
                TouchRecentFile(filePath);
                NotifyRecentFilesUpdated();
                SendMessageToWebView(new { action = "subTableFileSelected", subTableId = subTableId, filePath = filePath.Replace('\\', '/') });
                System.Diagnostics.Debug.WriteLine("Message sent to webview: subTableFileSelected");
                LoadWorksheetList(filePath);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("OpenFileDialog canceled");
            }
        }
    }

    /// <summary>
    /// 快速读取工作表名（不使用EPPlus，避免大文件加载时卡顿）
    /// xlsx 本质是 zip：只解析 xl/workbook.xml
    /// </summary>
    private static List<string> GetWorksheetNamesFast(string filePath)
    {
        // 只支持 xlsx；其他格式回退到空列表（由上层决定如何处理）
        if (!filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            return new List<string>();

        using var archive = ZipFile.OpenRead(filePath);
        var entry = archive.GetEntry("xl/workbook.xml");
        if (entry == null) return new List<string>();

        using var stream = entry.Open();
        var doc = XDocument.Load(stream);

        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        return doc.Descendants(ns + "sheet")
            .Select(e => (string?)e.Attribute("name"))
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .Distinct()
            .ToList()!;
    }

    private async void LoadWorksheetList(string filePath)
    {
        try
        {
            // 后台线程读取，避免UI假死
            var worksheets = await Task.Run(() =>
            {
                // 大文件优先走快速路径
                var fast = GetWorksheetNamesFast(filePath);
                if (fast.Count > 0) return fast;
                // 回退：使用现有EPPlus逻辑（可能较慢）
                return _excelAnalyzer?.GetWorksheetNames(filePath) ?? new List<string>();
            });

            BeginInvoke(new Action(() =>
            {
                SendMessageToWebView(new { action = "worksheetListLoaded", worksheets });
            }));
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"加载工作表列表失败: {ex.Message}" });
        }
    }

    private async void StartAnalysis(string filePath)
    {
        try
        {
            System.Diagnostics.Debug.WriteLine($"Starting analysis for file: {filePath}");

            if (_excelAnalyzer == null)
            {
                SendMessageToWebView(new { action = "error", message = "Excel分析器未初始化" });
                return;
            }

            // 避免阻塞UI线程导致“假死”
            var analysisResult = await Task.Run(() => _excelAnalyzer.Analyze(filePath));
            _currentAnalysisResult = analysisResult;

            var results = new
            {
                filePath = filePath,
                fileName = analysisResult.FileName,
                fileSize = FormatFileSize(analysisResult.FileSize),
                totalRows = analysisResult.TotalRowCount,
                sheetCount = analysisResult.WorksheetCount,
                fieldCount = analysisResult.TotalColumnCount,
                analyzeTime = analysisResult.AnalyzeTime,
                dataQuality = $"{(analysisResult.DataQualityScore):F1}% 完整",
                sheets = analysisResult.Worksheets.Select(s => new
                {
                    name = s.Name,
                    rowCount = s.RowCount,
                    colCount = s.ColumnCount,
                    size = s.Size,
                    completeness = s.Completeness
                }).ToArray()
            };

            // WebView2调用建议在UI线程
            BeginInvoke(new Action(() =>
            {
                SendMessageToWebView(new { action = "analysisComplete", results });
            }));
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"分析失败: {ex.Message}" });
        }
    }

    /// <summary>
    /// 文件分析（SQLite 数据源）：对已导入的表做快速统计
    /// </summary>
    private async void StartAnalysisSqlite(string tableName)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }

            var sw = Stopwatch.StartNew();
            var results = await Task.Run(() =>
            {
                var schema = _sqliteManager.GetTableSchema(tableName);
                int colCount = schema.Count;
                int rowCount = _sqliteManager.GetRowCount(tableName);

                // 使用导入统计估算“完整率”（智能转换模式下更有意义）
                double completeness = 100;
                if (_lastConversionStats != null && _lastConversionStats.Count > 0)
                {
                    long nonEmpty = _lastConversionStats.Sum(s => (long)s.nonEmptyCount);
                    long ok = _lastConversionStats.Sum(s => (long)s.okCount);
                    if (nonEmpty > 0) completeness = Math.Round(ok * 100.0 / nonEmpty, 1);
                }

                return new
                {
                    fileName = Path.GetFileName(_currentFilePath ?? string.Empty),
                    fileSize = "",
                    totalRows = rowCount,
                    sheetCount = 1,
                    fieldCount = colCount,
                    analyzeTime = 0.0,
                    dataQuality = $"{completeness:F1}%（SQLite导入统计）",
                    sheets = new[]
                    {
                        new
                        {
                            name = tableName,
                            rowCount = rowCount,
                            colCount = colCount,
                            size = "",
                            completeness = completeness
                        }
                    }
                };
            });

            sw.Stop();
            var resultsWithTime = new
            {
                results.fileName,
                results.fileSize,
                results.totalRows,
                results.sheetCount,
                results.fieldCount,
                analyzeTime = Math.Round(sw.Elapsed.TotalSeconds, 2),
                results.dataQuality,
                results.sheets
            };

            BeginInvoke(new Action(() =>
            {
                SendMessageToWebView(new { action = "analysisComplete", results = resultsWithTime });
            }));
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"SQLite分析失败: {ex.Message}" });
        }
    }

    private void AnalyzeWorksheet(string sheetName, string filePath)
    {
        try
        {
            if (string.IsNullOrEmpty(filePath) || _excelAnalyzer == null)
            {
                SendMessageToWebView(new { action = "error", message = "请先选择文件" });
                return;
            }

            var fields = _excelAnalyzer.GetWorksheetFields(filePath, sheetName);
            var dataPreview = _excelAnalyzer.GetWorksheetData(filePath, sheetName, maxRows: 10);

            var result = new
            {
                sheetName,
                fields,
                preview = dataPreview
            };

            SendMessageToWebView(new { action = "worksheetAnalysisComplete", result });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"工作表分析失败: {ex.Message}" });
        }
    }

    /// <summary>
    /// 工作表解析（SQLite 数据源）：对表做字段+预览+行数统计
    /// </summary>
    private async void AnalyzeSqliteWorksheet(string tableName)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }

            var result = await Task.Run(() =>
            {
                var schema = _sqliteManager.GetTableSchema(tableName);
                var fields = schema.Select(s => s.ColumnName).ToList();
                int totalRows = _sqliteManager.GetRowCount(tableName);

                // 预览：前10行
                var preview = _sqliteManager.Query($"SELECT * FROM [{tableName}] LIMIT 10");

                // 完整率：沿用导入统计（若有）
                double completeness = 100;
                if (_lastConversionStats != null && _lastConversionStats.Count > 0)
                {
                    long nonEmpty = _lastConversionStats.Sum(s => (long)s.nonEmptyCount);
                    long ok = _lastConversionStats.Sum(s => (long)s.okCount);
                    if (nonEmpty > 0) completeness = Math.Round(ok * 100.0 / nonEmpty, 1);
                }

                return new
                {
                    sheetName = tableName,
                    fields = fields,
                    preview = preview,
                    totalRows = totalRows,
                    size = "",
                    completeness = completeness
                };
            });

            BeginInvoke(new Action(() =>
            {
                SendMessageToWebView(new { action = "worksheetAnalysisComplete", result });
            }));
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"SQLite工作表解析失败: {ex.Message}" });
        }
    }

    private async void StartMetadataScan(System.Text.Json.JsonElement data)
    {
        try
        {
            string scope = (data.TryGetProperty("scope", out var sc) ? sc.GetString() : null) ?? "currentFile";
            string folderPath = (data.TryGetProperty("folderPath", out var fp) ? fp.GetString() : null) ?? string.Empty;

            var sw = Stopwatch.StartNew();
            var items = await Task.Run(() =>
            {
                var list = new List<object>();

                IEnumerable<string> files = Enumerable.Empty<string>();
                if (string.Equals(scope, "folder", StringComparison.OrdinalIgnoreCase))
                {
                    if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
                        throw new DirectoryNotFoundException("请选择有效的文件夹路径");
                    files = Directory.EnumerateFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(_currentFilePath) || !File.Exists(_currentFilePath))
                        throw new FileNotFoundException("请先选择 Excel 文件");
                    files = new[] { _currentFilePath };
                }

                foreach (var f in files)
                {
                    try
                    {
                        var fi = new FileInfo(f);
                        var sheets = GetWorksheetNamesFast(f);
                        list.Add(new
                        {
                            fileName = fi.Name,
                            filePath = f.Replace('\\', '/'),
                            sheetCount = sheets.Count,
                            totalRows = 0,
                            fileSize = FormatFileSize(fi.Length),
                            lastModified = fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm")
                        });
                    }
                    catch (Exception ex)
                    {
                        list.Add(new
                        {
                            fileName = Path.GetFileName(f),
                            filePath = f.Replace('\\', '/'),
                            sheetCount = 0,
                            totalRows = 0,
                            fileSize = "",
                            lastModified = "",
                            error = ex.Message
                        });
                    }
                }

                return list;
            });
            sw.Stop();

            var result = new
            {
                scope = scope,
                scopeLabel = string.Equals(scope, "folder", StringComparison.OrdinalIgnoreCase) ? "指定文件夹" : "当前文件",
                scanTime = Math.Round(sw.Elapsed.TotalSeconds, 2),
                items = items
            };

            SendMessageToWebView(new { action = "metadataScanComplete", results = result });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"元数据扫描失败: {ex.Message}" });
        }
    }

    private async void ImportWorksheet(System.Text.Json.JsonElement data)
    {
        try
        {
            string worksheetName = data.GetProperty("worksheetName").GetString() ?? string.Empty;
            _currentWorksheetName = worksheetName;

            string importMode =
                (data.TryGetProperty("importMode", out var im) ? im.GetString() : null)
                ?? "text";

            bool resetDb = data.TryGetProperty("resetDb", out var rd) && rd.ValueKind == System.Text.Json.JsonValueKind.True;

            if (string.IsNullOrEmpty(_currentFilePath) || _dataImporter == null)
            {
                SendMessageToWebView(new { action = "error", message = "请先选择文件" });
                return;
            }

            // 导入前清空 SQLite（避免表混杂/旧数据误用）
            if (resetDb)
            {
                ResetSqliteDatabase();
                _excelSqliteImportedTables.Clear();
            }

            // 进度回调：降低WebView消息频率，避免UI被刷爆
            var lastReportAt = DateTime.MinValue;
            var progress = new Progress<ImportProgress>(p =>
            {
                var now = DateTime.Now;
                if ((now - lastReportAt).TotalMilliseconds < 150 && p.Percentage < 100) return;
                lastReportAt = now;
                SendMessageToWebView(new
                {
                    action = "importProgressUpdate",
                    stage = p.Stage,
                    percent = p.Percentage,
                    currentRow = p.CurrentRow,
                    totalRows = p.TotalRows
                });
            });

            var result = await _dataImporter.ImportWorksheetAsync(
                _currentFilePath,
                worksheetName,
                tableName: "Main",
                importMode: importMode,
                progress: progress,
                cancellationToken: CancellationToken.None);

            // 保存导入统计，供 SQLite 模式下的“文件分析/工作表解析”展示数据质量
            _lastConversionStats = result.ConversionStats ?? new List<ConversionStat>();

            SendMessageToWebView(new
            {
                action = "importComplete",
                result = new
                {
                    success = result.Success,
                    message = result.Message,
                    tableName = result.TableName,
                    rowCount = result.RowCount,
                    columnCount = result.ColumnCount,
                    importTime = result.ImportTime,
                    conversionStats = result.ConversionStats
                }
            });

            // 导入完成后刷新表列表（驱动SQL编辑器/其他功能）
            GetTableList();
        }
        catch (Exception ex)
        {
            WriteErrorLog("导入SQLite失败", ex);
            SendMessageToWebView(new { action = "error", message = $"导入失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private async void ExecuteQuery(System.Text.Json.JsonElement data)
    {
        try
        {
            string sql = data.GetProperty("sql").GetString() ?? string.Empty;
            string dbType = (data.TryGetProperty("dbType", out var dt) ? dt.GetString() : null) ?? "sqlite";
            string importMode =
                (data.TryGetProperty("importMode", out var im) ? im.GetString() : null)
                ?? "text";

            if (_queryEngine == null)
            {
                SendMessageToWebView(new { action = "error", message = "查询引擎未初始化" });
                return;
            }

            if (string.Equals(dbType, "excel", StringComparison.OrdinalIgnoreCase))
            {
                if (string.IsNullOrWhiteSpace(_currentFilePath))
                {
                    SendMessageToWebView(new { action = "error", message = "请先选择Excel文件" });
                    return;
                }
                if (_dataImporter == null)
                {
                    SendMessageToWebView(new { action = "error", message = "数据导入器未初始化" });
                    return;
                }

                // Excel SQL：按 SQL 语句提取表名（=工作表名），缺什么导什么
                await EnsureExcelTablesForSqlAsync(_currentFilePath, sql, importMode);
            }

            var result = await Task.Run(() => _queryEngine.ExecuteQuery(sql));

            SendMessageToWebView(new
            {
                action = "queryComplete",
                result = new
                {
                    columns = result.Columns,
                    rows = result.Rows,
                    totalRows = result.TotalRows,
                    queryTime = result.QueryTime,
                    sql = result.Sql
                }
            });
        }
        catch (Exception ex)
        {
            WriteErrorLog("执行SQL失败", ex);
            SendMessageToWebView(new { action = "error", message = $"查询执行失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private static List<string> ExtractTableNamesFromSql(string sql)
    {
        // 轻量解析：抓 FROM / JOIN / UPDATE / INTO 后的第一个 token
        // 支持 [表名] / "表名" / `表名` / 纯 token（中文/字母/数字/_/$）
        var list = new List<string>();
        if (string.IsNullOrWhiteSpace(sql)) return list;

        var rx = new System.Text.RegularExpressions.Regex(
            @"\b(from|join|update|into)\s+(\[[^\]]+\]|""[^""]+""|`[^`]+`|[A-Za-z0-9_\u4e00-\u9fa5\$]+)",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        foreach (System.Text.RegularExpressions.Match m in rx.Matches(sql))
        {
            var token = m.Groups[2].Value.Trim();
            if (token.StartsWith("[") && token.EndsWith("]")) token = token[1..^1].Replace("]]", "]");
            else if (token.StartsWith("\"") && token.EndsWith("\"")) token = token[1..^1].Replace("\"\"", "\"");
            else if (token.StartsWith("`") && token.EndsWith("`")) token = token[1..^1];

            if (!string.IsNullOrWhiteSpace(token) && !list.Contains(token, StringComparer.OrdinalIgnoreCase))
                list.Add(token);
        }

        return list;
    }

    private async Task EnsureExcelTablesForSqlAsync(string filePath, string sql, string importMode)
    {
        // 表名=工作表名：SQL中引用到的表，如果 SQLite 里没有，就从 Excel 导入一份同名表
        var tableNames = ExtractTableNamesFromSql(sql);
        if (tableNames.Count == 0) return;

        foreach (var table in tableNames)
        {
            if (_excelSqliteImportedTables.Contains(table)) continue;

            SendMessageToWebView(new { action = "importProgressUpdate", stage = $"Excel→SQLite 导入 {table}", percent = 0 });

            var progress = new Progress<ImportProgress>(p =>
            {
                SendMessageToWebView(new { action = "importProgressUpdate", stage = $"Excel→SQLite 导入 {table}", percent = p.Percentage });
            });

            var r = await _dataImporter!.ImportWorksheetAsync(
                filePath,
                worksheetName: table,
                tableName: table,           // 关键：表名=工作表名
                importMode: importMode,      // text/smart
                progress: progress,
                cancellationToken: CancellationToken.None);

            if (!r.Success)
                throw new InvalidOperationException(r.Message);

            _excelSqliteImportedTables.Add(table);
        }

        // 更新表列表，让前端（SQL编辑器/下拉）能看到新表
        GetTableList();
    }

    private void ExecuteSqlEditor(System.Text.Json.JsonElement data)
    {
        try
        {
            // 统一走 ExecuteQuery（支持 dbType=sqlite/excel）
            ExecuteQuery(data);
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"SQL执行失败: {ex.Message}" });
        }
    }

    private void GetSqliteTableFields(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }

            string tableName = (data.TryGetProperty("tableName", out var tn) ? tn.GetString() : null) ?? "Main";
            var schema = _sqliteManager.GetTableSchema(tableName);
            var fields = schema.Select(s => s.ColumnName).ToList();

            SendMessageToWebView(new
            {
                action = "worksheetFieldsLoaded",
                sheetName = tableName,
                fields = fields
            });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"获取SQLite字段失败: {ex.Message}" });
        }
    }

    private void GetActiveFields(System.Text.Json.JsonElement data)
    {
        try
        {
            string dbType = (data.TryGetProperty("dbType", out var dt) ? dt.GetString() : null) ?? "sqlite";

            if (string.Equals(dbType, "sqlite", StringComparison.OrdinalIgnoreCase))
            {
                if (_sqliteManager == null)
                {
                    SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                    return;
                }

                string tableName = (data.TryGetProperty("tableName", out var tn) ? tn.GetString() : null) ?? "Main";
                var schema = _sqliteManager.GetTableSchema(tableName);
                var fields = schema.Select(s => s.ColumnName).ToList();

                SendMessageToWebView(new
                {
                    action = "activeFieldsLoaded",
                    dbType = "sqlite",
                    tableName,
                    fields
                });
                return;
            }

            // excel
            if (_excelAnalyzer == null)
            {
                SendMessageToWebView(new { action = "error", message = "Excel分析器未初始化" });
                return;
            }

            string filePath =
                (data.TryGetProperty("filePath", out var fp) ? fp.GetString() : null)
                ?? _currentFilePath
                ?? string.Empty;

            string worksheetName =
                (data.TryGetProperty("worksheetName", out var wn) ? wn.GetString() : null)
                ?? _currentWorksheetName
                ?? string.Empty;

            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                SendMessageToWebView(new { action = "error", message = "请先选择Excel文件" });
                return;
            }
            if (string.IsNullOrWhiteSpace(worksheetName))
            {
                SendMessageToWebView(new { action = "error", message = "请先选择工作表" });
                return;
            }

            var fieldsExcel = _excelAnalyzer.GetWorksheetFields(filePath, worksheetName);
            SendMessageToWebView(new
            {
                action = "activeFieldsLoaded",
                dbType = "excel",
                filePath = filePath.Replace('\\', '/'),
                worksheetName,
                fields = fieldsExcel
            });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"加载字段失败: {ex.Message}" });
        }
    }

    private void ExecuteGlobalSearch(System.Text.Json.JsonElement data)
    {
        try
        {
            string keyword = data.GetProperty("keyword").GetString() ?? string.Empty;
            bool searchAllSheets = data.GetProperty("searchAllSheets").GetBoolean();
            bool searchFieldNames = data.GetProperty("searchFieldNames").GetBoolean();
            bool searchDataContent = data.GetProperty("searchDataContent").GetBoolean();
            bool searchExactMatch = data.GetProperty("searchExactMatch").GetBoolean();
            string searchOption = data.GetProperty("searchOption").GetString() ?? "contains";

            if (_queryEngine == null || _sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "查询引擎未初始化" });
                return;
            }

            var tables = _sqliteManager.GetTables();
            var results = _queryEngine.ExecuteGlobalSearch(
                keyword, tables, searchFieldNames, searchDataContent, searchExactMatch, searchOption);

            SendMessageToWebView(new
            {
                action = "globalSearchComplete",
                results = new
                {
                    keyword,
                    count = results.Count,
                    items = results
                }
            });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"全局搜索失败: {ex.Message}" });
        }
    }

    private void ExecuteSingleTableStats(System.Text.Json.JsonElement data)
    {
        try
        {
            string tableName = data.GetProperty("tableName").GetString() ?? string.Empty;
            var groupByFields = data.GetProperty("groupByFields").EnumerateArray()
                .Select(e => e.GetString() ?? string.Empty).ToList();
            var metrics = data.GetProperty("metrics").EnumerateArray()
                .Select(e => new StatisticsMetric
                {
                    FieldName = e.GetProperty("fieldName").GetString() ?? string.Empty,
                    AggregateFunction = e.GetProperty("function").GetString() ?? "COUNT",
                    Alias = e.GetProperty("alias").GetString() ?? string.Empty
                }).ToList();

            if (_statisticsEngine == null)
            {
                SendMessageToWebView(new { action = "error", message = "统计引擎未初始化" });
                return;
            }

            var result = _statisticsEngine.ExecuteSingleTableStatistics(tableName, groupByFields, metrics);

            SendMessageToWebView(new
            {
                action = "singleTableStatsComplete",
                results = new
                {
                    groupBy = string.Join(", ", groupByFields),
                    metrics = string.Join(", ", metrics.Select(m => $"{m.FieldName} ({m.AggregateFunction})")),
                    totalRows = result.TotalRows,
                    statsTime = result.StatisticsTime,
                    columns = result.Columns,
                    data = result.Rows,
                    sql = result.Sql
                }
            });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"单表统计失败: {ex.Message}" });
        }
    }

    private void ExecuteMultiTableStats(System.Text.Json.JsonElement data)
    {
        try
        {
            string mainTable = data.GetProperty("mainTable").GetString() ?? string.Empty;
            string joinTable = data.GetProperty("joinTable").GetString() ?? string.Empty;
            string mainField = data.GetProperty("mainField").GetString() ?? string.Empty;
            string joinField = data.GetProperty("joinField").GetString() ?? string.Empty;
            var groupByFields = data.GetProperty("groupByFields").EnumerateArray()
                .Select(e => e.GetString() ?? string.Empty).ToList();
            var metrics = data.GetProperty("metrics").EnumerateArray()
                .Select(e => new StatisticsMetric
                {
                    FieldName = e.GetProperty("fieldName").GetString() ?? string.Empty,
                    AggregateFunction = e.GetProperty("function").GetString() ?? "COUNT",
                    Alias = e.GetProperty("alias").GetString() ?? string.Empty
                }).ToList();

            if (_statisticsEngine == null)
            {
                SendMessageToWebView(new { action = "error", message = "统计引擎未初始化" });
                return;
            }

            var result = _statisticsEngine.ExecuteMultiTableStatistics(
                mainTable, joinTable, mainField, joinField, groupByFields, metrics);

            SendMessageToWebView(new
            {
                action = "multiTableStatsComplete",
                results = new
                {
                    joinCondition = $"{mainTable}.{mainField} = {joinTable}.{joinField}",
                    groupBy = string.Join(", ", groupByFields),
                    metrics = string.Join(", ", metrics.Select(m => $"{m.FieldName} ({m.AggregateFunction})")),
                    totalRows = result.TotalRows,
                    statsTime = result.StatisticsTime,
                    columns = result.Columns,
                    data = result.Rows,
                    sql = result.Sql
                }
            });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"多表统计失败: {ex.Message}" });
        }
    }

    private void ExecuteSplit(System.Text.Json.JsonElement data)
    {
        try
        {
            string sourceTable = data.GetProperty("sourceTable").GetString() ?? string.Empty;
            string splitField = data.GetProperty("splitField").GetString() ?? string.Empty;
            string outputDirectory = data.GetProperty("outputDirectory").GetString() ?? string.Empty;
            string splitType = data.GetProperty("splitType").GetString() ?? "byField";

            if (_splitEngine == null)
            {
                SendMessageToWebView(new { action = "error", message = "分拆引擎未初始化" });
                return;
            }

            SplitResult result;
            if (splitType == "byField")
            {
                result = _splitEngine.SplitByField(sourceTable, splitField, outputDirectory);
            }
            else
            {
                int rowsPerFile = data.GetProperty("rowsPerFile").GetInt32();
                result = _splitEngine.SplitByRowCount(sourceTable, rowsPerFile, outputDirectory);
            }

            SendMessageToWebView(new
            {
                action = "splitComplete",
                result = new
                {
                    fileCount = result.Files.Count,
                    totalRows = result.TotalRows,
                    splitTime = result.SplitTime,
                    files = result.Files.Select(f => new
                    {
                        f.FileName,
                        f.RowCount,
                        f.SplitValue
                    })
                }
            });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"数据分拆失败: {ex.Message}" });
        }
    }

    private void GetWorksheetFields(System.Text.Json.JsonElement data)
    {
        try
        {
            // 兼容不同前端字段命名：worksheetName / sheetName
            string worksheetName =
                (data.TryGetProperty("worksheetName", out var wn) ? wn.GetString() : null)
                ?? (data.TryGetProperty("sheetName", out var sn) ? sn.GetString() : null)
                ?? string.Empty;

            // 兼容：filePath可不传（使用当前已选文件）
            string filePath =
                (data.TryGetProperty("filePath", out var fp) ? fp.GetString() : null)
                ?? _currentFilePath
                ?? string.Empty;

            if (_excelAnalyzer == null)
            {
                SendMessageToWebView(new { action = "error", message = "Excel分析器未初始化" });
                return;
            }

            if (string.IsNullOrWhiteSpace(filePath))
            {
                SendMessageToWebView(new { action = "error", message = "请先选择文件" });
                return;
            }
            if (string.IsNullOrWhiteSpace(worksheetName))
            {
                SendMessageToWebView(new { action = "error", message = "请先选择工作表" });
                return;
            }

            var fields = _excelAnalyzer.GetWorksheetFields(filePath, worksheetName);
            SendMessageToWebView(new { action = "worksheetFieldsLoaded", sheetName = worksheetName, fields });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"获取字段列表失败: {ex.Message}" });
        }
    }

    private void GetTableList()
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }

            var tables = _sqliteManager.GetTables();
            SendMessageToWebView(new { action = "tableListLoaded", tables });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"获取表列表失败: {ex.Message}" });
        }
    }

    /// <summary>
    /// 清空当前 SQLite 内存库（删除所有业务表；保留 sqlite_master）
    /// </summary>
    private void ResetSqliteDatabase()
    {
        if (_sqliteManager == null) return;
        try
        {
            var tables = _sqliteManager.GetTables();
            foreach (var t in tables)
            {
                if (string.IsNullOrWhiteSpace(t)) continue;
                _sqliteManager.Execute($"DROP TABLE IF EXISTS [{t}];");
            }
        }
        catch (Exception ex)
        {
            // 不阻塞导入：仅记录
            System.Diagnostics.Debug.WriteLine("ResetSqliteDatabase failed: " + ex.Message);
        }
    }

    private void WriteErrorLog(string title, Exception ex)
    {
        try
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "ExcelSQLiteWeb",
                "logs");
            Directory.CreateDirectory(dir);

            var file = $"error-{DateTime.Now:yyyyMMdd-HHmmss}.log";
            var path = Path.Combine(dir, file);

            var sb = new StringBuilder();
            sb.AppendLine("ExcelSQLite 错误记录");
            sb.AppendLine($"时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"标题: {title}");
            sb.AppendLine($"当前文件: {_currentFilePath}");
            sb.AppendLine($"当前工作表: {_currentWorksheetName}");
            sb.AppendLine(new string('-', 60));
            sb.AppendLine(ex.ToString());

            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
            _lastErrorLogPath = path;

            SendMessageToWebView(new { action = "errorLogAvailable" });
        }
        catch (Exception e)
        {
            System.Diagnostics.Debug.WriteLine("WriteErrorLog failed: " + e.Message);
        }
    }

    private void DownloadErrorLog()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(_lastErrorLogPath) || !File.Exists(_lastErrorLogPath))
            {
                SendMessageToWebView(new { action = "error", message = "暂无可下载的错误记录", hasErrorLog = false });
                return;
            }

            using var sfd = new SaveFileDialog();
            sfd.Title = "保存错误记录";
            sfd.Filter = "日志文件 (*.log)|*.log|文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*";
            sfd.FileName = Path.GetFileName(_lastErrorLogPath);

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                File.Copy(_lastErrorLogPath, sfd.FileName, overwrite: true);
            }
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"下载错误记录失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void ExportQueryToFile(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null || _queryEngine == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite/查询引擎未初始化" });
                return;
            }

            string sql = (data.TryGetProperty("sql", out var s) ? s.GetString() : null) ?? string.Empty;
            string format = (data.TryGetProperty("format", out var f) ? f.GetString() : null) ?? "xlsx";
            int limit = (data.TryGetProperty("limit", out var l) && l.ValueKind == System.Text.Json.JsonValueKind.Number)
                ? l.GetInt32()
                : 0;

            if (string.IsNullOrWhiteSpace(sql))
            {
                SendMessageToWebView(new { action = "error", message = "导出失败：SQL为空" });
                return;
            }

            // 大数量提示：尽量计算行数（对大表 count(*) 也可能耗时，但比导出更快）
            long cnt = 0;
            bool forceAllText = string.Equals(format, "xlsx_text", StringComparison.OrdinalIgnoreCase);

            try
            {
                var sqlNoSemi = sql.Trim().TrimEnd(';');
                var countSql = $"SELECT COUNT(*) AS cnt FROM ({sqlNoSemi}) t";
                var cntRows = _sqliteManager.Query(countSql);
                if (cntRows.Count > 0 && cntRows[0].TryGetValue("cnt", out var v) && v != null)
                    cnt = Convert.ToInt64(v);

                if ((string.Equals(format, "xlsx", StringComparison.OrdinalIgnoreCase) || forceAllText) && cnt > 1048576)
                {
                    MessageBox.Show(
                        $"查询结果预计 {cnt:n0} 行，已超过 Excel 单工作表最大行数 1,048,576。\n请改用 CSV 导出或增加过滤条件。",
                        "无法导出为 Excel",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                if (limit <= 0 && cnt > 100000)
                {
                    var dr = MessageBox.Show(
                        $"查询结果预计 {cnt:n0} 行，导出可能耗时较长并占用较大磁盘空间。\n\n是否继续导出全部？\n（建议：先把预览行数调大做抽查；或增加 WHERE 过滤）",
                        "导出提示",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);
                    if (dr != DialogResult.Yes) return;
                }
            }
            catch
            {
                // count 失败不阻塞导出
            }

            using var sfd = new SaveFileDialog();
            if (string.Equals(format, "csv", StringComparison.OrdinalIgnoreCase))
            {
                sfd.Title = "导出查询结果（CSV）";
                sfd.Filter = "CSV 文件 (*.csv)|*.csv|所有文件 (*.*)|*.*";
                sfd.FileName = $"query-{DateTime.Now:yyyyMMdd-HHmmss}.csv";
            }
            else
            {
                sfd.Title = forceAllText ? "导出查询结果（Excel .xlsx - 全列文本）" : "导出查询结果（Excel .xlsx）";
                sfd.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
                sfd.FileName = $"query-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx";
            }
            if (sfd.ShowDialog() != DialogResult.OK) return;

            // 执行导出：为避免一次性吃内存，这里用 SqliteDataReader 流式写 CSV
            var exportSql = sql.Trim().TrimEnd(';');
            if (limit > 0 && !exportSql.Contains(" limit ", StringComparison.OrdinalIgnoreCase) && exportSql.TrimStart().StartsWith("select", StringComparison.OrdinalIgnoreCase))
            {
                exportSql = $"SELECT * FROM ({exportSql}) t LIMIT {limit}";
            }

            if (string.Equals(format, "csv", StringComparison.OrdinalIgnoreCase))
            {
                ExportSqlToCsv(exportSql, sfd.FileName, protectForExcel: true);
            }
            else
            {
                // 说明：需要安装 ClosedXML NuGet：ClosedXML
                ExportSqlToXlsx(exportSql, sfd.FileName, forceAllText);
            }
        }
        catch (Exception ex)
        {
            WriteErrorLog("导出查询结果失败", ex);
            SendMessageToWebView(new { action = "error", message = $"导出失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void ExportSqlToCsv(string sql, string outputPath, bool protectForExcel)
    {
        if (_sqliteManager == null) throw new InvalidOperationException("SQLite管理器未初始化");

        _sqliteManager.Open();
        using var cmd = _sqliteManager.Connection!.CreateCommand();
        cmd.CommandText = sql;

        using var reader = cmd.ExecuteReader();
        using var writer = new StreamWriter(outputPath, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));

        // header
        var colCount = reader.FieldCount;
        for (int i = 0; i < colCount; i++)
        {
            if (i > 0) writer.Write(",");
            writer.Write(CsvEscape(reader.GetName(i)));
        }
        writer.WriteLine();

        while (reader.Read())
        {
            for (int i = 0; i < colCount; i++)
            {
                if (i > 0) writer.Write(",");
                var v = reader.IsDBNull(i) ? "" : Convert.ToString(reader.GetValue(i));
                var cell = v ?? "";
                if (protectForExcel && LooksLikeSensitiveNumber(cell))
                {
                    // 防止 Excel 科学计数法/前导0丢失：
                    // 方案：输出为 ="原值"（Excel 打开会按文本显示；其它 CSV 解析器也能读到字符串）
                    cell = "=\"" + cell.Replace("\"", "\"\"") + "\"";
                }
                writer.Write(CsvEscape(cell));
            }
            writer.WriteLine();
        }
    }

    private static bool LooksLikeSensitiveNumber(string s)
    {
        var t = (s ?? "").Trim();
        if (t.Length < 8) return false;
        // 纯数字且较长、或以0开头的编码类字段
        if (t.All(char.IsDigit))
        {
            if (t.StartsWith("0")) return true;
            if (t.Length >= 12) return true;     // 银行卡/长编码
            if (t.Length >= 15) return true;     // Excel 15位精度风险
        }
        return false;
    }

    private void ExportSqlToXlsx(string sql, string outputPath, bool forceAllText)
    {
        // 需要 NuGet：ClosedXML
        // using ClosedXML.Excel;
        if (_sqliteManager == null) throw new InvalidOperationException("SQLite管理器未初始化");

        _sqliteManager.Open();
        using var cmd = _sqliteManager.Connection!.CreateCommand();
        cmd.CommandText = sql;

        using var reader = cmd.ExecuteReader();
        var colCount = reader.FieldCount;

        var wb = new ClosedXML.Excel.XLWorkbook();
        var ws = wb.Worksheets.Add("Query");

        // header
        for (int c = 0; c < colCount; c++)
        {
            ws.Cell(1, c + 1).Value = reader.GetName(c);
            ws.Cell(1, c + 1).Style.Font.Bold = true;
        }

        int row = 2;
        while (reader.Read())
        {
            for (int c = 0; c < colCount; c++)
            {
                var v = reader.IsDBNull(c) ? null : reader.GetValue(c);
                if (v == null)
                {
                    ws.Cell(row, c + 1).Value = "";
                    continue;
                }

                var s = Convert.ToString(v);
                if (forceAllText)
                {
                    // 全列文本：任何值都按文本写，避免科学计数法/前导0/精度截断
                    ws.Cell(row, c + 1).SetValue(s ?? "");
                    ws.Cell(row, c + 1).Style.NumberFormat.Format = "@";
                }
                else if (!string.IsNullOrEmpty(s) && LooksLikeSensitiveNumber(s))
                {
                    // 强制文本格式，避免科学计数法/前导0丢失
                    ws.Cell(row, c + 1).SetValue(s);
                    ws.Cell(row, c + 1).Style.NumberFormat.Format = "@";
                }
                else
                {
                    // 尽量保留数值类型（统计/排序更友好）
                    if (double.TryParse(s, out var d) && s!.Length < 15)
                        ws.Cell(row, c + 1).Value = d;
                    else
                        ws.Cell(row, c + 1).SetValue(s);
                }
            }
            row++;
        }

        ws.Columns().AdjustToContents();
        wb.SaveAs(outputPath);
    }

    private static string CsvEscape(string s)
    {
        if (s.Contains('"') || s.Contains(',') || s.Contains('\n') || s.Contains('\r'))
        {
            return "\"" + s.Replace("\"", "\"\"") + "\"";
        }
        return s;
    }

    private void OpenRecentFile(string fileName)
    {
        try
        {
            System.Diagnostics.Debug.WriteLine($"Opening recent file: {fileName}");
            var record = _recentFiles.FirstOrDefault(x =>
                string.Equals(Path.GetFileName(x.FullPath), fileName, StringComparison.OrdinalIgnoreCase));

            if (record == null || string.IsNullOrWhiteSpace(record.FullPath))
            {
                SendMessageToWebView(new { action = "error", message = $"最近文件未找到: {fileName}" });
                return;
            }

            if (!File.Exists(record.FullPath))
            {
                SendMessageToWebView(new { action = "error", message = $"文件不存在: {record.FullPath}" });
                return;
            }

            _currentFilePath = record.FullPath;
            _excelSqliteImportedTables.Clear();
            TouchRecentFile(record.FullPath);
            NotifyRecentFilesUpdated();

            // 将最近文件作为“当前文件数据源”打开
            SendMessageToWebView(new { action = "fileSelected", filePath = record.FullPath.Replace('\\', '/') });
            LoadWorksheetList(record.FullPath);
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"打开最近文件失败: {ex.Message}" });
        }
    }

    private void ApplyTemplate(string templateName)
    {
        System.Diagnostics.Debug.WriteLine($"Applying template: {templateName}");
    }

    private void OpenTool(string toolName)
    {
        System.Diagnostics.Debug.WriteLine($"Opening tool: {toolName}");
    }

    private void ShowHelp()
    {
        MessageBox.Show("帮助功能正在开发中...", "帮助");
    }

    private void ShowAbout()
    {
        MessageBox.Show("ExcelSQLite - 数据处理工具\n版本 1.0.0\n© 2024 ExcelSQLite", "关于");
    }

    private void ShowLearningCases()
    {
        string casesText = "📚 ExcelSQLite 学习案例\n\n" +
                          "1. 销售数据统计\n" +
                          "   - 按地区统计销售额\n" +
                          "   - 按产品类别统计销量\n" +
                          "   - 按时间维度分析趋势\n\n" +
                          "2. 客户数据分析\n" +
                          "   - 客户画像标签\n" +
                          "   - 购买行为分析\n" +
                          "   - 客户分群统计\n\n" +
                          "3. 库存管理\n" +
                          "   - 库存周转率计算\n" +
                          "   - 安全库存预警\n" +
                          "   - 库存ABC分类";
        MessageBox.Show(casesText, "学习案例");
    }

    private void SendMessageToWebView(object message)
    {
        try
        {
            string json = System.Text.Json.JsonSerializer.Serialize(message);
            webView21.CoreWebView2?.PostWebMessageAsJson(json);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error sending message: {ex.Message}");
        }
    }

    private string FormatFileSize(long bytes)
    {
        if (bytes < 1024) return $"{bytes} B";
        if (bytes < 1024 * 1024) return $"{bytes / 1024.0:F1} KB";
        if (bytes < 1024 * 1024 * 1024) return $"{bytes / (1024.0 * 1024.0):F1} MB";
        return $"{bytes / (1024.0 * 1024.0 * 1024.0):F1} GB";
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        _sqliteManager?.Dispose();
        base.OnFormClosing(e);
    }
}
