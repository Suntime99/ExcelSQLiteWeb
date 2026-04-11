using ExcelSQLiteWeb.Models;
using ExcelSQLiteWeb.Services;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeOpenXml;
using Microsoft.Data.Sqlite;
using Dapper;
using System.Drawing.Imaging;

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
    private string? _lastCleansedFilePath;
    private string? _lastCleansingReportPath;
    // Excel 模式 SQL：已导入到 SQLite 的工作表（表名=工作表名）
    private readonly HashSet<string> _excelSqliteImportedTables = new(StringComparer.OrdinalIgnoreCase);

    // SQL实验室：专家模式下的“未提交事务”（用于提交/回滚）
    private SqliteTransaction? _sqlLabTxn;
    // SQL实验室：当前执行的取消令牌（用于取消执行/超时）
    private CancellationTokenSource? _sqlExecCts;
    // 图表：最近一次生成的图表 PNG（用于导出）
    private string? _lastChartPngPath;

    // 最近一次错误日志（用于“错误记录下载”）
    private string? _lastErrorLogPath;

    // 最近文件（持久化到本机设置文件）
    private const int MaxRecentFiles = 10;
    private readonly List<RecentFileRecord> _recentFiles = new();
    // 最近一次打开的“主表文件”（用于启动时自动回填）
    private string? _lastMainFilePath;
    // 当前“主表”在 SQLite 内存库中的真实表名（默认兼容 Main）
    private string? _currentMainTableName;

    // ==================== 方案管理（DataConfig，位于主程序目录） ====================
    // 说明：
    // - 方案文件：<主程序目录>/DataConfig/schemes/{方案名}.ini
    // - 临时库：  <主程序目录>/DataConfig/db/{方案名}.db
    // - 最近列表：<主程序目录>/DataConfig/recent.json
    private const int DefaultMaxRecentSchemesDisplay = 10; // UI 显示数量
    private const int DefaultMaxRecentSchemesKeep = 50;    // 最大保留数量
    private int _maxSchemesDisplay = DefaultMaxRecentSchemesDisplay;
    private int _maxSchemesKeep = DefaultMaxRecentSchemesKeep;

    private readonly List<SchemeMeta> _schemes = new();
    private string? _activeSchemeId;     // 方案“文件名”基名（安全化）
    private string? _activeSchemeDbPath; // 绝对路径：DataConfig/db/{id}.db
    private string _dataConfigDir = "";  // 实际 DataConfig 根目录（优先主程序目录，不可写则回退 AppData）

    private string MainTableNameOrDefault()
        => string.IsNullOrWhiteSpace(_currentMainTableName) ? "Main" : _currentMainTableName!;

    public Form1()
    {
        InitializeComponent();
        InitializeDragDropSupport();
        _dataConfigDir = ResolveDataConfigDir();
        InitializeServices();
        EnsureDataConfigDirs();
        LoadRecentFilesFromDisk();
        LoadRecentSchemesFromDisk();
        _lastMainFilePath = _recentFiles.FirstOrDefault()?.FullPath;
        InitializeWebView2();
    }

    private string ResolveDataConfigDir()
    {
        // 优先：主程序目录下的 DataConfig（便于测试/调试/打包）
        // 回退：%AppData%/ExcelSQLiteWeb/DataConfig（当主程序目录不可写时）
        string primary = Path.Combine(AppContext.BaseDirectory, "DataConfig");
        if (TryEnsureWritableDir(primary)) return primary;

        string fallback = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelSQLiteWeb",
            "DataConfig");
        TryEnsureWritableDir(fallback);
        return fallback;
    }

    private static bool TryEnsureWritableDir(string dir)
    {
        try
        {
            Directory.CreateDirectory(dir);
            // 简单写入测试，判断是否可写
            var test = Path.Combine(dir, ".write_test.tmp");
            File.WriteAllText(test, DateTime.Now.ToString("O"), Encoding.UTF8);
            File.Delete(test);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private void InitializeDragDropSupport()
    {
        try
        {
            // 允许把 Excel 文件直接拖放到界面（WebView2 控件上）
            // 注意：WebView2 自身的 AllowDrop 属性是只读（会隐藏 Control.AllowDrop）。
            // 这里强制使用基类 Control.AllowDrop 以启用 WinForms 拖放事件。
            ((Control)webView21).AllowDrop = true;
            webView21.DragEnter += (_, e) =>
            {
                try
                {
                    if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
                        e.Effect = DragDropEffects.Copy;
                    else
                        e.Effect = DragDropEffects.None;
                }
                catch { e.Effect = DragDropEffects.None; }
            };
            webView21.DragDrop += (_, e) =>
            {
                try
                {
                    if (e.Data == null || !e.Data.GetDataPresent(DataFormats.FileDrop)) return;
                    var files = (string[]?)e.Data.GetData(DataFormats.FileDrop);
                    if (files == null || files.Length == 0) return;
                    var excel = files
                        .Where(f => !string.IsNullOrWhiteSpace(f))
                        .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .Take(20)
                        .Select(f => f.Replace('\\', '/'))
                        .ToArray();
                    if (excel.Length == 0) return;
                    SendMessageToWebView(new { action = "hostFilesDropped", files = excel });
                }
                catch { }
            };
        }
        catch { }
    }

    private sealed class RecentSchemesStore
    {
        public string? LastSchemeId { get; set; }
        public List<SchemeMeta> Schemes { get; set; } = new();
    }

    private sealed class RecentFileRecord
    {
        public string FullPath { get; set; } = string.Empty;
        public DateTime LastOpenTime { get; set; } = DateTime.Now;
    }

    private sealed class SchemeMeta
    {
        public string Id { get; set; } = "";
        public string Name { get; set; } = "未命名方案";
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

    private string GetDataConfigDir()
        => string.IsNullOrWhiteSpace(_dataConfigDir) ? Path.Combine(AppContext.BaseDirectory, "DataConfig") : _dataConfigDir;

    private string GetSchemesDir()
        => Path.Combine(GetDataConfigDir(), "schemes");

    private string GetDbDir()
        => Path.Combine(GetDataConfigDir(), "db");

    private string GetRecentSchemesStorePath()
        => Path.Combine(GetDataConfigDir(), "recent.json");

    private void EnsureDataConfigDirs()
    {
        Directory.CreateDirectory(GetDataConfigDir());
        Directory.CreateDirectory(GetSchemesDir());
        Directory.CreateDirectory(GetDbDir());
    }

    private string GetSchemeIniPath(string schemeId)
        => Path.Combine(GetSchemesDir(), $"{schemeId}.ini");

    private string GetSchemeDbPath(string schemeId)
        => Path.Combine(GetDbDir(), $"{schemeId}.db");

    private static string SanitizeFileName(string name, int maxLen = 80)
    {
        var s = (name ?? "").Trim();
        if (string.IsNullOrWhiteSpace(s)) s = "未命名方案";
        // Windows 文件名非法字符：\/:*?"<>|
        foreach (var ch in Path.GetInvalidFileNameChars())
            s = s.Replace(ch, '_');
        s = s.Replace('/', '_').Replace('\\', '_').Replace(':', '_').Replace('*', '_').Replace('?', '_').Replace('"', '_')
             .Replace('<', '_').Replace('>', '_').Replace('|', '_');
        s = s.Trim().TrimEnd('.'); // Windows 不允许尾部 '.'
        if (s.Length > maxLen) s = s.Substring(0, maxLen).Trim();
        if (string.IsNullOrWhiteSpace(s)) s = "未命名方案";
        // 处理保留名
        var upper = s.ToUpperInvariant();
        var reserved = new HashSet<string> { "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9" };
        if (reserved.Contains(upper)) s = "_" + s;
        return s;
    }

    private static string ToB64(string s)
        => Convert.ToBase64String(Encoding.UTF8.GetBytes(s ?? string.Empty));

    private static string FromB64(string b64)
    {
        try
        {
            var bytes = Convert.FromBase64String(b64 ?? string.Empty);
            return Encoding.UTF8.GetString(bytes);
        }
        catch { return string.Empty; }
    }

    private static Dictionary<string, Dictionary<string, string>> ReadIni(string path)
    {
        var dict = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
        string current = "General";
        dict[current] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (!File.Exists(path)) return dict;

        foreach (var raw in File.ReadAllLines(path, Encoding.UTF8))
        {
            var line = raw.Trim();
            if (string.IsNullOrWhiteSpace(line)) continue;
            if (line.StartsWith(";") || line.StartsWith("#")) continue;
            if (line.StartsWith("[") && line.EndsWith("]"))
            {
                current = line.Substring(1, line.Length - 2).Trim();
                if (!dict.ContainsKey(current))
                    dict[current] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                continue;
            }
            var idx = line.IndexOf('=');
            if (idx <= 0) continue;
            var key = line.Substring(0, idx).Trim();
            var val = line.Substring(idx + 1).Trim();
            if (!dict.ContainsKey(current))
                dict[current] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            dict[current][key] = val;
        }
        return dict;
    }

    private static void WriteIni(string path, Dictionary<string, Dictionary<string, string>> data)
    {
        var sb = new StringBuilder();
        foreach (var sec in data.Keys)
        {
            sb.AppendLine($"[{sec}]");
            foreach (var kv in data[sec])
            {
                sb.AppendLine($"{kv.Key}={kv.Value}");
            }
            sb.AppendLine();
        }
        File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
    }

    private void LoadRecentSchemesFromDisk()
    {
        try
        {
            var path = GetRecentSchemesStorePath();
            _schemes.Clear();
            if (!File.Exists(path)) return;
            var json = File.ReadAllText(path, Encoding.UTF8);
            var store = System.Text.Json.JsonSerializer.Deserialize<RecentSchemesStore>(json) ?? new RecentSchemesStore();
            _activeSchemeId = string.IsNullOrWhiteSpace(store.LastSchemeId) ? null : store.LastSchemeId;
            if (store.Schemes != null)
                _schemes.AddRange(store.Schemes.Where(s => !string.IsNullOrWhiteSpace(s.Id)));

            // 清理不存在的 ini（防止 recent.json 里残留）
            _schemes.RemoveAll(s => !File.Exists(GetSchemeIniPath(s.Id)));

            _schemes.Sort((a, b) => b.LastOpenTime.CompareTo(a.LastOpenTime));
            if (_schemes.Count > _maxSchemesKeep)
                _schemes.RemoveRange(_maxSchemesKeep, _schemes.Count - _maxSchemesKeep);

            // 若未记录 LastSchemeId，则默认取最近一个方案（提升“第二次打开自动加载”成功率）
            if (string.IsNullOrWhiteSpace(_activeSchemeId) && _schemes.Count > 0)
            {
                _activeSchemeId = _schemes[0].Id;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine("LoadRecentSchemesFromDisk failed: " + ex.Message);
        }
    }

    private void SaveRecentSchemesToDisk()
    {
        try
        {
            EnsureDataConfigDirs();
            var store = new RecentSchemesStore
            {
                LastSchemeId = _activeSchemeId,
                Schemes = _schemes
                    .OrderByDescending(s => s.LastOpenTime)
                    .Take(_maxSchemesKeep)
                    .ToList()
            };
            var json = System.Text.Json.JsonSerializer.Serialize(store, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(GetRecentSchemesStorePath(), json, Encoding.UTF8);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine("SaveRecentSchemesToDisk failed: " + ex.Message);
        }
    }

    private object BuildRecentSchemesPayload()
    {
        return _schemes
            .OrderByDescending(x => x.LastOpenTime)
            .Take(_maxSchemesDisplay)
            .Select(x => new
            {
                id = x.Id,
                name = x.Name,
                time = x.LastOpenTime.ToString("yyyy-MM-dd HH:mm")
            })
            .ToArray();
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

    private static string BuildSqliteTableName(string excelPath, string worksheetName)
    {
        var baseName = Path.GetFileNameWithoutExtension(excelPath ?? string.Empty);
        // 避免极端字符影响 SQL 标识符（前端/后端均用 [] 引用）
        baseName = baseName.Replace("]", "）").Replace("[", "（");
        worksheetName = (worksheetName ?? string.Empty).Replace("]", "）").Replace("[", "（");
        return $"{baseName}|{worksheetName}";
    }

    private void OpenMainFileInternal(string fullPath, bool clearPrevious)
    {
        if (string.IsNullOrWhiteSpace(fullPath) || !File.Exists(fullPath))
        {
            SendMessageToWebView(new { action = "error", message = $"文件不存在: {fullPath}" });
            return;
        }

        // 按要求：重新引入前清空先前内存库与加载状态
        if (clearPrevious)
        {
            ResetSqliteDatabase();
            _excelSqliteImportedTables.Clear();
            _lastConversionStats = new();
            _currentMainTableName = null;
        }

        _currentFilePath = fullPath;
        _excelSqliteImportedTables.Clear();
        TouchRecentFile(fullPath);
        NotifyRecentFilesUpdated();

        // 以“主表选择”方式回填到页面（而不是 fileSelected）
        SendMessageToWebView(new
        {
            action = "mainTableFileSelected",
            filePath = fullPath.Replace('\\', '/'),
            clearPrevious = clearPrevious
        });
        LoadWorksheetList(fullPath);
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
                    recentSchemes = BuildRecentSchemesPayload(),
                    recentFiles = BuildRecentFilesPayload(), // 兼容：保留旧字段
                    dataConfigDir = GetDataConfigDir(),
                    activeSchemeId = _activeSchemeId,
                    activeSchemeDbPath = _activeSchemeDbPath
                });
            };

            // 优先从磁盘读取（便于“就近覆盖”快速迭代）；若不存在再回退到嵌入式资源
            string htmlContent = "";
            // 1) 优先尝试读取 webView21.Source 指向的 file:// 路径（调试/开发场景常用）
            string htmlPath = "";
            try
            {
                var src = webView21.Source;
                if (src != null && src.IsFile)
                {
                    var p = src.LocalPath;
                    if (!string.IsNullOrWhiteSpace(p) && File.Exists(p)) htmlPath = p;
                }
            }
            catch { }

            // 2) 其次：exe 输出目录（发布时可直接覆盖文件）
            if (string.IsNullOrWhiteSpace(htmlPath))
                htmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "index.html");
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
                        string tableName = (data.TryGetProperty("tableName", out var tn) ? tn.GetString() : null) ?? MainTableNameOrDefault();
                        StartAnalysisSqlite(tableName);
                        break;
                    }
                case "startAnalysisSqliteAll":
                    {
                        StartAnalysisSqliteAll();
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
                        string tableName = (data.TryGetProperty("tableName", out var tn) ? tn.GetString() : null) ?? MainTableNameOrDefault();
                        AnalyzeSqliteWorksheet(tableName);
                        break;
                    }
                case "startMetadataScan":
                    StartMetadataScan(data);
                    break;
                case "detectRelations":
                    DetectRelations(data);
                    break;
                case "exportRelationReport":
                    ExportRelationReport(data);
                    break;
                case "exportRelationEnums":
                    ExportRelationEnums(data);
                    break;
                case "getRelationEnums":
                    GetRelationEnums(data);
                    break;
                case "exportRelationEnumPreview":
                    ExportRelationEnumPreview(data);
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
                case "executeMultiTableSplit":
                    ExecuteMultiTableSplit(data);
                    break;
                case "executeDataCleansing":
                    ExecuteDataCleansing(data);
                    break;
                case "executeBusinessRuleVerify":
                    ExecuteBusinessRuleVerify(data);
                    break;
                case "generateDataCleansingSql":
                    GenerateDataCleansingSql(data);
                    break;
                case "getDcPreviewStats":
                    GetDcPreviewStats(data);
                    break;
                case "browseDcOutputPath":
                    BrowseDcOutputPath();
                    break;
                case "executeDataCompare":
                    ExecuteDataCompare(data);
                    break;
                case "importWorksheet":
                    ImportWorksheet(data);
                    break;
                case "openRecentFile":
                    string recentFileName = data.GetProperty("fileName").GetString() ?? string.Empty;
                    OpenRecentFile(recentFileName);
                    break;
                case "openRecentScheme":
                    {
                        string schemeId = (data.TryGetProperty("schemeId", out var sid) ? sid.GetString() : null) ?? string.Empty;
                        if (string.IsNullOrWhiteSpace(schemeId))
                        {
                            SendMessageToWebView(new { action = "error", message = "方案ID为空" });
                            break;
                        }
                        LoadSchemeInternal(schemeId, autoLoad: false);
                        break;
                    }
                case "saveScheme":
                    SaveScheme(data);
                    break;
                case "cleanupTempDb":
                    CleanupTempDb();
                    break;
                case "rebuildSchemeDb":
                    RebuildSchemeDb();
                    break;
                case "setSubTableFile":
                    {
                        string subId = (data.TryGetProperty("subTableId", out var sid) ? sid.GetString() : null) ?? "";
                        string droppedPath = (data.TryGetProperty("filePath", out var fp) ? fp.GetString() : null) ?? "";
                        if (!string.IsNullOrWhiteSpace(droppedPath)) droppedPath = droppedPath.Replace('/', '\\');
                        if (string.IsNullOrWhiteSpace(subId) || string.IsNullOrWhiteSpace(droppedPath))
                        {
                            SendMessageToWebView(new { action = "error", message = "拖放导入失败：参数缺失" });
                            break;
                        }
                        if (!File.Exists(droppedPath))
                        {
                            SendMessageToWebView(new { action = "error", message = $"文件不存在：{droppedPath}" });
                            break;
                        }
                        TouchRecentFile(droppedPath);
                        NotifyRecentFilesUpdated();
                        SendMessageToWebView(new { action = "subTableFileSelected", subTableId = subId, filePath = droppedPath.Replace('\\', '/') });
                        try { LoadWorksheetList(droppedPath); } catch { }
                        break;
                    }
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
                case "setMainTable":
                    SetMainTable(data);
                    break;
                case "saveSettings":
                    SaveSettings(data);
                    break;
                case "loadSettings":
                    LoadSettings(data);
                    break;
                case "executeSqlEditor":
                    ExecuteSqlEditor(data);
                    break;
                case "sqlLabCommit":
                    SqlLabCommit();
                    break;
                case "sqlLabRollback":
                    SqlLabRollback();
                    break;
                case "cancelSqlExecution":
                    CancelSqlExecution();
                    break;
                case "executeMultiTableQuery":
                    System.Diagnostics.Debug.WriteLine("Executing multi-table query");
                    ExecuteQuery(data);
                    break;
                case "exportSqlResult":
                    System.Diagnostics.Debug.WriteLine("Exporting SQL result");
                    // 兼容旧前端：转发到 exportQueryToFile
                    ExportQueryToFile(data);
                    break;
                case "exportQueryToFile":
                    ExportQueryToFile(data);
                    break;
                case "exportGridToFile":
                    ExportGridToFile(data);
                    break;
                case "exportCleansingTemplate":
                    ExportCleansingTemplate(data);
                    break;
                case "exportCompareReport":
                    ExportCompareReport(data);
                    break;
                case "saveHtmlReport":
                    SaveHtmlReport(data);
                    break;
                case "setClipboard":
                    SetClipboard(data);
                    break;
                case "downloadImportSuccess":
                    DownloadImportSuccess();
                    break;
                case "downloadImportFail":
                    DownloadImportFail();
                    break;
                case "openNativeFile":
                    OpenNativeFile();
                    break;
                case "exportChart":
                    System.Diagnostics.Debug.WriteLine("Exporting chart");
                    ExportChart();
                    break;
                case "refreshData":
                    System.Diagnostics.Debug.WriteLine("Refreshing data");
                    RefreshData();
                    break;
                case "openCleansedFile":
                    OpenCleansedFile();
                    break;
                case "viewCleansingReport":
                    ViewCleansingReport();
                    break;
                case "openBatchLog":
                    System.Diagnostics.Debug.WriteLine("Opening batch log");
                    SendMessageToWebView(new { action = "error", message = "批量处理：打开日志暂未实现（批量引擎未接入）" });
                    break;
                case "openBatchOutputFolder":
                    System.Diagnostics.Debug.WriteLine("Opening batch output folder");
                    SendMessageToWebView(new { action = "error", message = "批量处理：打开输出文件夹暂未实现（批量引擎未接入）" });
                    break;
                case "generatePivotTable":
                    GeneratePivotTable(data);
                    break;
                case "pivotDrilldown":
                    PivotDrilldown(data);
                    break;
                case "generateChart":
                    GenerateChart(data);
                    break;
                case "executeBatchProcess":
                    SendMessageToWebView(new { action = "error", message = "批量处理：当前版本暂未实现（前端预留，后端未接入）" });
                    break;
                case "executePersonalDataMasking":
                case "executeEnterpriseDataMasking":
                case "executeComprehensiveDataMasking":
                case "executeDataRestoration":
                    SendMessageToWebView(new { action = "error", message = "脱敏/还原：当前版本暂未实现（前端预留，后端未接入）" });
                    break;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error processing message: {ex.Message}");
            SendMessageToWebView(new { action = "error", message = ex.Message });
        }
    }

    private void SetMainTable(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }
            var tableName = (data.TryGetProperty("tableName", out var tn) ? tn.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(tableName))
            {
                SendMessageToWebView(new { action = "error", message = "主表切换失败：tableName 为空" });
                return;
            }
            var tables = _sqliteManager.GetTables();
            if (!tables.Any(t => string.Equals(t, tableName, StringComparison.OrdinalIgnoreCase)))
            {
                SendMessageToWebView(new { action = "error", message = $"主表切换失败：表不存在 {tableName}" });
                return;
            }

            _currentMainTableName = tableName;
            SendMessageToWebView(new { action = "mainTableSet", tableName = tableName });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"主表切换失败: {ex.Message}" });
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

    private void SaveScheme(System.Text.Json.JsonElement data)
    {
        try
        {
            string schemeId = (data.TryGetProperty("schemeId", out var sid) ? sid.GetString() : null) ?? string.Empty;
            string schemeName = (data.TryGetProperty("schemeName", out var sn) ? sn.GetString() : null) ?? "未命名方案";
            var settingsEl = (data.TryGetProperty("settings", out var st) ? st : default);
            string settingsJson = settingsEl.ValueKind != System.Text.Json.JsonValueKind.Undefined ? settingsEl.GetRawText() : "{}";
            schemeName = string.IsNullOrWhiteSpace(schemeName) ? "未命名方案" : schemeName.Trim();

            // 方案文件名（安全化）：优先用传入 schemeId，否则由 schemeName 推导
            if (string.IsNullOrWhiteSpace(schemeId))
            {
                schemeId = SanitizeFileName(schemeName);
                // 若同名已存在，自动追加后缀避免覆盖
                var ini0 = GetSchemeIniPath(schemeId);
                if (File.Exists(ini0))
                {
                    for (int i = 2; i <= 99; i++)
                    {
                        var tryId = SanitizeFileName($"{schemeName}({i})");
                        if (!File.Exists(GetSchemeIniPath(tryId)))
                        {
                            schemeId = tryId;
                            break;
                        }
                    }
                }
            }

            // 文件指纹（修改时间）
            string fileStampsJson = "[]";
            try
            {
                using var doc = System.Text.Json.JsonDocument.Parse(settingsJson);
                if (doc.RootElement.TryGetProperty("sources", out var sources) && sources.ValueKind == System.Text.Json.JsonValueKind.Array)
                {
                    var list = new List<object>();
                    foreach (var s in sources.EnumerateArray())
                    {
                        var fp = (s.TryGetProperty("filePath", out var fpEl) ? fpEl.GetString() : null) ?? "";
                        if (string.IsNullOrWhiteSpace(fp)) continue;
                        long ticks = 0;
                        try
                        {
                            var fi = new FileInfo(fp);
                            if (fi.Exists) ticks = fi.LastWriteTimeUtc.Ticks;
                        }
                        catch { }
                        list.Add(new { filePath = fp.Replace('\\', '/'), lastWriteUtcTicks = ticks });
                    }
                    fileStampsJson = System.Text.Json.JsonSerializer.Serialize(list);
                }
            }
            catch { }

            // 写入方案 ini（DataConfig/schemes/{方案名}.ini）
            EnsureDataConfigDirs();
            var schemeIniPath = GetSchemeIniPath(schemeId);
            var ini = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
            ini["meta"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["id"] = schemeId,
                ["displayName"] = schemeName,
                ["dbFile"] = $"{schemeId}.db",
                ["lastSaveTime"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            };
            ini["settings"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["settingsJsonB64"] = ToB64(settingsJson),
                ["fileStampsJsonB64"] = ToB64(fileStampsJson)
            };
            // 预留：模板库/分组空间（后续迭代扩展）
            ini["templates"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["note"] = "templates are embedded in this scheme file"
            };
            WriteIni(schemeIniPath, ini);

            // 更新 recent.json
            var meta = _schemes.FirstOrDefault(x => string.Equals(x.Id, schemeId, StringComparison.OrdinalIgnoreCase));
            if (meta == null)
            {
                meta = new SchemeMeta { Id = schemeId };
                _schemes.Add(meta);
            }
            meta.Name = schemeName;
            meta.LastOpenTime = DateTime.Now;
            _schemes.Sort((a, b) => b.LastOpenTime.CompareTo(a.LastOpenTime));
            if (_schemes.Count > _maxSchemesKeep)
                _schemes.RemoveRange(_maxSchemesKeep, _schemes.Count - _maxSchemesKeep);

            _activeSchemeId = schemeId;
            _activeSchemeDbPath = GetSchemeDbPath(schemeId);
            Directory.CreateDirectory(Path.GetDirectoryName(_activeSchemeDbPath) ?? AppContext.BaseDirectory);
            SaveRecentSchemesToDisk();

            // 保存方案后，立即切换到方案库（文件 SQLite），并尽量将当前内存库内容备份过去（便于“保存后继续用”）
            try { SwitchSqliteToFileDb(_activeSchemeDbPath, backupFromExisting: true); } catch { }

            SendMessageToWebView(new
            {
                action = "schemeSaved",
                schemeId = schemeId,
                recentSchemes = BuildRecentSchemesPayload(),
                dbPath = _activeSchemeDbPath,
                dataConfigDir = GetDataConfigDir()
            });
        }
        catch (Exception ex)
        {
            WriteErrorLog("保存方案失败", ex);
            SendMessageToWebView(new { action = "error", message = $"保存方案失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void LoadSchemeInternal(string schemeId, bool autoLoad)
    {
        // 读取方案 ini
        var schemeIniPath = GetSchemeIniPath(schemeId);
        if (!File.Exists(schemeIniPath))
        {
            SendMessageToWebView(new { action = "error", message = $"方案未找到: {schemeId}" });
            return;
        }
        var ini = ReadIni(schemeIniPath);
        string displayName = schemeId;
        string settingsJson = "{}";
        string stampsB64 = "";
        if (ini.TryGetValue("meta", out var meta))
        {
            if (meta.TryGetValue("displayName", out var dn) && !string.IsNullOrWhiteSpace(dn)) displayName = dn;
        }
        if (ini.TryGetValue("settings", out var st))
        {
            if (st.TryGetValue("settingsJsonB64", out var sj) && !string.IsNullOrWhiteSpace(sj))
                settingsJson = FromB64(sj);
            if (st.TryGetValue("fileStampsJsonB64", out var fs) && !string.IsNullOrWhiteSpace(fs))
                stampsB64 = fs;
        }
        if (string.IsNullOrWhiteSpace(settingsJson)) settingsJson = "{}";

        _activeSchemeId = schemeId;
        _activeSchemeDbPath = GetSchemeDbPath(schemeId);
        Directory.CreateDirectory(Path.GetDirectoryName(_activeSchemeDbPath) ?? AppContext.BaseDirectory);

        // 更新 recent.json（最近方案）
        var rec = _schemes.FirstOrDefault(x => string.Equals(x.Id, schemeId, StringComparison.OrdinalIgnoreCase));
        if (rec == null)
        {
            rec = new SchemeMeta { Id = schemeId };
            _schemes.Add(rec);
        }
        rec.Name = displayName;
        rec.LastOpenTime = DateTime.Now;
        _schemes.Sort((a, b) => b.LastOpenTime.CompareTo(a.LastOpenTime));
        if (_schemes.Count > _maxSchemesKeep)
            _schemes.RemoveRange(_maxSchemesKeep, _schemes.Count - _maxSchemesKeep);
        SaveRecentSchemesToDisk();

        // 用方案里的“主表文件”回填当前文件（便于后续字段/导入/打开原生文件等功能正常工作）
        try
        {
            using var doc = System.Text.Json.JsonDocument.Parse(settingsJson);
            var root = doc.RootElement;
            if (root.TryGetProperty("main", out var main) && main.ValueKind == System.Text.Json.JsonValueKind.Object)
            {
                var fp = (main.TryGetProperty("filePath", out var fpEl) ? fpEl.GetString() : null) ?? "";
                if (!string.IsNullOrWhiteSpace(fp) && File.Exists(fp))
                {
                    _currentFilePath = fp;
                    // 仅刷新工作表下拉（不清空SQLite，避免破坏临时库）
                    try { LoadWorksheetList(fp); } catch { }
                }
            }
            if (root.TryGetProperty("sqliteMainTableName", out var mt) && mt.ValueKind == System.Text.Json.JsonValueKind.String)
            {
                var tn = mt.GetString();
                if (!string.IsNullOrWhiteSpace(tn)) _currentMainTableName = tn;
            }
        }
        catch { }

        // 校验文件修改时间（提示需重新导入）
        string notice = "";
        try
        {
            var stampsJson = FromB64(stampsB64);
            if (!string.IsNullOrWhiteSpace(stampsJson))
            {
                var stamps = System.Text.Json.JsonSerializer.Deserialize<List<Dictionary<string, object>>>(stampsJson) ?? new();
                var changed = new List<string>();
                foreach (var stamp in stamps)
                {
                    var fp = stamp.TryGetValue("filePath", out var fpo) ? (fpo?.ToString() ?? "") : "";
                    if (string.IsNullOrWhiteSpace(fp)) continue;
                    long oldTicks = 0;
                    if (stamp.TryGetValue("lastWriteUtcTicks", out var to) && long.TryParse(to?.ToString(), out var tv)) oldTicks = tv;
                    long curTicks = 0;
                    try
                    {
                        var fi = new FileInfo(fp);
                        if (fi.Exists) curTicks = fi.LastWriteTimeUtc.Ticks;
                    }
                    catch { }
                    if (oldTicks != 0 && curTicks != 0 && oldTicks != curTicks) changed.Add(Path.GetFileName(fp));
                }
                if (changed.Count > 0)
                {
                    notice = "提示：方案中的部分源文件已发生变化（" + string.Join("，", changed.Take(5)) + (changed.Count > 5 ? "..." : "") + "），建议重新导入SQLite以保证一致性。";
                }
            }
        }
        catch { }

        // 打开/切换方案临时库（文件数据库）
        try
        {
            SwitchSqliteToFileDb(_activeSchemeDbPath, backupFromExisting: false);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine("SwitchSqliteToFileDb failed: " + ex.Message);
        }

        // 下发到前端
        object? settingsObj = null;
        try
        {
            settingsObj = System.Text.Json.JsonSerializer.Deserialize<object>(settingsJson);
        }
        catch
        {
            settingsObj = new { };
        }

        SendMessageToWebView(new
        {
            action = "schemeLoaded",
            schemeId = schemeId,
            schemeName = displayName,
            settings = settingsObj,
            notice = string.IsNullOrWhiteSpace(notice) ? null : notice,
            dbPath = _activeSchemeDbPath,
            dataConfigDir = GetDataConfigDir()
        });

        // 如果临时库中已有表，顺带刷新表列表与主表显示
        try
        {
            GetTableList();
            if (!string.IsNullOrWhiteSpace(_currentMainTableName))
                SendMessageToWebView(new { action = "mainTableSet", tableName = _currentMainTableName });
        }
        catch { }

        if (!autoLoad)
        {
            SendMessageToWebView(new { action = "recentSchemesUpdated", recentSchemes = BuildRecentSchemesPayload() });
        }
    }

    private void SwitchSqliteToFileDb(string dbPath, bool backupFromExisting)
    {
        // 重新创建 SQLite 管理器与依赖服务
        try
        {
            // 若需要从现有连接（可能是内存库）备份到文件库，先做备份
            if (backupFromExisting && _sqliteManager?.Connection != null)
            {
                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(dbPath) ?? AppContext.BaseDirectory);
                    using var dest = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={dbPath}");
                    dest.Open();
                    _sqliteManager.Connection.BackupDatabase(dest);
                    dest.Close();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("BackupDatabase failed: " + ex.Message);
                }
            }
            _sqliteManager?.Dispose();
        }
        catch { }

        _sqliteManager = new SqliteManager(dbPath);
        _sqliteManager.Open();

        // 依赖服务重新指向新的连接
        _dataImporter = new DataImporter(_excelAnalyzer!, _sqliteManager);
        _queryEngine = new QueryEngine(_sqliteManager);
        _statisticsEngine = new StatisticsEngine(_sqliteManager);
        _splitEngine = new SplitEngine(_sqliteManager, _excelAnalyzer!);
    }

    private void CleanupTempDb()
    {
        try
        {
            // 说明：用户主动触发的“清理”，仅清理“孤儿库”（没有对应方案 ini 的 db）
            // 且默认只清理较旧的文件，避免误删用户希望保留的历史库。
            EnsureDataConfigDirs();
            var dir = GetDbDir();
            var schemesDir = GetSchemesDir();
            var keepPath = _activeSchemeDbPath;
            int deleted = 0;
            int total = 0;
            const int orphanRetentionDays = 30;

            foreach (var f in Directory.EnumerateFiles(dir, "*.db", SearchOption.TopDirectoryOnly))
            {
                total++;
                if (!string.IsNullOrWhiteSpace(keepPath) && string.Equals(Path.GetFullPath(f), Path.GetFullPath(keepPath), StringComparison.OrdinalIgnoreCase))
                    continue;
                var fi = new FileInfo(f);
                if (!fi.Exists) continue;

                var baseName = Path.GetFileNameWithoutExtension(fi.Name);
                var iniPath = Path.Combine(schemesDir, baseName + ".ini");
                if (File.Exists(iniPath)) continue; // 有方案，不清理

                var ageDays = (DateTime.UtcNow - fi.LastWriteTimeUtc).TotalDays;
                if (ageDays >= orphanRetentionDays)
                {
                    try { fi.Delete(); deleted++; } catch { }
                }
            }

            SendMessageToWebView(new { action = "status", message = $"清理完成：扫描 {total} 个方案库，删除 {deleted} 个孤儿库（无对应方案且超期≥{orphanRetentionDays}天）" });
        }
        catch (Exception ex)
        {
            WriteErrorLog("清理临时数据库失败", ex);
            SendMessageToWebView(new { action = "error", message = $"清理失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void RebuildSchemeDb()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(_activeSchemeId))
            {
                SendMessageToWebView(new { action = "error", message = "未绑定方案：请先保存/加载方案后再重建数据库。" });
                return;
            }
            var dbPath = GetSchemeDbPath(_activeSchemeId!);
            try
            {
                // 关闭事务（如果存在）
                _sqlLabTxn?.Dispose();
                _sqlLabTxn = null;
            }
            catch { }

            try { _sqliteManager?.Dispose(); } catch { }

            try
            {
                if (File.Exists(dbPath)) File.Delete(dbPath);
            }
            catch { }

            _activeSchemeDbPath = dbPath;
            SwitchSqliteToFileDb(dbPath, backupFromExisting: false);
            _currentMainTableName = null;

            // 刷新表列表
            try { GetTableList(); } catch { }

            SendMessageToWebView(new
            {
                action = "status",
                message = $"方案数据库已重建：{dbPath}"
            });
        }
        catch (Exception ex)
        {
            WriteErrorLog("重建方案数据库失败", ex);
            SendMessageToWebView(new { action = "error", message = $"重建方案数据库失败: {ex.Message}", hasErrorLog = true });
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
                System.Diagnostics.Debug.WriteLine($"File selected: {filePath}");
                OpenMainFileInternal(filePath, clearPrevious: true);
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

    private sealed record WorksheetStat(string Name, int RowCount, int ColCount);

    /// <summary>
    /// 快速读取每个工作表的行列数（使用 EPPlus 读取 Dimension，不遍历全数据）
    /// </summary>
    private static List<WorksheetStat> GetWorksheetStatsFast(string filePath, List<string>? preferSheetNames = null)
    {
        var list = new List<WorksheetStat>();
        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filePath));
            var wsList = package.Workbook.Worksheets;
            foreach (var ws in wsList)
            {
                if (preferSheetNames != null && preferSheetNames.Count > 0 && !preferSheetNames.Contains(ws.Name))
                    continue;
                var dim = ws.Dimension;
                int r = dim?.End.Row ?? 0;
                int c = dim?.End.Column ?? 0;
                list.Add(new WorksheetStat(ws.Name, r, c));
            }
        }
        catch
        {
            // ignore
        }
        return list;
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

    /// <summary>
    /// 文件分析（SQLite 数据源）：分析当前已导入的全部表（修复“只能识别主表”）
    /// </summary>
    private async void StartAnalysisSqliteAll()
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
                var tables = _sqliteManager.GetTables();
                // 为空时兜底主表
                if (tables == null || tables.Count == 0)
                {
                    tables = new List<string> { MainTableNameOrDefault() };
                }

                var sheets = new List<object>();
                long totalRows = 0;
                long totalFields = 0;
                var fileKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var t in tables)
                {
                    // 表名命名约定：{fileBase}|{sheet}（来自前端 mkTableName）
                    // 用于估算“来源文件数量”
                    try
                    {
                        var k = t;
                        if (k.Contains("|")) k = k.Split('|')[0];
                        if (!string.IsNullOrWhiteSpace(k)) fileKeys.Add(k);
                    }
                    catch { }

                    var schema = _sqliteManager.GetTableSchema(t);
                    int colCount = schema.Count;
                    int rowCount = _sqliteManager.GetRowCount(t);
                    totalRows += rowCount;
                    totalFields += colCount;

                    // 完整率：沿用导入统计（若有），否则默认 100
                    double completeness = 100;
                    if (_lastConversionStats != null && _lastConversionStats.Count > 0)
                    {
                        long nonEmpty = _lastConversionStats.Sum(s => (long)s.nonEmptyCount);
                        long ok = _lastConversionStats.Sum(s => (long)s.okCount);
                        if (nonEmpty > 0) completeness = Math.Round(ok * 100.0 / nonEmpty, 1);
                    }

                    sheets.Add(new
                    {
                        name = t,
                        rowCount,
                        colCount,
                        size = "",
                        completeness
                    });
                }

                return new
                {
                    fileName = Path.GetFileName(_currentFilePath ?? string.Empty),
                    fileSize = "",
                    fileCount = fileKeys.Count > 0 ? fileKeys.Count : 1,
                    totalRows = totalRows,
                    sheetCount = tables.Count,
                    fieldCount = totalFields,
                    analyzeTime = 0.0,
                    dataQuality = $"{100:F1}%（SQLite）",
                    sheets = sheets.ToArray()
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
            // 前端在“源文件清单模式”下可能未触发 BrowseMainTableFile，导致 _currentFilePath 为空
            // 因此允许前端显式传入 filePath 作为 currentFile 扫描目标
            string filePath = (data.TryGetProperty("filePath", out var fpp) ? fpp.GetString() : null) ?? string.Empty;
            bool includeBasic = !(data.TryGetProperty("includeBasic", out var ib) && ib.ValueKind == System.Text.Json.JsonValueKind.False);
            bool includeSheets = !(data.TryGetProperty("includeSheets", out var isf) && isf.ValueKind == System.Text.Json.JsonValueKind.False);
            bool includeStats = (data.TryGetProperty("includeStats", out var ist) && ist.ValueKind == System.Text.Json.JsonValueKind.True);
            bool includeRelations = (data.TryGetProperty("includeRelations", out var ir) && ir.ValueKind == System.Text.Json.JsonValueKind.True);

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
                    var cur = filePath;
                    if (string.IsNullOrWhiteSpace(cur)) cur = _currentFilePath ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(cur) || !File.Exists(cur))
                        throw new FileNotFoundException("请先选择 Excel 文件");
                    files = new[] { cur };
                }

                var fileList = files.ToList();
                int total = fileList.Count;
                int idx = 0;
                foreach (var f in fileList)
                {
                    idx++;
                    try
                    {
                        BeginInvoke(new Action(() =>
                        {
                            SendMessageToWebView(new { action = "metadataScanProgress", current = idx, total = total, fileName = Path.GetFileName(f) });
                        }));
                    }
                    catch { }

                    try
                    {
                        var fi = new FileInfo(f);
                        var sheetNames = includeSheets ? GetWorksheetNamesFast(f) : new List<string>();
                        var sheets = new List<object>();
                        long totalRows = 0;
                        if (includeSheets)
                        {
                            if (includeStats)
                            {
                                foreach (var s in GetWorksheetStatsFast(f, sheetNames))
                                {
                                    sheets.Add(new { name = s.Name, rowCount = s.RowCount, colCount = s.ColCount });
                                    totalRows += s.RowCount;
                                }
                            }
                            else
                            {
                                foreach (var s in sheetNames)
                                    sheets.Add(new { name = s });
                            }
                        }
                        list.Add(new
                        {
                            fileName = fi.Name,
                            filePath = f.Replace('\\', '/'),
                            sheetCount = includeSheets ? sheetNames.Count : 0,
                            totalRows = includeStats ? totalRows : 0,
                            fileSize = includeBasic ? FormatFileSize(fi.Length) : "",
                            lastModified = includeBasic ? fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm") : "",
                            sheets = sheets
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
                items = items,
                includeBasic,
                includeSheets,
                includeStats,
                includeRelations
            };

            SendMessageToWebView(new { action = "metadataScanComplete", results = result });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"元数据扫描失败: {ex.Message}" });
        }
    }

    private async void DetectRelations(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }

            string mode = (data.TryGetProperty("mode", out var m) ? m.GetString() : null) ?? "primary-first";
            string mainTable = (data.TryGetProperty("mainTable", out var mt) ? mt.GetString() : null) ?? MainTableNameOrDefault();
            double minCoverage = (data.TryGetProperty("minCoverage", out var mc) && mc.TryGetDouble(out var mcv)) ? mcv : 0.7;
            int sampleRows = (data.TryGetProperty("sampleRows", out var sr) && sr.TryGetInt32(out var srv)) ? srv : 50000;
            int maxPairs = (data.TryGetProperty("maxPairs", out var mp) && mp.TryGetInt32(out var mpv)) ? mpv : 80;
            bool useName = !(data.TryGetProperty("useName", out var un) && un.ValueKind == System.Text.Json.JsonValueKind.False);
            bool useValues = !(data.TryGetProperty("useValues", out var uv) && uv.ValueKind == System.Text.Json.JsonValueKind.False);
            bool useComposite = !(data.TryGetProperty("useComposite", out var uc) && uc.ValueKind == System.Text.Json.JsonValueKind.False);

            // 参与比对的表（多表）
            List<string> selectedTables = new();
            try
            {
                if (data.TryGetProperty("tables", out var t) && t.ValueKind == System.Text.Json.JsonValueKind.Array)
                {
                    foreach (var it in t.EnumerateArray())
                    {
                        var s = it.GetString();
                        if (!string.IsNullOrWhiteSpace(s)) selectedTables.Add(s);
                    }
                }
            }
            catch { }

            var sw = Stopwatch.StartNew();

            var results = await Task.Run(() =>
            {
                var mgr = _sqliteManager;
                var allTables = mgr.GetTables();
                if (allTables.Count == 0) return new List<object>();

                var tables = (selectedTables != null && selectedTables.Count > 0)
                    ? selectedTables.Where(t => allTables.Contains(t)).Distinct(StringComparer.OrdinalIgnoreCase).ToList()
                    : allTables;
                if (tables.Count == 0) return new List<object>();

                static string NormName(string s)
                {
                    s = (s ?? "").Trim().ToLowerInvariant();
                    s = s.Replace("_", "").Replace("-", "").Replace(" ", "");
                    return s;
                }
                static double JaroWinkler(string s1, string s2)
                {
                    s1 = s1 ?? ""; s2 = s2 ?? "";
                    if (s1 == s2) return 1.0;
                    int l1 = s1.Length, l2 = s2.Length;
                    if (l1 == 0 || l2 == 0) return 0.0;
                    int matchDist = Math.Max(l1, l2) / 2 - 1;
                    bool[] s1Matches = new bool[l1];
                    bool[] s2Matches = new bool[l2];
                    int matches = 0;
                    for (int i = 0; i < l1; i++)
                    {
                        int start = Math.Max(0, i - matchDist);
                        int end = Math.Min(i + matchDist + 1, l2);
                        for (int j = start; j < end; j++)
                        {
                            if (s2Matches[j]) continue;
                            if (s1[i] != s2[j]) continue;
                            s1Matches[i] = true;
                            s2Matches[j] = true;
                            matches++;
                            break;
                        }
                    }
                    if (matches == 0) return 0.0;
                    double t = 0;
                    for (int i = 0, k = 0; i < l1; i++)
                    {
                        if (!s1Matches[i]) continue;
                        while (!s2Matches[k]) k++;
                        if (s1[i] != s2[k]) t++;
                        k++;
                    }
                    t /= 2.0;
                    double m = matches;
                    double jaro = (m / l1 + m / l2 + (m - t) / m) / 3.0;
                    // winkler prefix
                    int prefix = 0;
                    for (int i = 0; i < Math.Min(4, Math.Min(l1, l2)); i++)
                    {
                        if (s1[i] == s2[i]) prefix++;
                        else break;
                    }
                    return jaro + prefix * 0.1 * (1 - jaro);
                }

                var schemaMap = new Dictionary<string, List<(string Col, string Type)>>(StringComparer.OrdinalIgnoreCase);
                foreach (var t in tables)
                {
                    try { schemaMap[t] = mgr.GetTableSchema(t); }
                    catch { schemaMap[t] = new List<(string, string)>(); }
                }

                // 枚举值采样（distinct）
                HashSet<string> SampleDistinct(string table, string col, int limit)
                {
                    var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    if (limit <= 0) return set;
                    try
                    {
                        var qt = SqliteManager.QuoteIdent(table);
                        var qc = SqliteManager.QuoteIdent(col);
                        var rows = mgr.Query($"SELECT DISTINCT CAST({qc} AS TEXT) AS v FROM {qt} WHERE {qc} IS NOT NULL LIMIT {limit}");
                        foreach (var r in rows)
                        {
                            if (!r.TryGetValue("v", out var o) || o == null) continue;
                            var s = o.ToString()?.Trim();
                            if (string.IsNullOrWhiteSpace(s)) continue;
                            // 轻量归一化
                            s = s!.Replace(" ", "").ToLowerInvariant();
                            set.Add(s);
                        }
                    }
                    catch { }
                    return set;
                }
                static double Jaccard(HashSet<string> a, HashSet<string> b)
                {
                    if (a.Count == 0 || b.Count == 0) return 0;
                    int inter = 0;
                    // iterate smaller
                    if (a.Count > b.Count) { var tmp = a; a = b; b = tmp; }
                    foreach (var x in a) if (b.Contains(x)) inter++;
                    int uni = a.Count + b.Count - inter;
                    return uni <= 0 ? 0 : (double)inter / uni;
                }

                double nameThreshold = 0.86;   // 字段名近似阈值
                double valueThreshold = 0.25;  // 枚举值近似阈值（Jaccard）
                int distinctLimit = Math.Min(5000, Math.Max(200, sampleRows / 10));

                // 生成需要扫描的表对
                var pairs = new List<(string A, string B)>();
                if (string.Equals(mode, "primary-first", StringComparison.OrdinalIgnoreCase))
                {
                    var a = string.IsNullOrWhiteSpace(mainTable) ? MainTableNameOrDefault() : mainTable;
                    if (tables.Contains(a))
                        pairs.AddRange(tables.Where(x => !string.Equals(x, a, StringComparison.OrdinalIgnoreCase)).Select(b => (a, b)));
                }
                else
                {
                    for (int i = 0; i < tables.Count; i++)
                        for (int j = i + 1; j < tables.Count; j++)
                            pairs.Add((tables[i], tables[j]));
                }

                var resultList = new List<(double Score, object Obj)>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // 1) 字段名近似
                var nameCandidates = new List<(string A, string aCol, string B, string bCol, double Sim)>();
                if (useName)
                {
                    foreach (var (A, B) in pairs)
                    {
                        var aCols = schemaMap[A].Select(x => x.Col).ToList();
                        var bCols = schemaMap[B].Select(x => x.Col).ToList();
                        foreach (var ac in aCols)
                        foreach (var bc in bCols)
                        {
                            var sim = JaroWinkler(NormName(ac), NormName(bc));
                            if (sim >= nameThreshold)
                                nameCandidates.Add((A, ac, B, bc, sim));
                        }
                    }
                }

                // 2) 枚举值近似（优先基于 nameCandidates，再补充少量直觉列）
                var valueCandidates = new List<(string A, string aCol, string B, string bCol, double NameSim)>();
                if (useValues)
                {
                    // 前 K 条 nameCandidates
                    foreach (var c in nameCandidates.OrderByDescending(x => x.Sim).Take(120))
                        valueCandidates.Add((c.A, c.aCol, c.B, c.bCol, c.Sim));

                    // 兜底：ID/CODE/NAME/编号/编码列两两配（限制规模）
                    bool IsKeyish(string col)
                    {
                        var s = (col ?? "").ToLowerInvariant();
                        return s == "id" || s.EndsWith("id") || s.EndsWith("_id")
                               || s.Contains("编号") || s.Contains("编码") || s.Contains("代码")
                               || s.Contains("name") || s.Contains("名称");
                    }
                    foreach (var (A, B) in pairs.Take(20))
                    {
                        var aCols = schemaMap[A].Select(x => x.Col).Where(IsKeyish).Take(8).ToList();
                        var bCols = schemaMap[B].Select(x => x.Col).Where(IsKeyish).Take(8).ToList();
                        foreach (var ac in aCols)
                        foreach (var bc in bCols)
                            valueCandidates.Add((A, ac, B, bc, 0));
                    }
                }

                // 缓存 distinct sets
                var distinctCache = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
                HashSet<string> GetDistinct(string t, string c)
                {
                    var key = t + "||" + c;
                    if (distinctCache.TryGetValue(key, out var set)) return set;
                    set = SampleDistinct(t, c, distinctLimit);
                    distinctCache[key] = set;
                    return set;
                }

                void AddResult(string leftTable, IEnumerable<string> leftCols, string rightTable, IEnumerable<string> rightCols,
                    string method, double score, double nameSim, double valueSim, double coverage, string? onSql)
                {
                    var lcols = leftCols.ToArray();
                    var rcols = rightCols.ToArray();
                    var sig = $"{leftTable}|{string.Join("+", lcols)}=>{rightTable}|{string.Join("+", rcols)}|{method}";
                    if (!seen.Add(sig)) return;
                    var obj = new
                    {
                        leftTable,
                        leftColumns = lcols,
                        rightTable,
                        rightColumns = rcols,
                        method,
                        score,
                        nameSimilarity = nameSim,
                        valueSimilarity = valueSim,
                        coverage,
                        onSql = onSql
                    };
                    resultList.Add((score, obj));
                }

                // 汇总：nameCandidates 直接给结果
                if (useName)
                {
                    foreach (var c in nameCandidates.OrderByDescending(x => x.Sim).Take(maxPairs))
                    {
                        var onSql = $"A.[{c.aCol}] = B.[{c.bCol}]";
                        AddResult(c.A, new[] { c.aCol }, c.B, new[] { c.bCol }, "字段名近似", c.Sim * 100, c.Sim, 0, 0, onSql);
                    }
                }

                // 汇总：valueCandidates 计算 Jaccard
                if (useValues)
                {
                    foreach (var c in valueCandidates)
                    {
                        var aSet = GetDistinct(c.A, c.aCol);
                        var bSet = GetDistinct(c.B, c.bCol);
                        var js = Jaccard(aSet, bSet);
                        if (js < valueThreshold) continue;
                        var ns = c.NameSim;
                        var score = js * 100 * 0.75 + ns * 100 * 0.25;
                        var onSql = $"A.[{c.aCol}] = B.[{c.bCol}]";
                        AddResult(c.A, new[] { c.aCol }, c.B, new[] { c.bCol }, "枚举值近似", score, ns, js, 0, onSql);
                    }
                }

                // 3) 组合字段匹配（2→1 与 1→2），基于“存在性覆盖率”
                if (useComposite)
                {
                    bool IsTexty(string type)
                    {
                        var t = (type ?? "").ToUpperInvariant();
                        return t.Contains("CHAR") || t.Contains("TEXT") || t.Contains("CLOB") || t.Contains("VARCHAR");
                    }
                    bool IsKeyish(string col)
                    {
                        var s = (col ?? "").ToLowerInvariant();
                        return s == "id" || s.EndsWith("id") || s.EndsWith("_id")
                               || s.Contains("编号") || s.Contains("编码") || s.Contains("代码")
                               || s.Contains("name") || s.Contains("名称");
                    }
                    foreach (var (A, B) in pairs)
                    {
                        var aCols = schemaMap[A].Where(x => IsTexty(x.Type) || string.IsNullOrWhiteSpace(x.Type)).Select(x => x.Col).ToList();
                        var bCols = schemaMap[B].Where(x => IsTexty(x.Type) || string.IsNullOrWhiteSpace(x.Type)).Select(x => x.Col).ToList();
                        var aPick = aCols.Where(IsKeyish).Take(10).ToList();
                        if (aPick.Count < 2) aPick = aCols.Take(10).ToList();
                        var bPick = bCols.Where(IsKeyish).Take(10).ToList();
                        if (bPick.Count < 1) bPick = bCols.Take(10).ToList();

                        // 2→1 : A.c1||c2  EXISTS in B.b
                        foreach (var b in bPick.Take(6))
                        {
                            for (int i = 0; i < aPick.Count; i++)
                            for (int j = i + 1; j < aPick.Count; j++)
                            {
                                var c1 = aPick[i];
                                var c2 = aPick[j];
                                string[] patterns =
                                {
                                    $"COALESCE(a.{SqliteManager.QuoteIdent(c1)},'') || COALESCE(a.{SqliteManager.QuoteIdent(c2)},'')",
                                    $"COALESCE(a.{SqliteManager.QuoteIdent(c1)},'') || '-' || COALESCE(a.{SqliteManager.QuoteIdent(c2)},'')",
                                    $"COALESCE(a.{SqliteManager.QuoteIdent(c1)},'') || '_' || COALESCE(a.{SqliteManager.QuoteIdent(c2)},'')",
                                };
                                double best = 0;
                                string bestOn = "";
                                foreach (var expr in patterns)
                                {
                                    try
                                    {
                                        string limit = sampleRows > 0 ? $" LIMIT {sampleRows} " : "";
                                        var q = $@"
WITH samp AS (
  SELECT {expr} AS v FROM {SqliteManager.QuoteIdent(A)} a
  WHERE (a.{SqliteManager.QuoteIdent(c1)} IS NOT NULL OR a.{SqliteManager.QuoteIdent(c2)} IS NOT NULL) {limit}
)
SELECT
  (SELECT COUNT(*) FROM samp) AS nn,
  (SELECT COUNT(*) FROM samp s WHERE EXISTS (
        SELECT 1 FROM {SqliteManager.QuoteIdent(B)} b
        WHERE CAST(b.{SqliteManager.QuoteIdent(b)} AS TEXT) = CAST(s.v AS TEXT)
  )) AS hit
";
                                        var rr = mgr.Query(q);
                                        if (rr.Count == 0) continue;
                                        var row = rr[0];
                                        var nn = Convert.ToDouble(row["nn"] ?? 0);
                                        if (nn <= 0) continue;
                                        var hit = Convert.ToDouble(row["hit"] ?? 0);
                                        var cov = hit / nn;
                                        if (cov > best)
                                        {
                                            best = cov;
                                            bestOn = $"A.{c1}+{c2} ≈ B.{b}";
                                        }
                                    }
                                    catch { }
                                }
                                if (best >= minCoverage)
                                {
                                    var score = best * 100;
                                    var onSql = $"(A.[{c1}]||A.[{c2}]) = B.[{b}]";
                                    AddResult(A, new[] { c1, c2 }, B, new[] { b }, "多字段组合", score, 0, 0, best, onSql);
                                }
                            }
                        }

                        // 1→2 : A.a  EXISTS in B.b1||b2
                        foreach (var a in aPick.Take(6))
                        {
                            for (int i = 0; i < bPick.Count; i++)
                            for (int j = i + 1; j < bPick.Count; j++)
                            {
                                var b1 = bPick[i];
                                var b2 = bPick[j];
                                string[] patterns =
                                {
                                    $"COALESCE(b.{SqliteManager.QuoteIdent(b1)},'') || COALESCE(b.{SqliteManager.QuoteIdent(b2)},'')",
                                    $"COALESCE(b.{SqliteManager.QuoteIdent(b1)},'') || '-' || COALESCE(b.{SqliteManager.QuoteIdent(b2)},'')",
                                    $"COALESCE(b.{SqliteManager.QuoteIdent(b1)},'') || '_' || COALESCE(b.{SqliteManager.QuoteIdent(b2)},'')",
                                };
                                double best = 0;
                                foreach (var expr in patterns)
                                {
                                    try
                                    {
                                        string limit = sampleRows > 0 ? $" LIMIT {sampleRows} " : "";
                                        var q = $@"
WITH samp AS (
  SELECT CAST(a.{SqliteManager.QuoteIdent(a)} AS TEXT) AS v FROM {SqliteManager.QuoteIdent(A)} a
  WHERE a.{SqliteManager.QuoteIdent(a)} IS NOT NULL {limit}
)
SELECT
  (SELECT COUNT(*) FROM samp) AS nn,
  (SELECT COUNT(*) FROM samp s WHERE EXISTS (
        SELECT 1 FROM {SqliteManager.QuoteIdent(B)} b
        WHERE CAST({expr} AS TEXT) = CAST(s.v AS TEXT)
  )) AS hit
";
                                        var rr = mgr.Query(q);
                                        if (rr.Count == 0) continue;
                                        var row = rr[0];
                                        var nn = Convert.ToDouble(row["nn"] ?? 0);
                                        if (nn <= 0) continue;
                                        var hit = Convert.ToDouble(row["hit"] ?? 0);
                                        var cov = hit / nn;
                                        if (cov > best) best = cov;
                                    }
                                    catch { }
                                }
                                if (best >= minCoverage)
                                {
                                    var score = best * 100;
                                    var onSql = $"A.[{a}] = (B.[{b1}]||B.[{b2}])";
                                    AddResult(A, new[] { a }, B, new[] { b1, b2 }, "多字段组合", score, 0, 0, best, onSql);
                                }
                            }
                        }
                    }
                }

                return resultList
                    .OrderByDescending(x => x.Score)
                    .Select(x => x.Obj)
                    .Take(Math.Max(10, maxPairs))
                    .ToList();
            });

            sw.Stop();
            SendMessageToWebView(new
            {
                action = "relationDetectComplete",
                results = new
                {
                    elapsed = Math.Round(sw.Elapsed.TotalSeconds, 2),
                    results
                }
            });
        }
        catch (Exception ex)
        {
            WriteErrorLog("关联关系识别失败", ex);
            SendMessageToWebView(new { action = "error", message = $"关联关系识别失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private sealed class RelationEnumPreview
    {
        public string Table { get; set; } = "";
        public string Column { get; set; } = "";
        public long TotalDistinct { get; set; }
        public List<string> Values { get; set; } = new();
    }

    private void GetRelationEnums(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }

            string leftTable = (data.TryGetProperty("leftTable", out var lt) ? lt.GetString() : null) ?? "";
            string leftColumn = (data.TryGetProperty("leftColumn", out var lc) ? lc.GetString() : null) ?? "";
            string rightTable = (data.TryGetProperty("rightTable", out var rt) ? rt.GetString() : null) ?? "";
            string rightColumn = (data.TryGetProperty("rightColumn", out var rc) ? rc.GetString() : null) ?? "";
            int limit = (data.TryGetProperty("limit", out var lim) && lim.ValueKind == System.Text.Json.JsonValueKind.Number) ? lim.GetInt32() : 200;

            if (string.IsNullOrWhiteSpace(leftTable) || string.IsNullOrWhiteSpace(leftColumn)
                || string.IsNullOrWhiteSpace(rightTable) || string.IsNullOrWhiteSpace(rightColumn))
            {
                SendMessageToWebView(new { action = "error", message = "枚举值预览失败：参数为空" });
                return;
            }

            RelationEnumPreview Load(string table, string col)
            {
                var res = new RelationEnumPreview { Table = table, Column = col };
                string qt = SqliteManager.QuoteIdent(table);
                string qc = SqliteManager.QuoteIdent(col);
                try
                {
                    var cnt = _sqliteManager!.Query($"SELECT COUNT(DISTINCT {qc}) AS c FROM {qt} WHERE {qc} IS NOT NULL");
                    if (cnt.Count > 0 && cnt[0].TryGetValue("c", out var v) && v != null)
                        res.TotalDistinct = Convert.ToInt64(v);
                }
                catch { }
                try
                {
                    var rows = _sqliteManager!.Query($"SELECT DISTINCT CAST({qc} AS TEXT) AS v FROM {qt} WHERE {qc} IS NOT NULL LIMIT {Math.Max(1, limit)}");
                    foreach (var r in rows)
                    {
                        if (r.TryGetValue("v", out var v) && v != null)
                        {
                            var s = v.ToString();
                            if (!string.IsNullOrWhiteSpace(s)) res.Values.Add(s!);
                        }
                    }
                }
                catch { }
                return res;
            }

            var left = Load(leftTable, leftColumn);
            var right = Load(rightTable, rightColumn);
            SendMessageToWebView(new
            {
                action = "relationEnumsLoaded",
                left = new { table = left.Table, column = left.Column, totalDistinct = left.TotalDistinct, values = left.Values },
                right = new { table = right.Table, column = right.Column, totalDistinct = right.TotalDistinct, values = right.Values }
            });
        }
        catch (Exception ex)
        {
            WriteErrorLog("查看枚举值失败", ex);
            SendMessageToWebView(new { action = "error", message = $"查看枚举值失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void ExportRelationReport(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }

            string format = (data.TryGetProperty("format", out var f) ? f.GetString() : null) ?? "xlsx";
            string mode = (data.TryGetProperty("mode", out var mo) ? mo.GetString() : null) ?? "open"; // open/saveas
            if (!string.Equals(format, "xlsx", StringComparison.OrdinalIgnoreCase))
            {
                SendMessageToWebView(new { action = "error", message = "仅支持导出 Excel（xlsx）" });
                return;
            }

            if (!data.TryGetProperty("results", out var resultsEl) || resultsEl.ValueKind != System.Text.Json.JsonValueKind.Array)
            {
                SendMessageToWebView(new { action = "error", message = "导出失败：results 为空" });
                return;
            }

            bool openAfter = string.Equals(mode, "open", StringComparison.OrdinalIgnoreCase);
            string outPath;
            if (openAfter)
            {
                outPath = Path.Combine(Path.GetTempPath(), $"relation-report-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx");
            }
            else
            {
                using var sfd = new SaveFileDialog();
                sfd.Title = "导出关联关系识别报告（Excel）";
                sfd.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
                sfd.FileName = $"relation-report-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx";
                if (sfd.ShowDialog() != DialogResult.OK) return;
                outPath = sfd.FileName;
            }

            using var wb = new ClosedXML.Excel.XLWorkbook();
            var ws = wb.AddWorksheet("候选关系");
            var headers = new[]
            {
                "置信度","左表","左字段","右表","右字段","命名相似","枚举相似","覆盖率","匹配方式","推荐ON"
            };
            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cell(1, i + 1).Value = headers[i];
                ws.Cell(1, i + 1).Style.Font.Bold = true;
                ws.Cell(1, i + 1).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#343a40");
                ws.Cell(1, i + 1).Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
            }

            int rowIdx = 2;
            foreach (var el in resultsEl.EnumerateArray())
            {
                string leftTable = (el.TryGetProperty("leftTable", out var lt) ? lt.GetString() : null) ?? "";
                string rightTable = (el.TryGetProperty("rightTable", out var rt) ? rt.GetString() : null) ?? "";
                string method = (el.TryGetProperty("method", out var md) ? md.GetString() : null) ?? "";
                string onSql = (el.TryGetProperty("onSql", out var os) ? os.GetString() : null) ?? "";

                string leftCols = "";
                if (el.TryGetProperty("leftColumns", out var lc) && lc.ValueKind == System.Text.Json.JsonValueKind.Array)
                    leftCols = string.Join("+", lc.EnumerateArray().Select(x => x.GetString()).Where(x => !string.IsNullOrWhiteSpace(x)));
                string rightCols = "";
                if (el.TryGetProperty("rightColumns", out var rc) && rc.ValueKind == System.Text.Json.JsonValueKind.Array)
                    rightCols = string.Join("+", rc.EnumerateArray().Select(x => x.GetString()).Where(x => !string.IsNullOrWhiteSpace(x)));

                double score = (el.TryGetProperty("score", out var sc) && sc.TryGetDouble(out var sv)) ? sv : 0;
                double nameSim = (el.TryGetProperty("nameSimilarity", out var ns) && ns.TryGetDouble(out var nsv)) ? nsv : 0;
                double valueSim = (el.TryGetProperty("valueSimilarity", out var vs) && vs.TryGetDouble(out var vsv)) ? vsv : 0;
                double cov = (el.TryGetProperty("coverage", out var cv) && cv.TryGetDouble(out var cvv)) ? cvv : 0;

                ws.Cell(rowIdx, 1).Value = score;
                ws.Cell(rowIdx, 2).Value = leftTable;
                ws.Cell(rowIdx, 3).Value = leftCols;
                ws.Cell(rowIdx, 4).Value = rightTable;
                ws.Cell(rowIdx, 5).Value = rightCols;
                ws.Cell(rowIdx, 6).Value = nameSim;
                ws.Cell(rowIdx, 7).Value = valueSim;
                ws.Cell(rowIdx, 8).Value = cov;
                ws.Cell(rowIdx, 9).Value = method;
                ws.Cell(rowIdx, 10).Value = onSql;
                rowIdx++;
            }

            ws.Columns().AdjustToContents();
            wb.SaveAs(outPath);

            if (openAfter)
            {
                try
                {
                    var psi = new System.Diagnostics.ProcessStartInfo(outPath) { UseShellExecute = true };
                    System.Diagnostics.Process.Start(psi);
                }
                catch { }
            }
            SendMessageToWebView(new { action = "status", message = "关联关系识别报告已导出" });
        }
        catch (Exception ex)
        {
            WriteErrorLog("导出关联关系识别报告失败", ex);
            SendMessageToWebView(new { action = "error", message = $"导出失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void ExportRelationEnums(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }
            if (!data.TryGetProperty("results", out var resultsEl) || resultsEl.ValueKind != System.Text.Json.JsonValueKind.Array)
            {
                SendMessageToWebView(new { action = "error", message = "导出失败：results 为空" });
                return;
            }
            string mode = (data.TryGetProperty("mode", out var mo) ? mo.GetString() : null) ?? "open"; // open/saveas

            // 收集所有“单字段”的 table.column
            var cols = new List<(string Table, string Col)>();
            foreach (var el in resultsEl.EnumerateArray())
            {
                if (!el.TryGetProperty("leftTable", out var lt) || !el.TryGetProperty("rightTable", out var rt)) continue;
                var leftTable = lt.GetString() ?? "";
                var rightTable = rt.GetString() ?? "";
                if (string.IsNullOrWhiteSpace(leftTable) || string.IsNullOrWhiteSpace(rightTable)) continue;

                var leftCols = new List<string>();
                if (el.TryGetProperty("leftColumns", out var lc) && lc.ValueKind == System.Text.Json.JsonValueKind.Array)
                    leftCols.AddRange(lc.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => !string.IsNullOrWhiteSpace(x)));
                var rightCols = new List<string>();
                if (el.TryGetProperty("rightColumns", out var rc) && rc.ValueKind == System.Text.Json.JsonValueKind.Array)
                    rightCols.AddRange(rc.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => !string.IsNullOrWhiteSpace(x)));

                if (leftCols.Count == 1) cols.Add((leftTable, leftCols[0]));
                if (rightCols.Count == 1) cols.Add((rightTable, rightCols[0]));
            }
            cols = cols.Distinct().ToList();
            if (cols.Count == 0)
            {
                SendMessageToWebView(new { action = "error", message = "没有可导出的单字段枚举值（组合字段暂不支持全量导出）" });
                return;
            }
            bool openAfter = string.Equals(mode, "open", StringComparison.OrdinalIgnoreCase);
            string outPath;
            if (openAfter)
            {
                outPath = Path.Combine(Path.GetTempPath(), $"relation-enums-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx");
            }
            else
            {
                using var sfd = new SaveFileDialog();
                sfd.Title = "导出全量枚举值（Excel）";
                sfd.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
                sfd.FileName = $"relation-enums-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx";
                if (sfd.ShowDialog() != DialogResult.OK) return;
                outPath = sfd.FileName;
            }

            using var wb = new ClosedXML.Excel.XLWorkbook();
            foreach (var (table, col) in cols.Take(120)) // 防炸：最多 120 个字段导出
            {
                var sheetName = $"{table}.{col}";
                // Excel sheet 名限制 31
                sheetName = sheetName.Length > 31 ? sheetName.Substring(0, 31) : sheetName;
                // 去非法字符
                foreach (var ch in new[] { ':', '\\', '/', '?', '*', '[', ']' })
                    sheetName = sheetName.Replace(ch, '_');
                if (string.IsNullOrWhiteSpace(sheetName)) sheetName = "Enums";
                if (wb.Worksheets.Any(w => string.Equals(w.Name, sheetName, StringComparison.OrdinalIgnoreCase)))
                    sheetName = sheetName.Substring(0, Math.Min(28, sheetName.Length)) + "_" + wb.Worksheets.Count;

                var ws = wb.AddWorksheet(sheetName);
                ws.Cell(1, 1).Value = "value";
                ws.Cell(1, 1).Style.Font.Bold = true;
                ws.Cell(1, 1).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#343a40");
                ws.Cell(1, 1).Style.Font.FontColor = ClosedXML.Excel.XLColor.White;

                string qt = SqliteManager.QuoteIdent(table);
                string qc = SqliteManager.QuoteIdent(col);
                var rows = _sqliteManager.Query($"SELECT DISTINCT CAST({qc} AS TEXT) AS v FROM {qt} WHERE {qc} IS NOT NULL");
                int r = 2;
                foreach (var rr in rows)
                {
                    if (r > 1048576) break; // Excel 行上限
                    if (rr.TryGetValue("v", out var v) && v != null)
                        ws.Cell(r++, 1).Value = v.ToString();
                }
                ws.Columns().AdjustToContents();
            }
            wb.SaveAs(outPath);

            if (openAfter)
            {
                try
                {
                    var psi = new System.Diagnostics.ProcessStartInfo(outPath) { UseShellExecute = true };
                    System.Diagnostics.Process.Start(psi);
                }
                catch { }
            }
            SendMessageToWebView(new { action = "status", message = "全量枚举值已导出" });
        }
        catch (Exception ex)
        {
            WriteErrorLog("导出全量枚举值失败", ex);
            SendMessageToWebView(new { action = "error", message = $"导出全量枚举值失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void ExportRelationEnumPreview(System.Text.Json.JsonElement data)
    {
        try
        {
            // mode=open：不弹保存框，直接生成临时xlsx并打开（信息量小的场景）
            // mode=saveas：弹保存框
            var mode = (data.TryGetProperty("mode", out var m) ? m.GetString() : null) ?? "open";

            if (!data.TryGetProperty("data", out var payload) || payload.ValueKind != System.Text.Json.JsonValueKind.Object)
            {
                SendMessageToWebView(new { action = "error", message = "导出失败：数据为空" });
                return;
            }

            if (!payload.TryGetProperty("left", out var left) || !payload.TryGetProperty("right", out var right))
            {
                SendMessageToWebView(new { action = "error", message = "导出失败：左右枚举值缺失" });
                return;
            }

            string lTable = (left.TryGetProperty("table", out var lt) ? lt.GetString() : null) ?? "";
            string lCol = (left.TryGetProperty("column", out var lc) ? lc.GetString() : null) ?? "";
            string rTable = (right.TryGetProperty("table", out var rt) ? rt.GetString() : null) ?? "";
            string rCol = (right.TryGetProperty("column", out var rc) ? rc.GetString() : null) ?? "";

            List<string> lVals = new();
            if (left.TryGetProperty("values", out var lvs) && lvs.ValueKind == System.Text.Json.JsonValueKind.Array)
                lVals = lvs.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
            List<string> rVals = new();
            if (right.TryGetProperty("values", out var rvs) && rvs.ValueKind == System.Text.Json.JsonValueKind.Array)
                rVals = rvs.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => !string.IsNullOrWhiteSpace(x)).ToList();

            string? outPath = null;
            if (string.Equals(mode, "saveas", StringComparison.OrdinalIgnoreCase))
            {
                using var sfd = new SaveFileDialog();
                sfd.Title = "导出枚举值预览（Excel）";
                sfd.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
                sfd.FileName = $"enum-preview-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx";
                if (sfd.ShowDialog() != DialogResult.OK) return;
                outPath = sfd.FileName;
            }
            else
            {
                // 临时目录：避免污染项目目录
                outPath = Path.Combine(Path.GetTempPath(), $"enum-preview-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx");
            }

            using var wb = new ClosedXML.Excel.XLWorkbook();

            static string MakeSheetName(string s)
            {
                s = (s ?? "").Trim();
                foreach (var ch in new[] { ':', '\\', '/', '?', '*', '[', ']' })
                    s = s.Replace(ch, '_');
                if (string.IsNullOrWhiteSpace(s)) s = "Sheet";
                if (s.Length > 31) s = s.Substring(0, 31);
                return s;
            }

            var ws1 = wb.AddWorksheet(MakeSheetName($"{lTable}.{lCol}"));
            ws1.Cell(1, 1).Value = "value";
            ws1.Cell(1, 1).Style.Font.Bold = true;
            ws1.Cell(1, 1).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#343a40");
            ws1.Cell(1, 1).Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
            for (int i = 0; i < lVals.Count && i < 1048575; i++) ws1.Cell(i + 2, 1).Value = lVals[i];
            ws1.Columns().AdjustToContents();

            var ws2 = wb.AddWorksheet(MakeSheetName($"{rTable}.{rCol}"));
            ws2.Cell(1, 1).Value = "value";
            ws2.Cell(1, 1).Style.Font.Bold = true;
            ws2.Cell(1, 1).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#343a40");
            ws2.Cell(1, 1).Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
            for (int i = 0; i < rVals.Count && i < 1048575; i++) ws2.Cell(i + 2, 1).Value = rVals[i];
            ws2.Columns().AdjustToContents();

            wb.SaveAs(outPath);

            if (!string.Equals(mode, "saveas", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    var psi = new System.Diagnostics.ProcessStartInfo(outPath) { UseShellExecute = true };
                    System.Diagnostics.Process.Start(psi);
                }
                catch { }
            }

            SendMessageToWebView(new { action = "status", message = "枚举值预览已导出" });
        }
        catch (Exception ex)
        {
            WriteErrorLog("导出枚举值预览失败", ex);
            SendMessageToWebView(new { action = "error", message = $"导出枚举值预览失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private async void ImportWorksheet(System.Text.Json.JsonElement data)
    {
        try
        {
            string worksheetName = data.GetProperty("worksheetName").GetString() ?? string.Empty;
            // 兼容：导入可能来自“主表”或“副表”，支持携带 filePath/tableName
            string filePath =
                (data.TryGetProperty("filePath", out var fp) ? fp.GetString() : null)
                ?? _currentFilePath
                ?? string.Empty;
            // WebView 侧可能传入 / 分隔符；统一转为 Windows 路径
            if (!string.IsNullOrWhiteSpace(filePath))
                filePath = filePath.Replace('/', '\\');

            string tableName =
                (data.TryGetProperty("tableName", out var tn) ? tn.GetString() : null)
                ?? (string.IsNullOrWhiteSpace(filePath) ? "Main" : BuildSqliteTableName(filePath, worksheetName));

            bool resetDb = data.TryGetProperty("resetDb", out var rd) && rd.ValueKind == System.Text.Json.JsonValueKind.True;
            bool dropExisting = data.TryGetProperty("dropExisting", out var de) && de.ValueKind == System.Text.Json.JsonValueKind.True;

            string role =
                (data.TryGetProperty("role", out var rr) ? rr.GetString() : null)
                ?? (resetDb ? "main" : "sub");

            if (!string.IsNullOrWhiteSpace(_currentFilePath)
                && string.Equals(filePath, _currentFilePath, StringComparison.OrdinalIgnoreCase))
            {
                _currentWorksheetName = worksheetName;
            }

            string importMode =
                (data.TryGetProperty("importMode", out var im) ? im.GetString() : null)
                ?? "text";

            if (string.IsNullOrEmpty(filePath) || _dataImporter == null)
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

            // 单表重导：先删除目标表，再导入（不影响其它表）
            if (dropExisting)
            {
                try
                {
                    _sqliteManager.Execute($"DROP TABLE IF EXISTS {SqliteManager.QuoteIdent(tableName)}");
                }
                catch { }
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
                filePath,
                worksheetName,
                tableName: tableName,
                importMode: importMode,
                progress: progress,
                cancellationToken: CancellationToken.None);

            // 记录“当前主表”的真实表名（长期方案：不再强依赖 Main）
            if (string.Equals(role, "main", StringComparison.OrdinalIgnoreCase))
            {
                _currentMainTableName = result.TableName;
            }

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
                    conversionStats = result.ConversionStats,
                    role = role
                }
            });

            // 导入完成后刷新表列表（驱动SQL编辑器/其他功能）
            GetTableList();
        }
        catch (Microsoft.Data.Sqlite.SqliteException ex)
        {
            // VS 输出窗口看到的“引发的异常”多为 First-chance，这里把关键信息打印出来方便定位
            WriteErrorLog("导入SQLite失败(SqliteException)", ex);
            System.Diagnostics.Debug.WriteLine($"[ImportWorksheet][SqliteException] {ex.SqliteErrorCode}/{ex.SqliteExtendedErrorCode}: {ex.Message}");
            SendMessageToWebView(new
            {
                action = "error",
                message = $"导入失败（SQLite异常 {ex.SqliteErrorCode}/{ex.SqliteExtendedErrorCode}）：{ex.Message}",
                hasErrorLog = true
            });
        }
        catch (Exception ex)
        {
            WriteErrorLog("导入SQLite失败", ex);
            SendMessageToWebView(new { action = "error", message = $"导入失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private async void ExecuteQuery(System.Text.Json.JsonElement data)
    {
        string requestId = "";
        string source = "";
        try
        {
            string sql = data.GetProperty("sql").GetString() ?? string.Empty;
            string dbType = (data.TryGetProperty("dbType", out var dt) ? dt.GetString() : null) ?? "sqlite";
            string importMode =
                (data.TryGetProperty("importMode", out var im) ? im.GetString() : null)
                ?? "text";
            requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
            source = (data.TryGetProperty("source", out var src) ? src.GetString() : null) ?? "";
            bool expertMode = !(data.TryGetProperty("expertMode", out var em) && em.ValueKind == System.Text.Json.JsonValueKind.False);
            bool allowDangerDdl = (data.TryGetProperty("allowDangerDdl", out var addl) && addl.ValueKind == System.Text.Json.JsonValueKind.True);
            bool confirmedDangerous = (data.TryGetProperty("confirmedDangerous", out var cd) && cd.ValueKind == System.Text.Json.JsonValueKind.True);
            int timeoutSeconds = (data.TryGetProperty("timeoutSeconds", out var ts) && ts.ValueKind == System.Text.Json.JsonValueKind.Number) ? ts.GetInt32() : 30;

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

            static bool IsSelectLike(string s)
            {
                var t = (s ?? string.Empty).TrimStart();
                return t.StartsWith("select", StringComparison.OrdinalIgnoreCase)
                       || t.StartsWith("with", StringComparison.OrdinalIgnoreCase)
                       || t.StartsWith("pragma", StringComparison.OrdinalIgnoreCase);
            }
            static string StripComments(string s)
            {
                if (string.IsNullOrEmpty(s)) return string.Empty;
                // remove -- ... endline
                var x = System.Text.RegularExpressions.Regex.Replace(s, @"--.*?$", "", System.Text.RegularExpressions.RegexOptions.Multiline);
                // remove /* ... */
                x = System.Text.RegularExpressions.Regex.Replace(x, @"/\*[\s\S]*?\*/", "");
                return x;
            }
            static bool IsSingleStatement(string s)
            {
                var t = (s ?? string.Empty).Trim();
                t = System.Text.RegularExpressions.Regex.Replace(t, @";+\s*$", ""); // trim trailing ;
                return !t.Contains(';');
            }

            bool selectLike = IsSelectLike(sql);

            // SQL实验室：非 SELECT 写入/DDL —— 专家模式 + 事务（提交/回滚）
            if (string.Equals(source, "sql-editor", StringComparison.OrdinalIgnoreCase) && !selectLike)
            {
                if (!expertMode)
                {
                    SendMessageToWebView(new { action = "error", message = "当前未开启【专家模式】，不允许执行写入/DDL。", hasErrorLog = false, requestId, source });
                    return;
                }
                if (!IsSingleStatement(sql))
                {
                    SendMessageToWebView(new { action = "error", message = "SQL实验室暂不支持多语句执行（请拆分后分别执行）。", hasErrorLog = false, requestId, source });
                    return;
                }

                // 后端二次校验：危险DDL & 无WHERE的UPDATE/DELETE，要求前端确认+开关
                var s0 = StripComments(sql).Trim().ToLowerInvariant();
                bool isDangerDdl = System.Text.RegularExpressions.Regex.IsMatch(s0, @"\b(drop|alter)\b");
                bool isUpdateNoWhere = System.Text.RegularExpressions.Regex.IsMatch(s0, @"\bupdate\b") && !System.Text.RegularExpressions.Regex.IsMatch(s0, @"\bwhere\b");
                bool isDeleteNoWhere = System.Text.RegularExpressions.Regex.IsMatch(s0, @"\bdelete\b") && !System.Text.RegularExpressions.Regex.IsMatch(s0, @"\bwhere\b");
                if (isDangerDdl && !allowDangerDdl)
                {
                    SendMessageToWebView(new { action = "error", message = "危险DDL（DROP/ALTER）被策略拦截：请勾选【允许危险DDL】并再次执行。", hasErrorLog = false, requestId, source });
                    return;
                }
                if ((isDangerDdl || isUpdateNoWhere || isDeleteNoWhere) && !confirmedDangerous)
                {
                    SendMessageToWebView(new { action = "error", message = "高风险语句未确认：请在弹窗确认后再次执行。", hasErrorLog = false, requestId, source });
                    return;
                }
                if (_sqliteManager?.Connection == null)
                {
                    SendMessageToWebView(new { action = "error", message = "SQLite连接未就绪", requestId, source });
                    return;
                }
                // 注意：当前方案使用文件 SQLite（{方案名}.db）以便复现与继承清洗结果；
                // SQL实验室写入被视为“实验动作”，由前端强硬风险提示+方案管理中的“重建/重导”兜底。

                var sw = Stopwatch.StartNew();
                _sqlLabTxn ??= _sqliteManager.Connection.BeginTransaction();
                _sqlExecCts?.Cancel();
                _sqlExecCts?.Dispose();
                _sqlExecCts = new CancellationTokenSource();
                int affected = await _sqliteManager.ExecuteAsync(sql, parameters: null, txn: _sqlLabTxn, timeoutSeconds: timeoutSeconds, cancellationToken: _sqlExecCts.Token);
                sw.Stop();

                SendMessageToWebView(new
                {
                    action = "queryComplete",
                    requestId = requestId,
                    source = source,
                    result = new
                    {
                        columns = Array.Empty<string>(),
                        rows = Array.Empty<object>(),
                        totalRows = affected,
                        queryTime = Math.Round(sw.Elapsed.TotalSeconds, 2),
                        sql = sql,
                        txnOpen = true
                    }
                });
                return;
            }

            // SELECT/PRAGMA：若存在未提交事务，需要绑定同一事务以便读取未提交变更
            var txn = _sqlLabTxn;
            _sqlExecCts?.Cancel();
            _sqlExecCts?.Dispose();
            _sqlExecCts = new CancellationTokenSource();
            var result = await _queryEngine.ExecuteQueryAsync(sql, txn: txn, timeoutSeconds: timeoutSeconds, cancellationToken: _sqlExecCts.Token);

            SendMessageToWebView(new
            {
                action = "queryComplete",
                requestId = requestId,
                source = source,
                result = new
                {
                    columns = result.Columns,
                    rows = result.Rows,
                    totalRows = result.TotalRows,
                    queryTime = result.QueryTime,
                    sql = result.Sql,
                    txnOpen = _sqlLabTxn != null
                }
            });
        }
        catch (OperationCanceledException)
        {
            SendMessageToWebView(new { action = "error", message = "已取消执行", hasErrorLog = false, requestId = requestId, source = source });
        }
        catch (Exception ex)
        {
            WriteErrorLog("执行SQL失败", ex);
            // 若写入/DDL 在事务中失败，自动回滚并关闭事务，避免连接卡死
            try
            {
                _sqlLabTxn?.Rollback();
                _sqlLabTxn?.Dispose();
                _sqlLabTxn = null;
                SendMessageToWebView(new { action = "sqlLabTxnState", txnOpen = false });
            }
            catch { }
            SendMessageToWebView(new { action = "error", message = $"查询执行失败: {ex.Message}", hasErrorLog = true, requestId = requestId, source = source });
        }
    }

    private void CancelSqlExecution()
    {
        try
        {
            _sqlExecCts?.Cancel();
        }
        catch { }
    }

    private void RefreshData()
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }
            // 刷新表列表
            GetTableList();
            // 刷新主表字段（用于右侧字段池/透视表字段池）
            var tableName = MainTableNameOrDefault();
            var schema = _sqliteManager.GetTableSchema(tableName);
            var fields = schema.Select(s => s.ColumnName).ToList();
            SendMessageToWebView(new { action = "activeFieldsLoaded", dbType = "sqlite", tableName, fields });
            SendMessageToWebView(new { action = "status", message = "数据已刷新" });
        }
        catch (Exception ex)
        {
            WriteErrorLog("刷新数据失败", ex);
            SendMessageToWebView(new { action = "error", message = $"刷新数据失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void ExportCleansingTemplate(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }
            // 这是一份“清洗配置模板/字段映射模板”，便于留档与复用
            using var sfd = new SaveFileDialog();
            sfd.Title = "导出清洗模板（配置）";
            sfd.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
            sfd.FileName = $"cleansing-template-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx";
            if (sfd.ShowDialog() != DialogResult.OK) return;

            var table = MainTableNameOrDefault();
            var schema = _sqliteManager.GetTableSchema(table);

            var wb = new ClosedXML.Excel.XLWorkbook();
            var wsMap = wb.Worksheets.Add("字段映射");
            wsMap.Cell(1, 1).Value = "源字段";
            wsMap.Cell(1, 2).Value = "目标字段";
            wsMap.Cell(1, 3).Value = "数据类型(text/number/date)";
            wsMap.Cell(1, 4).Value = "转换规则(trim/upper/lower)";
            wsMap.Row(1).Style.Font.Bold = true;

            int r = 2;
            foreach (var c in schema)
            {
                wsMap.Cell(r, 1).Value = c.ColumnName;
                wsMap.Cell(r, 2).Value = c.ColumnName;
                wsMap.Cell(r, 3).Value = "text";
                wsMap.Cell(r, 4).Value = "";
                r++;
            }
            wsMap.Columns().AdjustToContents();

            var wsRules = wb.Worksheets.Add("清洗规则");
            wsRules.Cell(1, 1).Value = "规则项";
            wsRules.Cell(1, 2).Value = "建议值";
            wsRules.Row(1).Style.Font.Bold = true;
            var rules = new (string, string)[]
            {
                ("trimSpaces", "true/false"),
                ("fillEmpty", "true/false"),
                ("fillMethod", "custom/forward/backward/mean/median/mode"),
                ("fillValue", "当 fillMethod=custom 时填充值"),
                ("removeEmptyRows", "true/false"),
                ("removeDuplicates", "true/false"),
                ("standardizeCase", "true/false"),
                ("output.format", "xlsx/csv"),
                ("output.target", "newfile/overwrite"),
            };
            for (int i = 0; i < rules.Length; i++)
            {
                wsRules.Cell(i + 2, 1).Value = rules[i].Item1;
                wsRules.Cell(i + 2, 2).Value = rules[i].Item2;
            }
            wsRules.Columns().AdjustToContents();

            wb.SaveAs(sfd.FileName);
            SendMessageToWebView(new { action = "status", message = "清洗模板已导出" });
        }
        catch (Exception ex)
        {
            WriteErrorLog("导出清洗模板失败", ex);
            SendMessageToWebView(new { action = "error", message = $"导出清洗模板失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void ExportCompareReport(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null || _dataImporter == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite/导入器未初始化" });
                return;
            }
            if (string.IsNullOrWhiteSpace(_currentFilePath))
            {
                SendMessageToWebView(new { action = "error", message = "请先选择主表文件" });
                return;
            }
            if (!data.TryGetProperty("settings", out var settings))
            {
                SendMessageToWebView(new { action = "error", message = "比对参数缺失" });
                return;
            }

            string afterTable = settings.TryGetProperty("afterTable", out var at) ? (at.GetString() ?? "") : "";
            string beforeSheet = settings.TryGetProperty("beforeSheet", out var bs) ? (bs.GetString() ?? "") : "";
            if (string.IsNullOrWhiteSpace(afterTable) || string.IsNullOrWhiteSpace(beforeSheet))
            {
                SendMessageToWebView(new { action = "error", message = "请先选择清洗前工作表与清洗后表" });
                return;
            }

            using var sfd = new SaveFileDialog();
            sfd.Title = "导出差异比对报告（Excel）";
            sfd.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
            sfd.FileName = $"compare-report-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx";
            if (sfd.ShowDialog() != DialogResult.OK) return;

            // 复用 ExecuteDataCompare 的核心逻辑（简化：不输出完整明细，只输出摘要 + 抽样）
            string beforeTable = "__CompareBefore";
            _sqliteManager.Execute($"DROP TABLE IF EXISTS [{beforeTable}];");
            var progress = new Progress<ImportProgress>(_ => { });
            var r = _dataImporter.ImportWorksheetAsync(_currentFilePath, beforeSheet, beforeTable, "text", progress, CancellationToken.None).GetAwaiter().GetResult();
            if (!r.Success) throw new InvalidOperationException(r.Message);

            var beforeSchema = _sqliteManager.GetTableSchema(beforeTable);
            var afterSchema = _sqliteManager.GetTableSchema(afterTable);
            var beforeCols = beforeSchema.Select(s => s.ColumnName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var afterCols = afterSchema.Select(s => s.ColumnName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var commonCols = beforeCols.Intersect(afterCols, StringComparer.OrdinalIgnoreCase).ToList();

            int beforeRows = _sqliteManager.GetRowCount(beforeTable);
            int afterRows = _sqliteManager.GetRowCount(afterTable);

            var fieldDiffs = new List<(string Type, string Field, string Before, string After)>();
            foreach (var c in beforeCols.Except(afterCols, StringComparer.OrdinalIgnoreCase))
                fieldDiffs.Add(("缺失字段", c, "存在", "-"));
            foreach (var c in afterCols.Except(beforeCols, StringComparer.OrdinalIgnoreCase))
                fieldDiffs.Add(("新增字段", c, "-", "存在"));

            long dataDiffCount = 0;
            var dataDiffSamples = new List<(string DiffType, string Payload)>();
            if (commonCols.Count > 0)
            {
                var colsSql = string.Join(", ", commonCols.Select(SqlIdent));
                var onlyBeforeCnt = SqlScalarLong($"SELECT COUNT(*) AS v FROM (SELECT {colsSql} FROM [{beforeTable}] EXCEPT SELECT {colsSql} FROM [{afterTable}]) t");
                var onlyAfterCnt = SqlScalarLong($"SELECT COUNT(*) AS v FROM (SELECT {colsSql} FROM [{afterTable}] EXCEPT SELECT {colsSql} FROM [{beforeTable}]) t");
                dataDiffCount = onlyBeforeCnt + onlyAfterCnt;

                var sampleRows = _sqliteManager.Query($"SELECT * FROM (SELECT {colsSql} FROM [{beforeTable}] EXCEPT SELECT {colsSql} FROM [{afterTable}]) t LIMIT 50");
                foreach (var row in sampleRows)
                    dataDiffSamples.Add(("仅清洗前存在", System.Text.Json.JsonSerializer.Serialize(row)));
                var sampleRows2 = _sqliteManager.Query($"SELECT * FROM (SELECT {colsSql} FROM [{afterTable}] EXCEPT SELECT {colsSql} FROM [{beforeTable}]) t LIMIT 50");
                foreach (var row in sampleRows2)
                    dataDiffSamples.Add(("仅清洗后存在", System.Text.Json.JsonSerializer.Serialize(row)));
            }

            var wb = new ClosedXML.Excel.XLWorkbook();
            var wsSum = wb.Worksheets.Add("摘要");
            wsSum.Cell(1, 1).Value = "项";
            wsSum.Cell(1, 2).Value = "值";
            wsSum.Row(1).Style.Font.Bold = true;
            var sum = new (string, object)[]
            {
                ("清洗前工作表", beforeSheet),
                ("清洗后表", afterTable),
                ("清洗前行数", beforeRows),
                ("清洗后行数", afterRows),
                ("字段差异数", fieldDiffs.Count),
                ("数据差异数(估算)", dataDiffCount),
            };
            for (int i = 0; i < sum.Length; i++)
            {
                wsSum.Cell(i + 2, 1).Value = sum[i].Item1;
                wsSum.Cell(i + 2, 2).Value = sum[i].Item2?.ToString() ?? "";
            }
            wsSum.Columns().AdjustToContents();

            var wsF = wb.Worksheets.Add("字段差异");
            wsF.Cell(1, 1).Value = "类型";
            wsF.Cell(1, 2).Value = "字段";
            wsF.Cell(1, 3).Value = "清洗前";
            wsF.Cell(1, 4).Value = "清洗后";
            wsF.Row(1).Style.Font.Bold = true;
            for (int i = 0; i < fieldDiffs.Count; i++)
            {
                wsF.Cell(i + 2, 1).Value = fieldDiffs[i].Type;
                wsF.Cell(i + 2, 2).Value = fieldDiffs[i].Field;
                wsF.Cell(i + 2, 3).Value = fieldDiffs[i].Before;
                wsF.Cell(i + 2, 4).Value = fieldDiffs[i].After;
            }
            wsF.Columns().AdjustToContents();

            var wsD = wb.Worksheets.Add("数据差异抽样");
            wsD.Cell(1, 1).Value = "差异类型";
            wsD.Cell(1, 2).Value = "行内容(JSON)";
            wsD.Row(1).Style.Font.Bold = true;
            for (int i = 0; i < dataDiffSamples.Count; i++)
            {
                wsD.Cell(i + 2, 1).Value = dataDiffSamples[i].DiffType;
                wsD.Cell(i + 2, 2).Value = dataDiffSamples[i].Payload;
            }
            wsD.Column(2).Style.Alignment.WrapText = true;
            wsD.Columns().AdjustToContents();

            wb.SaveAs(sfd.FileName);
            SendMessageToWebView(new { action = "status", message = "差异比对报告已导出" });
        }
        catch (Exception ex)
        {
            WriteErrorLog("导出比对报告失败", ex);
            SendMessageToWebView(new { action = "error", message = $"导出比对报告失败: {ex.Message}", hasErrorLog = true });
        }
        finally
        {
            try { _sqliteManager?.Execute("DROP TABLE IF EXISTS [__CompareBefore];"); } catch { }
        }
    }

    private void GeneratePivotTable(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }
            var rowFields = new List<string>();
            var colFields = new List<string>();
            var valFields = new List<string>();

            if (data.TryGetProperty("rowFields", out var rf) && rf.ValueKind == System.Text.Json.JsonValueKind.Array)
                rowFields = rf.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
            if (data.TryGetProperty("colFields", out var cf) && cf.ValueKind == System.Text.Json.JsonValueKind.Array)
                colFields = cf.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
            if (data.TryGetProperty("valueFields", out var vf) && vf.ValueKind == System.Text.Json.JsonValueKind.Array)
                valFields = vf.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => !string.IsNullOrWhiteSpace(x)).ToList();

            if (valFields.Count == 0)
            {
                SendMessageToWebView(new { action = "error", message = "透视表：请至少配置一个值字段" });
                return;
            }

            var table = MainTableNameOrDefault();

            // 仅实现：最多 1 个列字段（多列字段可在前端用拼接字段实现）
            string? colField = colFields.FirstOrDefault();
            var groupExprs = new List<string>();
            var selectParts = new List<string>();

            foreach (var rfName in rowFields)
            {
                groupExprs.Add(SqlIdent(rfName));
                selectParts.Add($"{SqlIdent(rfName)} AS {SqlIdent(rfName)}");
            }

            if (!string.IsNullOrWhiteSpace(colField))
            {
                groupExprs.Add(SqlIdent(colField));
                selectParts.Add($"{SqlIdent(colField)} AS {SqlIdent(colField)}");
            }

            foreach (var v in valFields)
            {
                // 默认 SUM，按 REAL 聚合
                selectParts.Add($"SUM(CAST(NULLIF(CAST({SqlIdent(v)} AS TEXT),'') AS REAL)) AS {SqlIdent("SUM_" + v)}");
            }

            // 过滤字段（前端“过滤字段”区）：简单等值/IN（逗号分隔多个值）
            var where = new List<string>();
            if (data.TryGetProperty("filters", out var filters) && filters.ValueKind == System.Text.Json.JsonValueKind.Object)
            {
                foreach (var p in filters.EnumerateObject())
                {
                    var field = p.Name ?? "";
                    var val = p.Value.ValueKind == System.Text.Json.JsonValueKind.String ? (p.Value.GetString() ?? "") : p.Value.ToString();
                    if (string.IsNullOrWhiteSpace(field) || string.IsNullOrWhiteSpace(val)) continue;
                    var parts = val.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).Where(x => x.Length > 0).ToList();
                    if (parts.Count <= 1)
                    {
                        where.Add($"CAST({SqlIdent(field)} AS TEXT) = {SqlValue(parts.Count == 1 ? parts[0] : val.Trim())}");
                    }
                    else
                    {
                        var inList = string.Join(", ", parts.Select(SqlValue));
                        where.Add($"CAST({SqlIdent(field)} AS TEXT) IN ({inList})");
                    }
                }
            }

            string sql = $"SELECT {string.Join(", ", selectParts)} FROM [{table}]";
            if (where.Count > 0) sql += " WHERE " + string.Join(" AND ", where);
            if (groupExprs.Count > 0)
                sql += $" GROUP BY {string.Join(", ", groupExprs)}";
            sql += " LIMIT 20000"; // 防爆：透视表先限制规模

            var raw = _sqliteManager.Query(sql);

            // 组装输出列/行（二维透视）
            var outCols = new List<string>();
            outCols.AddRange(rowFields);
            var outRows = new List<Dictionary<string, object>>();

            if (string.IsNullOrWhiteSpace(colField))
            {
                // 无列字段：直接输出 rowFields + SUM_xxx
                outCols.AddRange(valFields.Select(v => "SUM_" + v));
                foreach (var r0 in raw)
                {
                    var o = new Dictionary<string, object>();
                    foreach (var k in outCols)
                        o[k] = r0.TryGetValue(k, out var vv) ? vv : "";
                    outRows.Add(o);
                }
            }
            else
            {
                // 有列字段：pivot
                var colValues = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var r0 in raw)
                {
                    var cv = r0.TryGetValue(colField, out var vv) ? (vv?.ToString() ?? "") : "";
                    colValues.Add(cv);
                }
                // 输出列：rowFields + (colValue + ':' + SUM_field)
                foreach (var cv in colValues)
                {
                    foreach (var v in valFields)
                        outCols.Add($"{cv}:SUM_{v}");
                }

                // key by rowFields tuple
                string KeyFor(Dictionary<string, object> r0)
                {
                    return string.Join("|", rowFields.Select(f => r0.TryGetValue(f, out var vv) ? (vv?.ToString() ?? "") : ""));
                }

                var rowMap = new Dictionary<string, Dictionary<string, object>>();
                foreach (var r0 in raw)
                {
                    var key = KeyFor(r0);
                    if (!rowMap.TryGetValue(key, out var obj))
                    {
                        obj = new Dictionary<string, object>();
                        foreach (var f in rowFields)
                            obj[f] = r0.TryGetValue(f, out var vv) ? vv : "";
                        rowMap[key] = obj;
                    }
                    var cv = r0.TryGetValue(colField, out var vv2) ? (vv2?.ToString() ?? "") : "";
                    foreach (var v in valFields)
                    {
                        var k = $"{cv}:SUM_{v}";
                        var srcKey = "SUM_" + v;
                        obj[k] = r0.TryGetValue(srcKey, out var vv3) ? vv3 : 0;
                    }
                }
                outRows.AddRange(rowMap.Values);
            }

            SendMessageToWebView(new { action = "pivotTableGenerated", columns = outCols, rows = outRows });
        }
        catch (Exception ex)
        {
            WriteErrorLog("生成透视表失败", ex);
            SendMessageToWebView(new { action = "error", message = $"生成透视表失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void PivotDrilldown(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }

            string table = (data.TryGetProperty("tableName", out var tn) ? tn.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(table)) table = MainTableNameOrDefault();

            int limit = 200;
            if (data.TryGetProperty("limit", out var lm) && lm.ValueKind == System.Text.Json.JsonValueKind.Number)
                limit = Math.Max(1, Math.Min(2000, lm.GetInt32()));

            var where = new List<string>();
            if (data.TryGetProperty("filters", out var filters) && filters.ValueKind == System.Text.Json.JsonValueKind.Object)
            {
                foreach (var p in filters.EnumerateObject())
                {
                    var field = p.Name ?? "";
                    var val = p.Value.ValueKind == System.Text.Json.JsonValueKind.String ? (p.Value.GetString() ?? "") : p.Value.ToString();
                    if (string.IsNullOrWhiteSpace(field) || string.IsNullOrWhiteSpace(val)) continue;
                    // 透视表分组/显示按 TEXT 处理：统一 CAST 为 TEXT 做等值匹配
                    where.Add($"CAST({SqlIdent(field)} AS TEXT) = {SqlValue(val)}");
                }
            }

            string whereSql = where.Count > 0 ? (" WHERE " + string.Join(" AND ", where)) : "";
            string sql = $"SELECT * FROM [{table}] {whereSql} LIMIT {limit}";

            var rows = _sqliteManager.Query(sql);
            var cols = new List<string>();
            if (rows.Count > 0)
            {
                // 保持字段顺序：按第一行 key 顺序（Query 实现一般能保持）
                cols = rows[0].Keys.ToList();
            }

            SendMessageToWebView(new
            {
                action = "pivotDrilldownResult",
                tableName = table,
                columns = cols,
                rows = rows
            });
        }
        catch (Exception ex)
        {
            WriteErrorLog("透视表钻取明细失败", ex);
            SendMessageToWebView(new { action = "error", message = $"透视表钻取明细失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void GenerateChart(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }

            string chartType = (data.TryGetProperty("chartType", out var ct) ? ct.GetString() : null) ?? "bar";
            string xField = (data.TryGetProperty("xField", out var xf) ? xf.GetString() : null) ?? "";
            string yField = (data.TryGetProperty("yField", out var yf) ? yf.GetString() : null) ?? "";
            string title = (data.TryGetProperty("title", out var ti) ? ti.GetString() : null) ?? "数据图表";
            string dataSource = (data.TryGetProperty("dataSource", out var ds) ? ds.GetString() : null) ?? "current";

            if (string.IsNullOrWhiteSpace(xField) || string.IsNullOrWhiteSpace(yField))
            {
                SendMessageToWebView(new { action = "error", message = "请选择 X 轴字段与 Y 轴字段" });
                return;
            }

            // 当前先支持：从“当前主表（SQLite）”生成图表
            // SQL查询结果作为数据源（dataSource=query）属于后续增强项
            if (!string.Equals(dataSource, "current", StringComparison.OrdinalIgnoreCase))
            {
                SendMessageToWebView(new { action = "error", message = "当前版本图表仅支持“当前工作表（SQLite主表）”作为数据源" });
                return;
            }

            var table = MainTableNameOrDefault();
            // 生成聚合数据：按 X 分组，对 Y 做 SUM（支持文本数字）
            string x = SqlIdent(xField);
            string y = SqlIdent(yField);
            string sql =
                $"SELECT CAST({x} AS TEXT) AS x, " +
                $"SUM(CAST(NULLIF(CAST({y} AS TEXT),'') AS REAL)) AS y " +
                $"FROM [{table}] " +
                $"GROUP BY CAST({x} AS TEXT) " +
                $"ORDER BY y DESC " +
                $"LIMIT 50";

            var rows = _sqliteManager.Query(sql);
            if (rows.Count == 0)
            {
                SendMessageToWebView(new { action = "error", message = "没有可用于生成图表的数据（结果为空）" });
                return;
            }

            var labels = new List<string>();
            var values = new List<double>();
            foreach (var r in rows)
            {
                var lx = r.TryGetValue("x", out var xv) ? (xv?.ToString() ?? "") : "";
                var ly = r.TryGetValue("y", out var yv) ? yv : null;
                double dv = 0;
                if (ly != null) double.TryParse(Convert.ToString(ly), out dv);
                labels.Add(lx);
                values.Add(dv);
            }

            // 输出 png 到临时目录
            var outPath = Path.Combine(Path.GetTempPath(), $"chart-{DateTime.Now:yyyyMMdd-HHmmss}.png");
            RenderSimpleChartPng(outPath, title, chartType, xField, yField, labels, values);
            _lastChartPngPath = outPath;

            var bytes = File.ReadAllBytes(outPath);
            var b64 = Convert.ToBase64String(bytes);
            SendMessageToWebView(new
            {
                action = "chartGenerated",
                chart = new
                {
                    title,
                    chartType,
                    xField,
                    yField,
                    imageBase64 = b64
                }
            });
        }
        catch (Exception ex)
        {
            WriteErrorLog("生成图表失败", ex);
            SendMessageToWebView(new { action = "error", message = $"生成图表失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private static void RenderSimpleChartPng(
        string outputPath,
        string title,
        string chartType,
        string xField,
        string yField,
        List<string> labels,
        List<double> values)
    {
        const int W = 960;
        const int H = 540;
        using var bmp = new Bitmap(W, H);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        g.Clear(Color.White);

        using var fontTitle = new Font("Microsoft YaHei", 14, FontStyle.Bold);
        using var fontAxis = new Font("Microsoft YaHei", 9, FontStyle.Regular);
        using var fontSmall = new Font("Microsoft YaHei", 8, FontStyle.Regular);
        using var penAxis = new Pen(Color.FromArgb(120, 120, 120), 1);
        using var penGrid = new Pen(Color.FromArgb(230, 230, 230), 1);
        using var brushMain = new SolidBrush(Color.FromArgb(0, 120, 212));

        // 标题
        g.DrawString(title, fontTitle, Brushes.Black, new PointF(16, 10));
        g.DrawString($"{xField} →  {yField}（SUM）", fontSmall, Brushes.Gray, new PointF(18, 36));

        var type = (chartType ?? "bar").ToLowerInvariant();
        var palette = new[]
        {
            Color.FromArgb(0,120,212),
            Color.FromArgb(82,196,26),
            Color.FromArgb(250,173,20),
            Color.FromArgb(245,34,45),
            Color.FromArgb(114,46,209),
            Color.FromArgb(19,194,194),
        };

        if (type == "pie")
        {
            var sum = values.Sum();
            if (sum <= 0) sum = 1;
            var pieRect = new Rectangle(60, 90, 360, 360);
            float start = 0;
            for (int i = 0; i < values.Count; i++)
            {
                float sweep = (float)(values[i] / sum * 360.0);
                using var b = new SolidBrush(palette[i % palette.Length]);
                g.FillPie(b, pieRect, start, sweep);
                start += sweep;
            }
            g.DrawEllipse(Pens.LightGray, pieRect);

            // legend
            float lx = 460, ly = 100;
            for (int i = 0; i < labels.Count; i++)
            {
                using var b = new SolidBrush(palette[i % palette.Length]);
                g.FillRectangle(b, lx, ly + i * 18, 10, 10);
                g.DrawString($"{labels[i]} ({values[i]:0.##})", fontAxis, Brushes.Black, lx + 14, ly + i * 18 - 2);
                if (i >= 18) break;
            }
        }
        else
        {
            // plot area
            int left = 70, top = 80, right = 930, bottom = 500;
            var plot = new Rectangle(left, top, right - left, bottom - top);

            // y range
            double maxY = values.Count > 0 ? values.Max() : 0;
            if (maxY <= 0) maxY = 1;
            maxY *= 1.1;

            // grid + y ticks
            int ticks = 5;
            for (int t = 0; t <= ticks; t++)
            {
                float y = bottom - (float)(t * 1.0 / ticks * plot.Height);
                g.DrawLine(penGrid, left, y, right, y);
                var val = maxY * t / ticks;
                g.DrawString(val.ToString("0.##"), fontSmall, Brushes.Gray, 10, y - 6);
            }

            // axes
            g.DrawLine(penAxis, left, bottom, right, bottom);
            g.DrawLine(penAxis, left, top, left, bottom);

            int n = labels.Count;
            if (n == 0) { bmp.Save(outputPath, ImageFormat.Png); return; }
            float step = plot.Width / (float)Math.Max(1, n);

            Func<double, float> yMap = (v) => bottom - (float)(v / maxY * plot.Height);

            if (type == "scatter")
            {
                for (int i = 0; i < n; i++)
                {
                    float x = left + step * (i + 0.5f);
                    float y = yMap(values[i]);
                    g.FillEllipse(brushMain, x - 4, y - 4, 8, 8);
                }
            }
            else if (type == "line" || type == "area")
            {
                var pts = new List<PointF>();
                for (int i = 0; i < n; i++)
                {
                    float x = left + step * (i + 0.5f);
                    float y = yMap(values[i]);
                    pts.Add(new PointF(x, y));
                }
                if (type == "area")
                {
                    var poly = new List<PointF>(pts);
                    poly.Add(new PointF(left + step * (n - 0.5f), bottom));
                    poly.Add(new PointF(left + step * 0.5f, bottom));
                    using var b = new SolidBrush(Color.FromArgb(60, 0, 120, 212));
                    g.FillPolygon(b, poly.ToArray());
                }
                using var penLine = new Pen(Color.FromArgb(0, 120, 212), 2);
                g.DrawLines(penLine, pts.ToArray());
                foreach (var p in pts) g.FillEllipse(brushMain, p.X - 3, p.Y - 3, 6, 6);
            }
            else // bar default
            {
                float barW = Math.Max(6, step * 0.6f);
                for (int i = 0; i < n; i++)
                {
                    float x = left + step * (i + 0.5f) - barW / 2;
                    float y = yMap(values[i]);
                    g.FillRectangle(brushMain, x, y, barW, bottom - y);
                }
            }

            // x labels（抽样显示，避免拥挤）
            int every = n <= 10 ? 1 : (n <= 25 ? 2 : 5);
            for (int i = 0; i < n; i += every)
            {
                float x = left + step * (i + 0.5f);
                var s = labels[i] ?? "";
                if (s.Length > 10) s = s[..10] + "…";
                var size = g.MeasureString(s, fontSmall);
                g.DrawString(s, fontSmall, Brushes.Gray, x - size.Width / 2, bottom + 6);
            }
        }

        bmp.Save(outputPath, ImageFormat.Png);
    }

    private void ExportChart()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(_lastChartPngPath) || !File.Exists(_lastChartPngPath))
            {
                SendMessageToWebView(new { action = "error", message = "暂无可导出的图表：请先生成图表" });
                return;
            }

            using var sfd = new SaveFileDialog();
            sfd.Title = "导出图表（PNG）";
            sfd.Filter = "PNG 图片 (*.png)|*.png|所有文件 (*.*)|*.*";
            sfd.FileName = $"chart-{DateTime.Now:yyyyMMdd-HHmmss}.png";
            if (sfd.ShowDialog() != DialogResult.OK) return;

            File.Copy(_lastChartPngPath, sfd.FileName, overwrite: true);
            try
            {
                var dr = MessageBox.Show("导出成功，是否打开图片？", "导出图表", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (dr == DialogResult.Yes)
                    Process.Start(new ProcessStartInfo { FileName = sfd.FileName, UseShellExecute = true });
            }
            catch { }
        }
        catch (Exception ex)
        {
            WriteErrorLog("导出图表失败", ex);
            SendMessageToWebView(new { action = "error", message = $"导出图表失败: {ex.Message}", hasErrorLog = true });
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

    private void SqlLabCommit()
    {
        try
        {
            if (_sqlLabTxn == null)
            {
                SendMessageToWebView(new { action = "sqlLabTxnState", txnOpen = false });
                return;
            }
            _sqlLabTxn.Commit();
            _sqlLabTxn.Dispose();
            _sqlLabTxn = null;
            SendMessageToWebView(new { action = "sqlLabTxnState", txnOpen = false });
            SendMessageToWebView(new { action = "status", message = "已提交变更" });
        }
        catch (Exception ex)
        {
            WriteErrorLog("提交变更失败", ex);
            SendMessageToWebView(new { action = "error", message = $"提交变更失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void SqlLabRollback()
    {
        try
        {
            if (_sqlLabTxn == null)
            {
                SendMessageToWebView(new { action = "sqlLabTxnState", txnOpen = false });
                return;
            }
            _sqlLabTxn.Rollback();
            _sqlLabTxn.Dispose();
            _sqlLabTxn = null;
            SendMessageToWebView(new { action = "sqlLabTxnState", txnOpen = false });
            SendMessageToWebView(new { action = "status", message = "已回滚变更" });
        }
        catch (Exception ex)
        {
            WriteErrorLog("回滚变更失败", ex);
            SendMessageToWebView(new { action = "error", message = $"回滚变更失败: {ex.Message}", hasErrorLog = true });
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

            string tableName = (data.TryGetProperty("tableName", out var tn) ? tn.GetString() : null) ?? MainTableNameOrDefault();
            var schema = _sqliteManager.GetTableSchema(tableName);
            var fields = schema.Select(s => s.ColumnName).ToList();

            // 新协议：专门给“多表关联/多表统计”等按表取字段使用
            SendMessageToWebView(new
            {
                action = "sqliteTableFieldsLoaded",
                tableName = tableName,
                fields = fields,
                schema = schema.Select(s => new { name = s.ColumnName, type = s.DataType }).ToList()
            });
            // 兼容旧前端：仍发送 worksheetFieldsLoaded
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

                string tableName = (data.TryGetProperty("tableName", out var tn) ? tn.GetString() : null) ?? MainTableNameOrDefault();
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

    private void ExecuteMultiTableSplit(System.Text.Json.JsonElement data)
    {
        string tempTable = "TempSplit_" + Guid.NewGuid().ToString("N");
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }
            if (_splitEngine == null)
            {
                SendMessageToWebView(new { action = "error", message = "分拆引擎未初始化" });
                return;
            }

            string sql = data.GetProperty("sql").GetString() ?? string.Empty;
            string splitField = data.GetProperty("splitField").GetString() ?? string.Empty;
            string outputDirectory = data.GetProperty("outputDirectory").GetString() ?? string.Empty;
            string splitType = data.GetProperty("splitType").GetString() ?? "byField";

            if (string.IsNullOrWhiteSpace(sql) || string.IsNullOrWhiteSpace(splitField) || string.IsNullOrWhiteSpace(outputDirectory))
            {
                SendMessageToWebView(new { action = "error", message = "多表分拆参数不完整" });
                return;
            }

            // 创建临时表（只保存关联后的主表列：由前端 SQL 决定）
            _sqliteManager.Execute($"DROP TABLE IF EXISTS [{tempTable}];");
            var sqlNoSemi = sql.Trim().TrimEnd(';');
            _sqliteManager.Execute($"CREATE TABLE [{tempTable}] AS {sqlNoSemi};");

            SplitResult result;
            if (splitType == "byField")
            {
                result = _splitEngine.SplitByField(tempTable, splitField, outputDirectory);
            }
            else
            {
                int rowsPerFile = data.GetProperty("rowsPerFile").GetInt32();
                result = _splitEngine.SplitByRowCount(tempTable, rowsPerFile, outputDirectory);
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
            SendMessageToWebView(new { action = "error", message = $"多表分拆失败: {ex.Message}" });
        }
        finally
        {
            try
            {
                _sqliteManager?.Execute($"DROP TABLE IF EXISTS [{tempTable}];");
            }
            catch { }
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
            string mode = (data.TryGetProperty("mode", out var md) ? md.GetString() : null) ?? "saveas"; // saveas/open
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

            bool openAfter = string.Equals(mode, "open", StringComparison.OrdinalIgnoreCase);
            string outPath;
            if (openAfter)
            {
                var ext = string.Equals(format, "csv", StringComparison.OrdinalIgnoreCase) ? "csv" : "xlsx";
                outPath = Path.Combine(Path.GetTempPath(), $"query-{DateTime.Now:yyyyMMdd-HHmmss}.{ext}");
            }
            else
            {
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
                outPath = sfd.FileName;
            }

            // 执行导出：为避免一次性吃内存，这里用 SqliteDataReader 流式写 CSV
            var exportSql = sql.Trim().TrimEnd(';');
            if (limit > 0 && !exportSql.Contains(" limit ", StringComparison.OrdinalIgnoreCase) && exportSql.TrimStart().StartsWith("select", StringComparison.OrdinalIgnoreCase))
            {
                exportSql = $"SELECT * FROM ({exportSql}) t LIMIT {limit}";
            }

            if (string.Equals(format, "csv", StringComparison.OrdinalIgnoreCase))
            {
                ExportSqlToCsv(exportSql, outPath, protectForExcel: true);
            }
            else
            {
                // 说明：需要安装 ClosedXML NuGet：ClosedXML
                ExportSqlToXlsx(exportSql, outPath, forceAllText);
            }

            if (openAfter)
            {
                try
                {
                    var psi = new System.Diagnostics.ProcessStartInfo(outPath) { UseShellExecute = true };
                    System.Diagnostics.Process.Start(psi);
                }
                catch { }
            }
        }
        catch (Exception ex)
        {
            WriteErrorLog("导出查询结果失败", ex);
            SendMessageToWebView(new { action = "error", message = $"导出失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void ExportGridToFile(System.Text.Json.JsonElement data)
    {
        try
        {
            string mode = (data.TryGetProperty("mode", out var md) ? md.GetString() : null) ?? "saveas"; // saveas/open
            string title = (data.TryGetProperty("title", out var ti) ? ti.GetString() : null) ?? "SQL编辑器导出";
            string fileName = (data.TryGetProperty("fileName", out var fn) ? fn.GetString() : null) ?? $"grid-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx";
            if (!data.TryGetProperty("columns", out var colsEl) || colsEl.ValueKind != System.Text.Json.JsonValueKind.Array)
            {
                SendMessageToWebView(new { action = "error", message = "导出失败：columns 为空" });
                return;
            }
            if (!data.TryGetProperty("rows", out var rowsEl) || rowsEl.ValueKind != System.Text.Json.JsonValueKind.Array)
            {
                SendMessageToWebView(new { action = "error", message = "导出失败：rows 为空" });
                return;
            }

            var cols = colsEl.EnumerateArray().Select(x => x.GetString() ?? "").ToList();
            if (cols.Count == 0)
            {
                SendMessageToWebView(new { action = "error", message = "导出失败：columns 为空" });
                return;
            }

            bool openAfter = string.Equals(mode, "open", StringComparison.OrdinalIgnoreCase);
            string outPath;
            if (openAfter)
            {
                outPath = Path.Combine(Path.GetTempPath(), fileName);
                if (!outPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    outPath += ".xlsx";
            }
            else
            {
                using var sfd = new SaveFileDialog();
                sfd.Title = title;
                sfd.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
                sfd.FileName = fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ? fileName : (fileName + ".xlsx");
                if (sfd.ShowDialog() != DialogResult.OK) return;
                outPath = sfd.FileName;
            }

            using var wb = new ClosedXML.Excel.XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            // Header
            for (int i = 0; i < cols.Count; i++)
            {
                ws.Cell(1, i + 1).Value = cols[i];
                ws.Cell(1, i + 1).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#343a40");
                ws.Cell(1, i + 1).Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
                ws.Cell(1, i + 1).Style.Font.Bold = true;
            }
            int rIdx = 2;
            foreach (var rowEl in rowsEl.EnumerateArray())
            {
                for (int c = 0; c < cols.Count; c++)
                {
                    string col = cols[c];
                    string text = "";
                    if (rowEl.ValueKind == System.Text.Json.JsonValueKind.Object)
                    {
                        if (rowEl.TryGetProperty(col, out var v))
                            text = v.ValueKind == System.Text.Json.JsonValueKind.String ? (v.GetString() ?? "") : v.ToString();
                        else
                            text = "";
                    }
                    else
                    {
                        text = rowEl.ToString();
                    }
                    ws.Cell(rIdx, c + 1).Value = text ?? "";
                }
                rIdx++;
            }
            ws.Columns().AdjustToContents();
            wb.SaveAs(outPath);

            if (openAfter)
            {
                try
                {
                    var psi = new System.Diagnostics.ProcessStartInfo(outPath) { UseShellExecute = true };
                    System.Diagnostics.Process.Start(psi);
                }
                catch { }
            }
            SendMessageToWebView(new { action = "status", message = "表格已导出" });
        }
        catch (Exception ex)
        {
            WriteErrorLog("导出表格失败", ex);
            SendMessageToWebView(new { action = "error", message = $"导出失败: {ex.Message}", hasErrorLog = true });
        }
    }

    /// <summary>
    /// 保存 HTML 报告到 exe 目录下（reports 子目录），并可选择保存后自动打开
    /// 前端用于：文件分析/工作表分析/后续元数据扫描等报告导出
    /// </summary>
    private void SaveHtmlReport(System.Text.Json.JsonElement data)
    {
        try
        {
            string html = (data.TryGetProperty("html", out var h) ? h.GetString() : null) ?? string.Empty;
            string fileName = (data.TryGetProperty("fileName", out var fn) ? fn.GetString() : null) ?? string.Empty;
            string reportName = (data.TryGetProperty("reportName", out var rn) ? rn.GetString() : null) ?? "report";
            bool openAfterSave = (data.TryGetProperty("openAfterSave", out var o) && o.ValueKind == System.Text.Json.JsonValueKind.True) ||
                                 (data.TryGetProperty("openAfterSave", out o) && o.ValueKind == System.Text.Json.JsonValueKind.False ? false : true);

            if (string.IsNullOrWhiteSpace(html))
            {
                SendMessageToWebView(new { action = "error", message = "导出失败：HTML 为空" });
                return;
            }

            // exe 目录
            var exeDir = AppContext.BaseDirectory;
            var reportDir = Path.Combine(exeDir, "reports");
            Directory.CreateDirectory(reportDir);

            if (string.IsNullOrWhiteSpace(fileName))
            {
                fileName = $"{reportName}-{DateTime.Now:yyyyMMdd-HHmmss}.html";
            }
            if (!fileName.EndsWith(".html", StringComparison.OrdinalIgnoreCase))
            {
                fileName += ".html";
            }

            // 文件名净化（避免非法字符）
            foreach (var c in Path.GetInvalidFileNameChars())
                fileName = fileName.Replace(c, '_');

            var fullPath = Path.Combine(reportDir, fileName);
            File.WriteAllText(fullPath, html, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));

            SendMessageToWebView(new { action = "htmlReportSaved", filePath = fullPath });

            if (openAfterSave)
            {
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = fullPath,
                        UseShellExecute = true
                    });
                }
                catch
                {
                    // ignore open errors
                }
            }
        }
        catch (Exception ex)
        {
            WriteErrorLog("导出HTML报告失败", ex);
            SendMessageToWebView(new { action = "error", message = $"导出失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void SetClipboard(System.Text.Json.JsonElement data)
    {
        try
        {
            string text = (data.TryGetProperty("text", out var t) ? t.GetString() : null) ?? string.Empty;
            string context = (data.TryGetProperty("context", out var c) ? c.GetString() : null) ?? string.Empty;

            if (string.IsNullOrEmpty(text))
            {
                SendMessageToWebView(new { action = "clipboardSet", success = false, context });
                return;
            }

            Clipboard.SetText(text);
            SendMessageToWebView(new { action = "clipboardSet", success = true, context });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "clipboardSet", success = false, context = "unknown", message = ex.Message });
        }
    }

    private void OpenNativeFile()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(_currentFilePath) || !File.Exists(_currentFilePath))
            {
                SendMessageToWebView(new { action = "error", message = "原生文件不存在或未选择" });
                return;
            }
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = _currentFilePath,
                UseShellExecute = true
            };
            System.Diagnostics.Process.Start(psi);
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"打开原生文件失败: {ex.Message}" });
        }
    }

    private void DownloadImportSuccess()
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }

            // 默认导出当前主表（即“引入成功数据”）
            using var sfd = new SaveFileDialog();
            sfd.Title = "下载引入成功（Excel .xlsx - 全列文本）";
            sfd.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
            sfd.FileName = $"import-success-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx";
            if (sfd.ShowDialog() != DialogResult.OK) return;

            var main = MainTableNameOrDefault();
            ExportSqlToXlsx($"SELECT * FROM [{main}]", sfd.FileName, forceAllText: true);
        }
        catch (Exception ex)
        {
            WriteErrorLog("下载引入成功失败", ex);
            SendMessageToWebView(new { action = "error", message = $"下载引入成功失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void DownloadImportFail()
    {
        try
        {
            if (_lastConversionStats == null || _lastConversionStats.Count == 0)
            {
                SendMessageToWebView(new { action = "error", message = "暂无引入失败统计（请先完成一次引入）" });
                return;
            }

            using var sfd = new SaveFileDialog();
            sfd.Title = "下载引入失败（转换失败统计）";
            sfd.Filter = "Excel 文件 (*.xlsx)|*.xlsx|CSV 文件 (*.csv)|*.csv|所有文件 (*.*)|*.*";
            sfd.FileName = $"import-fail-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx";
            if (sfd.ShowDialog() != DialogResult.OK) return;

            var ext = Path.GetExtension(sfd.FileName).ToLowerInvariant();
            if (ext == ".csv")
            {
                using var writer = new StreamWriter(sfd.FileName, false, new UTF8Encoding(true));
                writer.WriteLine("列名,非空数,成功数,失败数,成功率");
                foreach (var s in _lastConversionStats)
                {
                    long nonEmpty = Convert.ToInt64(s.nonEmptyCount);
                    long ok = Convert.ToInt64(s.okCount);
                    long fail = Math.Max(0, nonEmpty - ok);
                    var rate = nonEmpty > 0 ? (ok * 100.0 / nonEmpty) : 100.0;
                    var colName = Convert.ToString(s.GetType().GetProperty("columnName")?.GetValue(s)) ?? "";
                    writer.WriteLine($"{CsvEscape(colName)},{nonEmpty},{ok},{fail},{rate:F2}%");
                }
            }
            else
            {
                var wb = new ClosedXML.Excel.XLWorkbook();
                var ws = wb.Worksheets.Add("ImportFail");
                ws.Cell(1, 1).Value = "列名";
                ws.Cell(1, 2).Value = "非空数";
                ws.Cell(1, 3).Value = "成功数";
                ws.Cell(1, 4).Value = "失败数";
                ws.Cell(1, 5).Value = "成功率";
                ws.Range(1, 1, 1, 5).Style.Font.Bold = true;

                int row = 2;
                foreach (var s in _lastConversionStats)
                {
                    long nonEmpty = Convert.ToInt64(s.nonEmptyCount);
                    long ok = Convert.ToInt64(s.okCount);
                    long fail = Math.Max(0, nonEmpty - ok);
                    var rate = nonEmpty > 0 ? (ok * 100.0 / nonEmpty) : 100.0;

                    var colName = Convert.ToString(s.GetType().GetProperty("columnName")?.GetValue(s)) ?? "";
                    ws.Cell(row, 1).SetValue(colName);
                    ws.Cell(row, 2).Value = nonEmpty;
                    ws.Cell(row, 3).Value = ok;
                    ws.Cell(row, 4).Value = fail;
                    ws.Cell(row, 5).SetValue($"{rate:F2}%");
                    row++;
                }
                ws.Columns().AdjustToContents();
                wb.SaveAs(sfd.FileName);
            }
        }
        catch (Exception ex)
        {
            WriteErrorLog("下载引入失败失败", ex);
            SendMessageToWebView(new { action = "error", message = $"下载引入失败失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void BrowseDcOutputPath()
    {
        try
        {
            using var sfd = new SaveFileDialog();
            sfd.Title = "选择清洗后输出文件";
            sfd.Filter = "Excel 文件 (*.xlsx)|*.xlsx|CSV 文件 (*.csv)|*.csv|所有文件 (*.*)|*.*";
            sfd.FileName = $"cleansed-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx";
            if (sfd.ShowDialog() != DialogResult.OK) return;
            SendMessageToWebView(new { action = "dcOutputPathSelected", outputPath = sfd.FileName.Replace('\\', '/') });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"选择输出路径失败: {ex.Message}" });
        }
    }

    private void GetDcPreviewStats(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }
            if (!data.TryGetProperty("settings", out var settings))
            {
                SendMessageToWebView(new { action = "error", message = "清洗参数缺失" });
                return;
            }

            var sourceTable = MainTableNameOrDefault();
            if (settings.TryGetProperty("sourceTable", out var stn) && stn.ValueKind == System.Text.Json.JsonValueKind.String)
            {
                var v = stn.GetString();
                if (!string.IsNullOrWhiteSpace(v)) sourceTable = v!;
            }

            // 映射
            var fieldMappings = new List<(string Source, string Target, string DataType, string Rule)>();
            if (settings.TryGetProperty("fieldMapping", out var fmap) && fmap.ValueKind == System.Text.Json.JsonValueKind.Array)
            {
                foreach (var e in fmap.EnumerateArray())
                {
                    var src = e.TryGetProperty("sourceField", out var sf) ? (sf.GetString() ?? "") : "";
                    var tgt = e.TryGetProperty("targetField", out var tf) ? (tf.GetString() ?? "") : "";
                    var dt = e.TryGetProperty("dataType", out var dtt) ? (dtt.GetString() ?? "text") : "text";
                    var rule = e.TryGetProperty("transformRule", out var tr) ? (tr.GetString() ?? "") : "";
                    if (!string.IsNullOrWhiteSpace(src) && !string.IsNullOrWhiteSpace(tgt))
                        fieldMappings.Add((src, tgt, dt, rule));
                }
            }
            if (fieldMappings.Count == 0)
            {
                var schema = _sqliteManager.GetTableSchema(sourceTable);
                foreach (var c in schema)
                    fieldMappings.Add((c.ColumnName, c.ColumnName, "text", ""));
            }

            var cleansingRules = settings.TryGetProperty("cleansingRules", out var cr) ? cr : default;
            bool fillEmpty = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("fillEmpty", out var fe) && fe.GetBoolean();
            string fillMethod = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("fillMethod", out var fm) ? (fm.GetString() ?? "custom") : "custom";
            string fillValue = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("fillValue", out var fv) ? (fv.GetString() ?? "") : "";
            bool removeEmptyRows = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("removeEmptyRows", out var rer) && rer.GetBoolean();
            bool removeDuplicates = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("removeDuplicates", out var rd) && rd.GetBoolean();
            bool trimSpaces = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("trimSpaces", out var ts) && ts.GetBoolean();
            bool standardizeCase = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("standardizeCase", out var sc) && sc.GetBoolean();

            int totalRows = _sqliteManager.GetRowCount(sourceTable);

            long filledCount = 0;
            long trimmedCount = 0;
            if (fillEmpty)
            {
                foreach (var m in fieldMappings)
                {
                    var col = SqlIdent(m.Source);
                    var cnt = SqlScalarLong($"SELECT SUM(CASE WHEN {col} IS NULL OR TRIM(CAST({col} AS TEXT))='' THEN 1 ELSE 0 END) AS v FROM [{sourceTable}]");
                    filledCount += cnt;
                }
            }
            if (trimSpaces)
            {
                foreach (var m in fieldMappings)
                {
                    var col = SqlIdent(m.Source);
                    var cnt = SqlScalarLong($"SELECT SUM(CASE WHEN {col} IS NOT NULL AND CAST({col} AS TEXT)<>TRIM(CAST({col} AS TEXT)) THEN 1 ELSE 0 END) AS v FROM [{sourceTable}]");
                    trimmedCount += cnt;
                }
            }

            // 生成 selectSql（与清洗执行一致），用于估算“去重/去空行”后的行数
            var selectExprs = new List<string>();
            foreach (var m in fieldMappings)
            {
                string expr = SqlIdent(m.Source);
                if (string.Equals(m.Rule, "trim", StringComparison.OrdinalIgnoreCase)) expr = $"TRIM(CAST({expr} AS TEXT))";
                else if (string.Equals(m.Rule, "upper", StringComparison.OrdinalIgnoreCase)) expr = $"UPPER(CAST({expr} AS TEXT))";
                else if (string.Equals(m.Rule, "lower", StringComparison.OrdinalIgnoreCase)) expr = $"LOWER(CAST({expr} AS TEXT))";

                if (trimSpaces) expr = $"TRIM(CAST({expr} AS TEXT))";
                if (standardizeCase) expr = $"UPPER(CAST({expr} AS TEXT))";

                switch ((m.DataType ?? "text").ToLowerInvariant())
                {
                    case "number":
                        expr = $"CAST(NULLIF(CAST({expr} AS TEXT),'') AS REAL)";
                        break;
                    case "boolean":
                        expr = $"CASE WHEN LOWER(CAST({expr} AS TEXT)) IN ('1','true','yes','y','是') THEN 1 WHEN LOWER(CAST({expr} AS TEXT)) IN ('0','false','no','n','否') THEN 0 ELSE NULL END";
                        break;
                    case "date":
                        expr = $"DATE(CAST({expr} AS TEXT))";
                        break;
                    case "datetime":
                        expr = $"DATETIME(CAST({expr} AS TEXT))";
                        break;
                    default:
                        expr = $"CAST({expr} AS TEXT)";
                        break;
                }

                if (fillEmpty)
                {
                    if (!string.Equals(fillMethod, "custom", StringComparison.OrdinalIgnoreCase))
                        fillMethod = "custom";
                    if (string.Equals(fillMethod, "custom", StringComparison.OrdinalIgnoreCase))
                    {
                        if (string.Equals(m.DataType, "number", StringComparison.OrdinalIgnoreCase) && double.TryParse(fillValue, out var dv))
                            expr = $"COALESCE({expr}, {dv.ToString(System.Globalization.CultureInfo.InvariantCulture)})";
                        else
                            expr = $"COALESCE(NULLIF({expr},''), {SqlValue(fillValue)})";
                    }
                }

                selectExprs.Add($"{expr} AS {SqlIdent(m.Target)}");
            }

            string distinct = removeDuplicates ? "DISTINCT " : "";
            string where = "";
            if (removeEmptyRows)
            {
                var conds = fieldMappings.Select(m => $"COALESCE(TRIM(CAST({SqlIdent(m.Source)} AS TEXT)),'')<>''").ToArray();
                where = "WHERE " + string.Join(" OR ", conds);
            }

            var selectSql = $"SELECT {distinct}{string.Join(", ", selectExprs)} FROM [{sourceTable}] {where}".Trim();
            long afterRows = SqlScalarLong($"SELECT COUNT(1) AS v FROM ({selectSql})");
            long removedCount = Math.Max(0, totalRows - afterRows);

            SendMessageToWebView(new
            {
                action = "dcPreviewStats",
                stats = new
                {
                    totalRows,
                    filledCount,
                    trimmedCount,
                    removedCount
                }
            });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"预览统计失败: {ex.Message}" });
        }
    }

    private async void ExecuteBusinessRuleVerify(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }
            if (!data.TryGetProperty("settings", out var settings))
            {
                SendMessageToWebView(new { action = "error", message = "校验参数缺失" });
                return;
            }

            var sw = System.Diagnostics.Stopwatch.StartNew();

            var sourceTable = MainTableNameOrDefault();
            if (settings.TryGetProperty("sourceTable", out var stn) && stn.ValueKind == System.Text.Json.JsonValueKind.String)
            {
                var v = stn.GetString();
                if (!string.IsNullOrWhiteSpace(v)) sourceTable = v!;
            }

            string verifySql = settings.TryGetProperty("verifySql", out var vs) && vs.ValueKind == System.Text.Json.JsonValueKind.String
                ? (vs.GetString() ?? "")
                : "";
            if (string.IsNullOrWhiteSpace(verifySql))
            {
                SendMessageToWebView(new { action = "error", message = "校验SQL为空（请先生成或配置规则）" });
                return;
            }

            // 输出选项（四选一）
            var outputOptions = settings.TryGetProperty("outputOptions", out var oo) ? oo : default;
            string outputTarget = outputOptions.ValueKind != System.Text.Json.JsonValueKind.Undefined && outputOptions.TryGetProperty("target", out var ot) ? (ot.GetString() ?? "file_new") : "file_new";
            string format = outputOptions.ValueKind != System.Text.Json.JsonValueKind.Undefined && outputOptions.TryGetProperty("format", out var fmt) ? (fmt.GetString() ?? "xlsx") : "xlsx";
            string outputPath = outputOptions.ValueKind != System.Text.Json.JsonValueKind.Undefined && outputOptions.TryGetProperty("outputPath", out var op) ? (op.GetString() ?? "") : "";
            string dbCopyName = outputOptions.ValueKind != System.Text.Json.JsonValueKind.Undefined && outputOptions.TryGetProperty("dbCopyTableName", out var dn) ? (dn.GetString() ?? "") : "";
            bool generateReport = outputOptions.ValueKind != System.Text.Json.JsonValueKind.Undefined && outputOptions.TryGetProperty("generateReport", out var gr) && gr.ValueKind == System.Text.Json.JsonValueKind.True;

            bool needExportFile = string.Equals(outputTarget, "file_new", StringComparison.OrdinalIgnoreCase)
                || string.Equals(outputTarget, "file_overwrite", StringComparison.OrdinalIgnoreCase);

            if (string.Equals(outputTarget, "file_overwrite", StringComparison.OrdinalIgnoreCase))
            {
                outputPath = _currentFilePath ?? outputPath;
            }
            if (needExportFile && string.IsNullOrWhiteSpace(outputPath))
            {
                SendMessageToWebView(new { action = "error", message = "输出路径为空" });
                return;
            }

            string targetTable = "Verify_Result";
            if (string.Equals(outputTarget, "db_copy", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(dbCopyName))
                targetTable = dbCopyName.Trim();

            SendMessageToWebView(new { action = "dcProgressUpdate", percent = 10, message = "准备校验..." });

            // 执行校验：落库到结果表（覆盖/副本）
            SendMessageToWebView(new { action = "dcProgressUpdate", percent = 40, message = "执行校验（生成校验结果表）..." });
            await Task.Run(() =>
            {
                _sqliteManager.Execute($"DROP TABLE IF EXISTS [{targetTable}];");
                _sqliteManager.Execute($"CREATE TABLE [{targetTable}] AS {verifySql};");
            });

            var violations = _sqliteManager.GetRowCount(targetTable);

            // 文件输出：导出异常记录
            if (needExportFile)
            {
                SendMessageToWebView(new { action = "dcProgressUpdate", percent = 75, message = "导出校验结果..." });
                var exportSql = $"SELECT * FROM [{targetTable}]";
                if (string.Equals(format, "csv", StringComparison.OrdinalIgnoreCase) || outputPath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                    ExportSqlToCsv(exportSql, outputPath, protectForExcel: true);
                else
                    ExportSqlToXlsx(exportSql, outputPath, forceAllText: true);
                _lastCleansedFilePath = outputPath;
            }

            // 生成报告（HTML）
            _lastCleansingReportPath = null;
            string reportPath = "";
            if (generateReport)
            {
                try
                {
                    string Html(string s) => (s ?? "").Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
                    var baseDir = (needExportFile && !string.IsNullOrWhiteSpace(outputPath))
                        ? (Path.GetDirectoryName(outputPath) ?? Path.GetTempPath())
                        : Path.GetTempPath();
                    reportPath = Path.Combine(baseDir, $"verify-report-{DateTime.Now:yyyyMMdd-HHmmss}.html");

                    var sampleRows = _sqliteManager.Query($"SELECT * FROM [{targetTable}] LIMIT 50") ?? new List<Dictionary<string, object?>>();
                    var cols = sampleRows.Count > 0 ? sampleRows[0].Keys.ToList() : new List<string>();

                    string RenderTable(List<string> columns, List<Dictionary<string, object?>> rows)
                    {
                        if (columns.Count == 0) return "<div style='color:#6c757d;'>（无样例数据）</div>";
                        var sb = new StringBuilder();
                        sb.Append("<div style='overflow:auto; border:1px solid #e1e4e8; border-radius:6px;'>");
                        sb.Append("<table style='border-collapse:collapse; width:100%; font-size:9pt;'>");
                        sb.Append("<thead style='background:#0d1b2a; color:#fff; position:sticky; top:0;'><tr>");
                        foreach (var c in columns) sb.Append("<th style='padding:8px 10px; border-bottom:1px solid #223; text-align:left; white-space:nowrap;'>").Append(Html(c)).Append("</th>");
                        sb.Append("</tr></thead><tbody>");
                        foreach (var r in rows)
                        {
                            sb.Append("<tr>");
                            foreach (var c in columns)
                            {
                                r.TryGetValue(c, out var v);
                                sb.Append("<td style='padding:8px 10px; border-bottom:1px solid #f0f0f0; white-space:nowrap;'>")
                                  .Append(Html(v?.ToString() ?? "")).Append("</td>");
                            }
                            sb.Append("</tr>");
                        }
                        sb.Append("</tbody></table></div>");
                        return sb.ToString();
                    }

                    string outputDesc = outputTarget switch
                    {
                        "db_copy" => $"生成数据库表副本：{Html(targetTable)}",
                        "file_overwrite" => $"覆盖源文件：{Html(outputPath)}",
                        "db_overwrite" => $"覆盖校验结果表：{Html(targetTable)}",
                        _ => $"生成新文件：{Html(outputPath)}"
                    };

                    var html = $@"
<!doctype html>
<html lang='zh-CN'>
<head>
  <meta charset='utf-8'>
  <meta name='viewport' content='width=device-width, initial-scale=1'>
  <title>业务规则验证报告</title>
  <style>
    body{{font-family:Segoe UI,Microsoft YaHei,Arial,sans-serif; margin:24px; color:#212529;}}
    .h1{{font-size:14pt; font-weight:700; margin:0 0 12px;}}
    .sec{{margin:14px 0; padding:12px 14px; border:1px solid #e1e4e8; border-radius:8px; background:#fff;}}
    .kvs{{display:grid; grid-template-columns: 160px 1fr; gap:6px 12px; font-size:9pt;}}
    .k{{color:#6c757d;}}
    pre{{background:#0b1320; color:#e6edf3; padding:12px; border-radius:8px; overflow:auto; font-size:9pt;}}
  </style>
</head>
<body>
  <div class='h1'>业务规则验证报告</div>
  <div style='color:#6c757d; font-size:9pt;'>生成时间：{DateTime.Now:yyyy-MM-dd HH:mm:ss}</div>

  <div class='sec'>
    <div style='font-weight:700; margin-bottom:8px;'>一、基本信息</div>
    <div class='kvs'>
      <div class='k'>源表（SQLite）</div><div>{Html(sourceTable)}</div>
      <div class='k'>结果表（SQLite）</div><div>{Html(targetTable)}</div>
      <div class='k'>输出方式</div><div>{outputDesc}</div>
      <div class='k'>异常记录数</div><div>{violations:N0}</div>
      <div class='k'>耗时</div><div>{Math.Round(sw.Elapsed.TotalSeconds,2)} 秒</div>
    </div>
  </div>

  <div class='sec'>
    <div style='font-weight:700; margin-bottom:8px;'>二、异常记录样例（前 50 行）</div>
    {RenderTable(cols, sampleRows)}
  </div>

  <div class='sec'>
    <div style='font-weight:700; margin-bottom:8px;'>三、可执行 SQL 脚本</div>
    <pre>{Html($"DROP TABLE IF EXISTS [{targetTable}];\\nCREATE TABLE [{targetTable}] AS {verifySql};")}</pre>
  </div>
</body>
</html>";

                    File.WriteAllText(reportPath, html, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
                    _lastCleansingReportPath = reportPath;
                }
                catch (Exception exr)
                {
                    WriteErrorLog("生成业务规则验证报告失败", exr);
                    _lastCleansingReportPath = null;
                }
            }

            sw.Stop();
            SendMessageToWebView(new { action = "dcProgressUpdate", percent = 100, message = "完成！" });
            SendMessageToWebView(new
            {
                action = "dcCleansingComplete",
                results = new
                {
                    fileName = Path.GetFileName(_currentFilePath ?? string.Empty),
                    totalRows = violations,
                    filledCount = 0,
                    removedCount = 0,
                    trimmedCount = 0,
                    processTime = Math.Round(sw.Elapsed.TotalSeconds, 2),
                    outputPath = (needExportFile ? outputPath.Replace('\\', '/') : ""),
                    outputTable = targetTable,
                    reportPath = (string.IsNullOrWhiteSpace(reportPath) ? "" : reportPath.Replace('\\', '/'))
                }
            });
        }
        catch (Exception ex)
        {
            WriteErrorLog("业务规则验证执行失败", ex);
            SendMessageToWebView(new { action = "error", message = $"业务规则验证执行失败: {ex.Message}" });
        }
    }

    private async void ExecuteDataCleansing(System.Text.Json.JsonElement data)
    {
        var sw = Stopwatch.StartNew();
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }
            if (_queryEngine == null)
            {
                SendMessageToWebView(new { action = "error", message = "查询引擎未初始化" });
                return;
            }

            if (!data.TryGetProperty("settings", out var settings))
            {
                SendMessageToWebView(new { action = "error", message = "清洗参数缺失" });
                return;
            }

            // 仅支持 SQLite：默认清洗“当前主表”；支持前端指定 sourceTable/targetTable（用于多工作表/三分组清洗并存）
            var sourceTable = MainTableNameOrDefault();
            string targetTable = "Cleansed";
            if (settings.TryGetProperty("sourceTable", out var stn) && stn.ValueKind == System.Text.Json.JsonValueKind.String)
            {
                var v = stn.GetString();
                if (!string.IsNullOrWhiteSpace(v)) sourceTable = v!;
            }
            if (settings.TryGetProperty("targetTable", out var ttn) && ttn.ValueKind == System.Text.Json.JsonValueKind.String)
            {
                var v = ttn.GetString();
                if (!string.IsNullOrWhiteSpace(v)) targetTable = v!;
            }

            var fieldMappings = new List<(string Source, string Target, string DataType, string Rule)>();
            if (settings.TryGetProperty("fieldMapping", out var fmap) && fmap.ValueKind == System.Text.Json.JsonValueKind.Array)
            {
                foreach (var e in fmap.EnumerateArray())
                {
                    var src = e.TryGetProperty("sourceField", out var sf) ? (sf.GetString() ?? "") : "";
                    var tgt = e.TryGetProperty("targetField", out var tf) ? (tf.GetString() ?? "") : "";
                    var dt = e.TryGetProperty("dataType", out var dtt) ? (dtt.GetString() ?? "text") : "text";
                    var rule = e.TryGetProperty("transformRule", out var tr) ? (tr.GetString() ?? "") : "";
                    if (!string.IsNullOrWhiteSpace(src) && !string.IsNullOrWhiteSpace(tgt))
                        fieldMappings.Add((src, tgt, dt, rule));
                }
            }
            if (fieldMappings.Count == 0)
            {
                // 没选字段时默认全字段
                var schema = _sqliteManager.GetTableSchema(sourceTable);
                foreach (var c in schema)
                {
                    fieldMappings.Add((c.ColumnName, c.ColumnName, "text", ""));
                }
            }

            var cleansingRules = settings.TryGetProperty("cleansingRules", out var cr) ? cr : default;
            bool fillEmpty = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("fillEmpty", out var fe) && fe.GetBoolean();
            string fillMethod = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("fillMethod", out var fm) ? (fm.GetString() ?? "custom") : "custom";
            string fillValue = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("fillValue", out var fv) ? (fv.GetString() ?? "") : "";
            bool removeEmptyRows = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("removeEmptyRows", out var rer) && rer.GetBoolean();
            bool removeDuplicates = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("removeDuplicates", out var rd) && rd.GetBoolean();
            bool trimSpaces = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("trimSpaces", out var ts) && ts.GetBoolean();
            bool standardizeCase = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("standardizeCase", out var sc) && sc.GetBoolean();

            var outputOptions = settings.TryGetProperty("outputOptions", out var oo) ? oo : default;
            string outputTarget = outputOptions.ValueKind != System.Text.Json.JsonValueKind.Undefined && outputOptions.TryGetProperty("target", out var ot) ? (ot.GetString() ?? "file_new") : "file_new";
            string format = outputOptions.ValueKind != System.Text.Json.JsonValueKind.Undefined && outputOptions.TryGetProperty("format", out var fmt) ? (fmt.GetString() ?? "xlsx") : "xlsx";
            string outputPath = outputOptions.ValueKind != System.Text.Json.JsonValueKind.Undefined && outputOptions.TryGetProperty("outputPath", out var op) ? (op.GetString() ?? "") : "";
            string dbCopyName = outputOptions.ValueKind != System.Text.Json.JsonValueKind.Undefined && outputOptions.TryGetProperty("dbCopyTableName", out var dn) ? (dn.GetString() ?? "") : "";
            bool generateReport = outputOptions.ValueKind != System.Text.Json.JsonValueKind.Undefined && outputOptions.TryGetProperty("generateReport", out var gr) && gr.ValueKind == System.Text.Json.JsonValueKind.True;

            bool needExportFile = string.Equals(outputTarget, "file_new", StringComparison.OrdinalIgnoreCase)
                || string.Equals(outputTarget, "file_overwrite", StringComparison.OrdinalIgnoreCase);

            if (string.Equals(outputTarget, "file_overwrite", StringComparison.OrdinalIgnoreCase))
            {
                // 覆盖：写回到原文件路径（风险由前端确认）
                outputPath = _currentFilePath ?? outputPath;
            }
            if (needExportFile && string.IsNullOrWhiteSpace(outputPath))
            {
                SendMessageToWebView(new { action = "error", message = "输出路径为空" });
                return;
            }

            if (string.Equals(outputTarget, "db_copy", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(dbCopyName))
            {
                targetTable = dbCopyName.Trim();
            }

            SendMessageToWebView(new { action = "dcProgressUpdate", percent = 5, message = "准备清洗..." });

            // 统计：源行数
            int beforeRows = _sqliteManager.GetRowCount(sourceTable);

            // 预估填充/去空格统计（简单估算）
            long filledCount = 0;
            long trimmedCount = 0;
            if (fillEmpty)
            {
                foreach (var m in fieldMappings)
                {
                    var col = SqlIdent(m.Source);
                    var cnt = SqlScalarLong($"SELECT SUM(CASE WHEN {col} IS NULL OR TRIM(CAST({col} AS TEXT))='' THEN 1 ELSE 0 END) AS v FROM [{sourceTable}]");
                    filledCount += cnt;
                }
            }
            if (trimSpaces)
            {
                foreach (var m in fieldMappings)
                {
                    var col = SqlIdent(m.Source);
                    var cnt = SqlScalarLong($"SELECT SUM(CASE WHEN {col} IS NOT NULL AND CAST({col} AS TEXT)<>TRIM(CAST({col} AS TEXT)) THEN 1 ELSE 0 END) AS v FROM [{sourceTable}]");
                    trimmedCount += cnt;
                }
            }

            SendMessageToWebView(new { action = "dcProgressUpdate", percent = 20, message = "生成清洗SQL..." });

            // 构造 SELECT 表达式
            var selectExprs = new List<string>();
            foreach (var m in fieldMappings)
            {
                string expr = SqlIdent(m.Source);
                // 字段级转换规则
                if (string.Equals(m.Rule, "trim", StringComparison.OrdinalIgnoreCase)) expr = $"TRIM(CAST({expr} AS TEXT))";
                else if (string.Equals(m.Rule, "upper", StringComparison.OrdinalIgnoreCase)) expr = $"UPPER(CAST({expr} AS TEXT))";
                else if (string.Equals(m.Rule, "lower", StringComparison.OrdinalIgnoreCase)) expr = $"LOWER(CAST({expr} AS TEXT))";

                // 全局规则
                if (trimSpaces) expr = $"TRIM(CAST({expr} AS TEXT))";
                if (standardizeCase) expr = $"UPPER(CAST({expr} AS TEXT))";

                // 类型转换
                switch ((m.DataType ?? "text").ToLowerInvariant())
                {
                    case "number":
                        expr = $"CAST(NULLIF(CAST({expr} AS TEXT),'') AS REAL)";
                        break;
                    case "boolean":
                        expr = $"CASE WHEN LOWER(CAST({expr} AS TEXT)) IN ('1','true','yes','y','是') THEN 1 WHEN LOWER(CAST({expr} AS TEXT)) IN ('0','false','no','n','否') THEN 0 ELSE NULL END";
                        break;
                    case "date":
                        expr = $"DATE(CAST({expr} AS TEXT))";
                        break;
                    case "datetime":
                        expr = $"DATETIME(CAST({expr} AS TEXT))";
                        break;
                    default:
                        expr = $"CAST({expr} AS TEXT)";
                        break;
                }

                // 空值填充（当前优先支持 custom；其它方式先回退为 custom）
                if (fillEmpty)
                {
                    if (!string.Equals(fillMethod, "custom", StringComparison.OrdinalIgnoreCase))
                    {
                        // TODO：forward/backward/median/mode/mean 可逐步增强
                        fillMethod = "custom";
                    }

                    if (string.Equals(fillMethod, "custom", StringComparison.OrdinalIgnoreCase))
                    {
                        if (string.Equals(m.DataType, "number", StringComparison.OrdinalIgnoreCase) && double.TryParse(fillValue, out var dv))
                            expr = $"COALESCE({expr}, {dv.ToString(System.Globalization.CultureInfo.InvariantCulture)})";
                        else
                            expr = $"COALESCE(NULLIF({expr},''), {SqlValue(fillValue)})";
                    }
                }

                selectExprs.Add($"{expr} AS {SqlIdent(m.Target)}");
            }

            string distinct = removeDuplicates ? "DISTINCT " : "";
            string where = "";
            if (removeEmptyRows)
            {
                // 至少一个字段非空
                var conds = fieldMappings.Select(m => $"COALESCE(TRIM(CAST({SqlIdent(m.Source)} AS TEXT)),'')<>''").ToArray();
                where = "WHERE " + string.Join(" OR ", conds);
            }

            var selectSql = $"SELECT {distinct}{string.Join(", ", selectExprs)} FROM [{sourceTable}] {where}";

            SendMessageToWebView(new { action = "dcProgressUpdate", percent = 45, message = "执行清洗（生成清洗后表）..." });

            // 输出到数据库：覆盖源表 / 生成副本
            if (string.Equals(outputTarget, "db_overwrite", StringComparison.OrdinalIgnoreCase))
            {
                // 安全覆盖：先生成临时表，再替换源表
                var tmp = $"{sourceTable}__tmp__{DateTime.Now:yyyyMMddHHmmssfff}";
                await Task.Run(() =>
                {
                    _sqliteManager.Execute($"DROP TABLE IF EXISTS [{tmp}];");
                    _sqliteManager.Execute($"CREATE TABLE [{tmp}] AS {selectSql};");
                    _sqliteManager.Execute($"DROP TABLE IF EXISTS [{sourceTable}];");
                    _sqliteManager.Execute($"ALTER TABLE [{tmp}] RENAME TO [{sourceTable}];");
                });
                targetTable = sourceTable;
            }
            else
            {
                await Task.Run(() =>
                {
                    _sqliteManager.Execute($"DROP TABLE IF EXISTS [{targetTable}];");
                    _sqliteManager.Execute($"CREATE TABLE [{targetTable}] AS {selectSql};");
                });
            }

            int afterRows = _sqliteManager.GetRowCount(targetTable);
            long removedCount = Math.Max(0, beforeRows - afterRows);

            // 文件输出：导出清洗结果
            if (needExportFile)
            {
                SendMessageToWebView(new { action = "dcProgressUpdate", percent = 75, message = "导出清洗结果..." });

                var exportSql = $"SELECT * FROM [{targetTable}]";
                if (string.Equals(format, "csv", StringComparison.OrdinalIgnoreCase) || outputPath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    ExportSqlToCsv(exportSql, outputPath, protectForExcel: true);
                }
                else
                {
                    ExportSqlToXlsx(exportSql, outputPath, forceAllText: true);
                }
                _lastCleansedFilePath = outputPath;
            }

            sw.Stop();
            // 生成清洗报告（HTML，单文件）
            _lastCleansingReportPath = null;
            string reportPath = "";
            if (generateReport)
            {
                try
                {
                    string Html(string s) => (s ?? "").Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");

                    // 报告保存位置：优先与导出文件同目录，否则写到临时目录
                    var baseDir = (needExportFile && !string.IsNullOrWhiteSpace(outputPath))
                        ? (Path.GetDirectoryName(outputPath) ?? Path.GetTempPath())
                        : Path.GetTempPath();
                    reportPath = Path.Combine(baseDir, $"cleansing-report-{DateTime.Now:yyyyMMdd-HHmmss}.html");

                    // SQL脚本（可执行）
                    string ddlSql;
                    if (string.Equals(outputTarget, "db_overwrite", StringComparison.OrdinalIgnoreCase))
                    {
                        var tmp = $"{sourceTable}__tmp__{DateTime.Now:yyyyMMddHHmmssfff}";
                        ddlSql =
                            $"DROP TABLE IF EXISTS [{tmp}];\n" +
                            $"CREATE TABLE [{tmp}] AS {selectSql};\n" +
                            $"DROP TABLE IF EXISTS [{sourceTable}];\n" +
                            $"ALTER TABLE [{tmp}] RENAME TO [{sourceTable}];";
                    }
                    else
                    {
                        ddlSql = $"DROP TABLE IF EXISTS [{targetTable}];\nCREATE TABLE [{targetTable}] AS {selectSql};";
                    }

                    // 采样数据：清洗后表前20行
                    var sampleRows = _sqliteManager.Query($"SELECT * FROM [{targetTable}] LIMIT 20") ?? new List<Dictionary<string, object?>>();
                    var cols = sampleRows.Count > 0 ? sampleRows[0].Keys.ToList() : new List<string>();

                    string RenderTable(List<string> columns, List<Dictionary<string, object?>> rows)
                    {
                        if (columns.Count == 0) return "<div style='color:#6c757d;'>（无样例数据）</div>";
                        var sb = new StringBuilder();
                        sb.Append("<div style='overflow:auto; border:1px solid #e1e4e8; border-radius:6px;'>");
                        sb.Append("<table style='border-collapse:collapse; width:100%; font-size:9pt;'>");
                        sb.Append("<thead style='background:#0d1b2a; color:#fff; position:sticky; top:0;'><tr>");
                        foreach (var c in columns) sb.Append("<th style='padding:8px 10px; border-bottom:1px solid #223; text-align:left; white-space:nowrap;'>").Append(Html(c)).Append("</th>");
                        sb.Append("</tr></thead><tbody>");
                        foreach (var r in rows)
                        {
                            sb.Append("<tr>");
                            foreach (var c in columns)
                            {
                                r.TryGetValue(c, out var v);
                                sb.Append("<td style='padding:8px 10px; border-bottom:1px solid #f0f0f0; white-space:nowrap;'>")
                                  .Append(Html(v?.ToString() ?? "")).Append("</td>");
                            }
                            sb.Append("</tr>");
                        }
                        sb.Append("</tbody></table></div>");
                        return sb.ToString();
                    }

                    // 字段映射配置
                    var mapCols = new List<string> { "源字段", "目标字段", "类型", "转换规则" };
                    var mapRows = fieldMappings.Select(m => new Dictionary<string, object?>
                    {
                        ["源字段"] = m.Source,
                        ["目标字段"] = m.Target,
                        ["类型"] = m.DataType,
                        ["转换规则"] = string.IsNullOrWhiteSpace(m.Rule) ? "（无）" : m.Rule
                    }).ToList();

                    // 统计与输出说明
                    string outputDesc = outputTarget switch
                    {
                        "db_overwrite" => $"覆盖当前数据库表：{Html(sourceTable)}",
                        "db_copy" => $"生成数据库表副本：{Html(targetTable)}",
                        "file_overwrite" => $"覆盖源文件：{Html(outputPath)}",
                        _ => $"生成新文件：{Html(outputPath)}"
                    };

                    var suggestions = new List<string>
                    {
                        "【建议】空值填充目前优先支持“自定义值”，如需均值/中位数/众数/前后填充，可在后续版本增强。",
                        "【建议】字段级转换规则可扩展：日期格式标准化、正则清洗、映射表替换等。",
                        "【建议】执行后建议进行差异比对（Excel ↔ SQLite 清洗后表）确保结果符合预期。"
                    };

                    var html = $@"
<!doctype html>
<html lang='zh-CN'>
<head>
  <meta charset='utf-8'>
  <meta name='viewport' content='width=device-width, initial-scale=1'>
  <title>数据清洗报告</title>
  <style>
    body{{font-family:Segoe UI,Microsoft YaHei,Arial,sans-serif; margin:24px; color:#212529;}}
    .h1{{font-size:14pt; font-weight:700; margin:0 0 12px;}}
    .sec{{margin:14px 0; padding:12px 14px; border:1px solid #e1e4e8; border-radius:8px; background:#fff;}}
    .kvs{{display:grid; grid-template-columns: 140px 1fr; gap:6px 12px; font-size:9pt;}}
    .k{{color:#6c757d;}}
    pre{{background:#0b1320; color:#e6edf3; padding:12px; border-radius:8px; overflow:auto; font-size:9pt;}}
    .tag{{display:inline-block; padding:2px 8px; border-radius:999px; background:#f6ffed; border:1px solid #b7eb8f; color:#2f7d32; font-size:8.5pt;}}
  </style>
</head>
<body>
  <div class='h1'>数据清洗报告 <span class='tag'>ExcelSQLite</span></div>
  <div style='color:#6c757d; font-size:9pt;'>生成时间：{DateTime.Now:yyyy-MM-dd HH:mm:ss}</div>

  <div class='sec'>
    <div style='font-weight:700; margin-bottom:8px;'>一、基本信息</div>
    <div class='kvs'>
      <div class='k'>源表（SQLite）</div><div>{Html(sourceTable)}</div>
      <div class='k'>清洗后表（SQLite）</div><div>{Html(targetTable)}</div>
      <div class='k'>输出方式</div><div>{outputDesc}</div>
      <div class='k'>耗时</div><div>{Math.Round(sw.Elapsed.TotalSeconds,2)} 秒</div>
    </div>
  </div>

  <div class='sec'>
    <div style='font-weight:700; margin-bottom:8px;'>二、转换映射配置</div>
    {RenderTable(mapCols, mapRows)}
  </div>

  <div class='sec'>
    <div style='font-weight:700; margin-bottom:8px;'>三、转换数据样例（清洗后前 20 行）</div>
    {RenderTable(cols, sampleRows)}
  </div>

  <div class='sec'>
    <div style='font-weight:700; margin-bottom:8px;'>四、记录统计 / 清洗统计</div>
    <div class='kvs'>
      <div class='k'>源表行数</div><div>{beforeRows:N0}</div>
      <div class='k'>清洗后行数</div><div>{afterRows:N0}</div>
      <div class='k'>删除重复/空行（估算）</div><div>{removedCount:N0}</div>
      <div class='k'>填充空值（估算）</div><div>{filledCount:N0}</div>
      <div class='k'>清理空格（估算）</div><div>{trimmedCount:N0}</div>
    </div>
  </div>

  <div class='sec'>
    <div style='font-weight:700; margin-bottom:8px;'>五、可执行 SQL 脚本</div>
    <pre>{Html(ddlSql)}</pre>
  </div>

  <div class='sec'>
    <div style='font-weight:700; margin-bottom:8px;'>六、剩余事项 / 改进建议</div>
    <ul style='margin:0; padding-left:18px; font-size:9pt;'>
      {string.Join("", suggestions.Select(x => "<li>"+Html(x)+"</li>"))}
    </ul>
  </div>
</body>
</html>";

                    File.WriteAllText(reportPath, html, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
                    _lastCleansingReportPath = reportPath;
                }
                catch (Exception exr)
                {
                    WriteErrorLog("生成清洗报告失败", exr);
                    _lastCleansingReportPath = null;
                }
            }

            SendMessageToWebView(new
            {
                action = "dcCleansingComplete",
                results = new
                {
                    fileName = Path.GetFileName(_currentFilePath ?? string.Empty),
                    totalRows = afterRows,
                    filledCount,
                    removedCount,
                    trimmedCount,
                    processTime = Math.Round(sw.Elapsed.TotalSeconds, 2),
                    outputPath = (needExportFile ? outputPath.Replace('\\', '/') : ""),
                    outputTable = targetTable,
                    reportPath = (string.IsNullOrWhiteSpace(reportPath) ? "" : reportPath.Replace('\\', '/'))
                }
            });

            // 刷新表列表（差异比对“清洗后表”下拉依赖）
            GetTableList();
        }
        catch (Exception ex)
        {
            WriteErrorLog("数据清洗失败", ex);
            SendMessageToWebView(new { action = "error", message = $"数据清洗失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void GenerateDataCleansingSql(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }
            if (!data.TryGetProperty("settings", out var settings))
            {
                SendMessageToWebView(new { action = "error", message = "清洗参数缺失" });
                return;
            }

            // source/target
            var sourceTable = MainTableNameOrDefault();
            string targetTable = "Cleansed";
            if (settings.TryGetProperty("sourceTable", out var stn) && stn.ValueKind == System.Text.Json.JsonValueKind.String)
            {
                var v = stn.GetString();
                if (!string.IsNullOrWhiteSpace(v)) sourceTable = v!;
            }
            if (settings.TryGetProperty("targetTable", out var ttn) && ttn.ValueKind == System.Text.Json.JsonValueKind.String)
            {
                var v = ttn.GetString();
                if (!string.IsNullOrWhiteSpace(v)) targetTable = v!;
            }

            var fieldMappings = new List<(string Source, string Target, string DataType, string Rule)>();
            if (settings.TryGetProperty("fieldMapping", out var fmap) && fmap.ValueKind == System.Text.Json.JsonValueKind.Array)
            {
                foreach (var e in fmap.EnumerateArray())
                {
                    var src = e.TryGetProperty("sourceField", out var sf) ? (sf.GetString() ?? "") : "";
                    var tgt = e.TryGetProperty("targetField", out var tf) ? (tf.GetString() ?? "") : "";
                    var dt = e.TryGetProperty("dataType", out var dtt) ? (dtt.GetString() ?? "text") : "text";
                    var rule = e.TryGetProperty("transformRule", out var tr) ? (tr.GetString() ?? "") : "";
                    if (!string.IsNullOrWhiteSpace(src) && !string.IsNullOrWhiteSpace(tgt))
                        fieldMappings.Add((src, tgt, dt, rule));
                }
            }
            if (fieldMappings.Count == 0)
            {
                var schema = _sqliteManager.GetTableSchema(sourceTable);
                foreach (var c in schema)
                {
                    fieldMappings.Add((c.ColumnName, c.ColumnName, "text", ""));
                }
            }

            var cleansingRules = settings.TryGetProperty("cleansingRules", out var cr) ? cr : default;
            bool fillEmpty = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("fillEmpty", out var fe) && fe.GetBoolean();
            string fillMethod = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("fillMethod", out var fm) ? (fm.GetString() ?? "custom") : "custom";
            string fillValue = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("fillValue", out var fv) ? (fv.GetString() ?? "") : "";
            bool removeEmptyRows = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("removeEmptyRows", out var rer) && rer.GetBoolean();
            bool removeDuplicates = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("removeDuplicates", out var rd) && rd.GetBoolean();
            bool trimSpaces = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("trimSpaces", out var ts) && ts.GetBoolean();
            bool standardizeCase = cleansingRules.ValueKind != System.Text.Json.JsonValueKind.Undefined && cleansingRules.TryGetProperty("standardizeCase", out var sc) && sc.GetBoolean();

            // 构造 SELECT 表达式（与 ExecuteDataCleansing 保持一致）
            var selectExprs = new List<string>();
            foreach (var m in fieldMappings)
            {
                string expr = SqlIdent(m.Source);
                if (string.Equals(m.Rule, "trim", StringComparison.OrdinalIgnoreCase)) expr = $"TRIM(CAST({expr} AS TEXT))";
                else if (string.Equals(m.Rule, "upper", StringComparison.OrdinalIgnoreCase)) expr = $"UPPER(CAST({expr} AS TEXT))";
                else if (string.Equals(m.Rule, "lower", StringComparison.OrdinalIgnoreCase)) expr = $"LOWER(CAST({expr} AS TEXT))";

                if (trimSpaces) expr = $"TRIM(CAST({expr} AS TEXT))";
                if (standardizeCase) expr = $"UPPER(CAST({expr} AS TEXT))";

                switch ((m.DataType ?? "text").ToLowerInvariant())
                {
                    case "number":
                        expr = $"CAST(NULLIF(CAST({expr} AS TEXT),'') AS REAL)";
                        break;
                    case "boolean":
                        expr = $"CASE WHEN LOWER(CAST({expr} AS TEXT)) IN ('1','true','yes','y','是') THEN 1 WHEN LOWER(CAST({expr} AS TEXT)) IN ('0','false','no','n','否') THEN 0 ELSE NULL END";
                        break;
                    case "date":
                        expr = $"DATE(CAST({expr} AS TEXT))";
                        break;
                    case "datetime":
                        expr = $"DATETIME(CAST({expr} AS TEXT))";
                        break;
                    default:
                        expr = $"CAST({expr} AS TEXT)";
                        break;
                }

                if (fillEmpty)
                {
                    if (!string.Equals(fillMethod, "custom", StringComparison.OrdinalIgnoreCase))
                        fillMethod = "custom";
                    if (string.Equals(fillMethod, "custom", StringComparison.OrdinalIgnoreCase))
                    {
                        if (string.Equals(m.DataType, "number", StringComparison.OrdinalIgnoreCase) && double.TryParse(fillValue, out var dv))
                            expr = $"COALESCE({expr}, {dv.ToString(System.Globalization.CultureInfo.InvariantCulture)})";
                        else
                            expr = $"COALESCE(NULLIF({expr},''), {SqlValue(fillValue)})";
                    }
                }

                selectExprs.Add($"{expr} AS {SqlIdent(m.Target)}");
            }

            string distinct = removeDuplicates ? "DISTINCT " : "";
            string where = "";
            if (removeEmptyRows)
            {
                var conds = fieldMappings.Select(m => $"COALESCE(TRIM(CAST({SqlIdent(m.Source)} AS TEXT)),'')<>''").ToArray();
                where = "WHERE " + string.Join(" OR ", conds);
            }

            var selectSql = $"SELECT {distinct}{string.Join(", ", selectExprs)} FROM [{sourceTable}] {where}".Trim();
            // 输出目标（用于提示 SQL）：覆盖源表/生成副本/文件输出
            var outputOptions = settings.TryGetProperty("outputOptions", out var oo) ? oo : default;
            string outputTarget = outputOptions.ValueKind != System.Text.Json.JsonValueKind.Undefined && outputOptions.TryGetProperty("target", out var ot) ? (ot.GetString() ?? "file_new") : "file_new";
            string dbCopyName = outputOptions.ValueKind != System.Text.Json.JsonValueKind.Undefined && outputOptions.TryGetProperty("dbCopyTableName", out var dn) ? (dn.GetString() ?? "") : "";
            if (string.Equals(outputTarget, "db_overwrite", StringComparison.OrdinalIgnoreCase))
            {
                targetTable = sourceTable;
            }
            else if (string.Equals(outputTarget, "db_copy", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(dbCopyName))
            {
                targetTable = dbCopyName.Trim();
            }

            string ddl;
            if (string.Equals(outputTarget, "db_overwrite", StringComparison.OrdinalIgnoreCase))
            {
                // 覆盖源表：用临时表替换（更安全）
                var tmp = $"{sourceTable}__tmp__{DateTime.Now:yyyyMMddHHmmssfff}";
                ddl =
                    $"DROP TABLE IF EXISTS [{tmp}];\n" +
                    $"CREATE TABLE [{tmp}] AS {selectSql};\n" +
                    $"DROP TABLE IF EXISTS [{sourceTable}];\n" +
                    $"ALTER TABLE [{tmp}] RENAME TO [{sourceTable}];";
            }
            else
            {
                ddl = $"DROP TABLE IF EXISTS [{targetTable}];\nCREATE TABLE [{targetTable}] AS {selectSql};";
            }

            SendMessageToWebView(new
            {
                action = "dcSqlGenerated",
                sourceTable,
                targetTable,
                selectSql,
                ddlSql = ddl
            });
        }
        catch (Exception ex)
        {
            WriteErrorLog("生成清洗SQL失败", ex);
            SendMessageToWebView(new { action = "error", message = $"生成清洗SQL失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private async void ExecuteDataCompare(System.Text.Json.JsonElement data)
    {
        var sw = Stopwatch.StartNew();
        string beforeTable = "__CompareBefore";
        try
        {
            if (_sqliteManager == null || _dataImporter == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite/导入器未初始化" });
                return;
            }
            if (string.IsNullOrWhiteSpace(_currentFilePath))
            {
                SendMessageToWebView(new { action = "error", message = "请先选择主表文件" });
                return;
            }
            if (!data.TryGetProperty("settings", out var settings))
            {
                SendMessageToWebView(new { action = "error", message = "比对参数缺失" });
                return;
            }

            string afterTable = settings.TryGetProperty("afterTable", out var at) ? (at.GetString() ?? "") : "";
            string beforeSheet = settings.TryGetProperty("beforeSheet", out var bs) ? (bs.GetString() ?? "") : "";
            if (string.IsNullOrWhiteSpace(afterTable) || string.IsNullOrWhiteSpace(beforeSheet))
            {
                SendMessageToWebView(new { action = "error", message = "请先选择清洗前工作表与清洗后表" });
                return;
            }

            // 临时导入清洗前（用 text 模式避免转换丢失）
            _sqliteManager.Execute($"DROP TABLE IF EXISTS [{beforeTable}];");
            var progress = new Progress<ImportProgress>(_ => { });
            var r = await _dataImporter.ImportWorksheetAsync(_currentFilePath, beforeSheet, beforeTable, "text", progress, CancellationToken.None);
            if (!r.Success) throw new InvalidOperationException(r.Message);

            var beforeSchema = _sqliteManager.GetTableSchema(beforeTable);
            var afterSchema = _sqliteManager.GetTableSchema(afterTable);
            var beforeCols = beforeSchema.Select(s => s.ColumnName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var afterCols = afterSchema.Select(s => s.ColumnName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var commonCols = beforeCols.Intersect(afterCols, StringComparer.OrdinalIgnoreCase).ToList();

            var fieldDiffs = new List<object>();
            foreach (var c in beforeCols.Except(afterCols, StringComparer.OrdinalIgnoreCase))
                fieldDiffs.Add(new { type = "缺失字段", field = c, before = "存在", after = "-" });
            foreach (var c in afterCols.Except(beforeCols, StringComparer.OrdinalIgnoreCase))
                fieldDiffs.Add(new { type = "新增字段", field = c, before = "-", after = "存在" });

            int beforeRows = _sqliteManager.GetRowCount(beforeTable);
            int afterRows = _sqliteManager.GetRowCount(afterTable);

            var dataDiffs = new List<object>();
            long dataDiffCount = 0;
            if (commonCols.Count > 0)
            {
                var colsSql = string.Join(", ", commonCols.Select(SqlIdent));
                var onlyBeforeCnt = SqlScalarLong($"SELECT COUNT(*) AS v FROM (SELECT {colsSql} FROM [{beforeTable}] EXCEPT SELECT {colsSql} FROM [{afterTable}]) t");
                var onlyAfterCnt = SqlScalarLong($"SELECT COUNT(*) AS v FROM (SELECT {colsSql} FROM [{afterTable}] EXCEPT SELECT {colsSql} FROM [{beforeTable}]) t");
                dataDiffCount = onlyBeforeCnt + onlyAfterCnt;

                // 抽样展示（用整行JSON字符串展示，避免复杂“主键对齐”）
                var sampleRows = _sqliteManager.Query($"SELECT * FROM (SELECT {colsSql} FROM [{beforeTable}] EXCEPT SELECT {colsSql} FROM [{afterTable}]) t LIMIT 5");
                int idx = 1;
                foreach (var row in sampleRows)
                {
                    dataDiffs.Add(new { rowNumber = idx++, fieldName = "*", beforeValue = System.Text.Json.JsonSerializer.Serialize(row), afterValue = "", diffType = "仅清洗前存在" });
                }
                var sampleRows2 = _sqliteManager.Query($"SELECT * FROM (SELECT {colsSql} FROM [{afterTable}] EXCEPT SELECT {colsSql} FROM [{beforeTable}]) t LIMIT 5");
                foreach (var row in sampleRows2)
                {
                    dataDiffs.Add(new { rowNumber = idx++, fieldName = "*", beforeValue = "", afterValue = System.Text.Json.JsonSerializer.Serialize(row), diffType = "仅清洗后存在" });
                }
            }

            sw.Stop();
            SendMessageToWebView(new
            {
                action = "compareComplete",
                results = new
                {
                    beforeSheet,
                    afterTable,
                    beforeRows,
                    afterRows,
                    fieldDiffCount = fieldDiffs.Count,
                    fieldDiffs,
                    dataDiffCount,
                    dataDiffs,
                    compareTime = Math.Round(sw.Elapsed.TotalSeconds, 2)
                }
            });
        }
        catch (Exception ex)
        {
            WriteErrorLog("差异比对失败", ex);
            SendMessageToWebView(new { action = "error", message = $"差异比对失败: {ex.Message}", hasErrorLog = true });
        }
        finally
        {
            try { _sqliteManager?.Execute($"DROP TABLE IF EXISTS [{beforeTable}];"); } catch { }
        }
    }

    private void OpenCleansedFile()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(_lastCleansedFilePath) || !File.Exists(_lastCleansedFilePath))
            {
                SendMessageToWebView(new { action = "error", message = "暂无可打开的清洗后文件" });
                return;
            }
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = _lastCleansedFilePath, UseShellExecute = true });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"打开清洗后文件失败: {ex.Message}" });
        }
    }

    private void ViewCleansingReport()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(_lastCleansingReportPath) || !File.Exists(_lastCleansingReportPath))
            {
                SendMessageToWebView(new { action = "error", message = "暂无可查看的清洗报告" });
                return;
            }
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = _lastCleansingReportPath, UseShellExecute = true });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"打开清洗报告失败: {ex.Message}" });
        }
    }

    private static string SqlIdent(string name)
    {
        var n = (name ?? "").Replace("]", "]]");
        return $"[{n}]";
    }

    private static string SqlValue(string? v)
    {
        var s = v ?? "";
        s = s.Replace("'", "''");
        return $"'{s}'";
    }

    private long SqlScalarLong(string sql)
    {
        if (_sqliteManager == null) return 0;
        var rows = _sqliteManager.Query(sql);
        if (rows == null || rows.Count == 0) return 0;
        var first = rows[0];
        if (first.Count == 0) return 0;
        // 优先取 key=v
        if (first.TryGetValue("v", out var v) && v != null) return Convert.ToInt64(v);
        // 否则取第一个字段
        var any = first.Values.FirstOrDefault();
        if (any == null) return 0;
        return Convert.ToInt64(any);
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

        // 全局样式（符合 UI/UX 规范：微软雅黑 9pt，单元格居中，合理行高）
        ws.Style.Font.FontName = "Microsoft YaHei";
        ws.Style.Font.FontSize = 9;
        ws.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
        ws.Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Center;

        // header
        for (int c = 0; c < colCount; c++)
        {
            ws.Cell(1, c + 1).Value = reader.GetName(c);
            ws.Cell(1, c + 1).Style.Font.Bold = true;
        }
        ws.Row(1).Height = 20;

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

        // 行高（按 9pt 合理默认）
        try
        {
            var used = ws.RangeUsed();
            if (used != null)
            {
                foreach (var r in used.Rows())
                {
                    // ClosedXML：IXLRangeRow 本身不提供 Height，需通过 worksheet row 设置
                    if (r.RowNumber() != 1) ws.Row(r.RowNumber()).Height = 18;
                }
            }
        }
        catch { }

        // 列宽：自适应 + 合理上限（避免超长文本拉爆）
        ws.Columns().AdjustToContents();
        try
        {
            foreach (var col in ws.ColumnsUsed())
            {
                if (col.Width < 8) col.Width = 8;
                if (col.Width > 60) col.Width = 60;
            }
        }
        catch { }
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

            // 按需求：提示用户确认，并告知将清空先前文件/数据库加载信息
            var msg = $"你确认要引入《{Path.GetFileName(record.FullPath)}》文件到SQL内存吗？\n\n本操作将清空先前的文件及数据库加载信息！";
            var ok = MessageBox.Show(msg, "确认重新引入", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (ok != DialogResult.OK) return;

            OpenMainFileInternal(record.FullPath, clearPrevious: true);
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"打开最近文件失败: {ex.Message}" });
        }
    }

    // ========= 设置保存/加载（Excel 同目录 INI，值为 Base64(JSON)） =========
    private static string GetSettingsIniPath(string excelPath)
    {
        var dir = Path.GetDirectoryName(excelPath) ?? AppDomain.CurrentDomain.BaseDirectory;
        var baseName = Path.GetFileNameWithoutExtension(excelPath);
        return Path.Combine(dir, $"{baseName}.excelsqlite.ini");
    }

    private void SaveSettings(System.Text.Json.JsonElement data)
    {
        try
        {
            var scope = (data.TryGetProperty("scope", out var sc) ? sc.GetString() : null) ?? "";
            var excelPath = (data.TryGetProperty("excelPath", out var ep) ? ep.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(scope) || string.IsNullOrWhiteSpace(excelPath))
            {
                SendMessageToWebView(new { action = "error", message = "保存设置失败：scope/excelPath 为空" });
                return;
            }
            if (!data.TryGetProperty("data", out var payload))
            {
                SendMessageToWebView(new { action = "error", message = "保存设置失败：data 为空" });
                return;
            }
            var rawJson = payload.GetRawText();
            var b64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(rawJson));
            var iniPath = GetSettingsIniPath(excelPath);
            var ini = ReadIni(iniPath);
            if (!ini.ContainsKey(scope)) ini[scope] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            ini[scope]["json"] = b64;
            ini[scope]["updatedAt"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            WriteIni(iniPath, ini);
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"保存设置失败: {ex.Message}" });
        }
    }

    private void LoadSettings(System.Text.Json.JsonElement data)
    {
        try
        {
            var scope = (data.TryGetProperty("scope", out var sc) ? sc.GetString() : null) ?? "";
            var excelPath = (data.TryGetProperty("excelPath", out var ep) ? ep.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(scope) || string.IsNullOrWhiteSpace(excelPath))
                return;
            var iniPath = GetSettingsIniPath(excelPath);
            var ini = ReadIni(iniPath);
            if (!ini.TryGetValue(scope, out var sec)) return;
            if (!sec.TryGetValue("json", out var b64) || string.IsNullOrWhiteSpace(b64)) return;
            var rawJson = Encoding.UTF8.GetString(Convert.FromBase64String(b64));
            var obj = System.Text.Json.JsonSerializer.Deserialize<object>(rawJson);
            SendMessageToWebView(new { action = "settingsLoaded", scope = scope, data = obj });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"加载设置失败: {ex.Message}" });
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
