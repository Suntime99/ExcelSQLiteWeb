using ExcelSQLiteWeb.Models;
using ExcelSQLiteWeb.Services;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeOpenXml;
using Microsoft.Data.Sqlite;
using Dapper;
using System.Drawing.Imaging;
using System.Security.Cryptography;

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

    // ==================== 脱敏体系（Vault DB + Policy Repo） ====================
    private SqliteConnection? _vaultConn;
    private SqliteConnection? _policyConn;
    private string? _vaultDbPath;
    private string? _policyDbPath;
    private byte[]? _vaultHmacSecret; // 从 DPAPI 保护的 secret 文件解密得到（CurrentUser）
    private string _outputMode = "raw"; // raw/masked
    private readonly HashSet<string> _maskedShadowViews = new(StringComparer.OrdinalIgnoreCase);
    private CancellationTokenSource? _enumArchiveCts;
    private string _webHtmlContentCache = "";
    private string _webBootUserMode = "normal"; // normal/expert（宿主注入给前端，用于“reload 切换”）
    private string? _webBootUserModeScriptId = null;
    private string? _webHtmlPathCache = null; // 若从磁盘加载，则保存 file path，便于 reload/诊断
    private string? _webBaseDirCache = null;
    private const string WebVirtualHost = "app.excelsqlite.local";

    private static string ToMode(string? m)
        => string.Equals(m, "expert", StringComparison.OrdinalIgnoreCase) ? "expert" : "normal";

    private string GetEntryHtmlFileNameForMode(string? mode)
    {
        var m = ToMode(mode);
        // 约定：两套入口壳；若不存在则回退 index.html
        var fn = (m == "expert") ? "index.expert.html" : "index.normal.html";
        try
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            if (File.Exists(Path.Combine(baseDir, fn))) return fn;
        }
        catch { }
        return "index.html";
    }

    private sealed class DesensitizationConfigV1
    {
        public string Namespace { get; set; } = "DEFAULT";
        public string VaultDbPath { get; set; } = "";
        public string PolicyDbPath { get; set; } = "";
        public string MaskedDbPath { get; set; } = "";      // 默认：<dbDir>/masked/masked_{namespace}.db
        public string OutputMode { get; set; } = "raw";     // raw/masked（用于项目持久化）
        public string MaskedTablePrefix { get; set; } = ""; // 可选：如 masked__；为空则 {table}_masked
        public bool KeepRawInMasked { get; set; } = false;  // 工具版默认 false
    }

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
    private CancellationTokenSource? _exportCts;
    private string? _exportJobId;
    // 元数据分析：扫描取消令牌（用于取消长时间扫描）
    private CancellationTokenSource? _metadataScanCts;
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

    // ==================== 项目管理（DataConfig，位于主程序目录） ====================
    // 说明：
    // - 项目文件：<主程序目录>/DataConfig/schemes/{项目名}.ini
    // - 临时库：  <主程序目录>/DataConfig/db/{项目名}.db
    // - 最近列表：<主程序目录>/DataConfig/recent.json
    private const int DefaultMaxRecentSchemesDisplay = 10; // UI 显示数量
    private const int DefaultMaxRecentSchemesKeep = 50;    // 最大保留数量
    private int _maxSchemesDisplay = DefaultMaxRecentSchemesDisplay;
    private int _maxSchemesKeep = DefaultMaxRecentSchemesKeep;

    private readonly List<SchemeMeta> _schemes = new();
    private string? _activeSchemeId;     // 项目“文件名”基名（安全化）
    private string? _activeSchemeDbPath; // 绝对路径：DataConfig/db/{id}.db
    private string _dataConfigDir = "";  // 实际 DataConfig 根目录（优先主程序目录，不可写则回退 AppData）

    private string MainTableNameOrDefault()
        => string.IsNullOrWhiteSpace(_currentMainTableName) ? "Main" : _currentMainTableName!;

    public Form1()
    {
        InitializeComponent();
        // WebView2 Runtime 缺失时：提前提示并退出（避免界面空白/初始化异常）
        if (!EnsureWebView2RuntimeReady())
        {
            try { Environment.Exit(2); } catch { }
            return;
        }
        InitializeDragDropSupport();
        _dataConfigDir = ResolveDataConfigDir();
        InitializeServices();
        EnsureDataConfigDirs();
        LoadRecentFilesFromDisk();
        LoadRecentSchemesFromDisk();
        _lastMainFilePath = _recentFiles.FirstOrDefault()?.FullPath;
        InitializeWebView2();
    }

    private static bool EnsureWebView2RuntimeReady()
    {
        try
        {
            // 若未安装 WebView2 Runtime，此调用会抛异常
            var v = Microsoft.Web.WebView2.Core.CoreWebView2Environment.GetAvailableBrowserVersionString();
            return !string.IsNullOrWhiteSpace(v);
        }
        catch
        {
            try
            {
                var msg =
                    "未检测到 WebView2 Runtime（Microsoft Edge WebView2 Runtime）。\n\n" +
                    "请先安装后再运行本程序。\n\n" +
                    "如果你的程序目录内包含 WebView2RuntimeInstallerX64.exe（或 X86），可直接运行该安装器。\n\n" +
                    "下载地址（微软官方）：\n" +
                    "https://developer.microsoft.com/microsoft-edge/webview2/\n\n" +
                    "提示：Win11/部分 Win10 可能已自带。";
                MessageBox.Show(msg, "缺少运行环境", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                // 优先尝试运行同目录内置安装器（用户选择安装）
                try
                {
                    var dir = AppContext.BaseDirectory;
                    var p1 = Path.Combine(dir, "WebView2RuntimeInstallerX64.exe");
                    var p2 = Path.Combine(dir, "WebView2RuntimeInstallerX86.exe");
                    var installer = File.Exists(p1) ? p1 : (File.Exists(p2) ? p2 : "");
                    if (!string.IsNullOrWhiteSpace(installer))
                    {
                        Process.Start(new ProcessStartInfo { FileName = installer, UseShellExecute = true });
                        return false;
                    }
                }
                catch { }

                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "https://developer.microsoft.com/microsoft-edge/webview2/",
                        UseShellExecute = true
                    });
                }
                catch { }
            }
            catch { }
            return false;
        }
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
        public string Name { get; set; } = "未命名项目";
        public DateTime LastOpenTime { get; set; } = DateTime.Now;
    }

    // ==================== Project Hub：项目元数据（跟随业务库目录） ====================
    // 说明：
    // - 项目 INI 仍保留用于“快速回填配置”（兼容历史实现）
    // - 项目元数据 JSON（*.project.json）用于“业务视图/仪表盘/审计/修订点”
    private sealed class ProjectMetaV1
    {
        public int SchemaVersion { get; set; } = 1;
        public string Id { get; set; } = "";
        public string DisplayName { get; set; } = "未命名项目";
        public string DbPath { get; set; } = "";
        public string IniPath { get; set; } = "";
        public string? BaseDbPath { get; set; } // 可选：挂载基础库（共享维表/口径表）
        public string? BaseDbAlias { get; set; } = "base";
        public List<string> BaseDbTables { get; set; } = new(); // 基础库表清单（用于项目中心展示）
        public DateTime CreatedAt { get; set; } = DateTime.Now;
        public DateTime UpdatedAt { get; set; } = DateTime.Now;
        public DateTime LastOpenAt { get; set; } = DateTime.Now;
        public long OpenCount { get; set; } = 0;
        public long DbSizeBytes { get; set; }
        public List<SourceFileMetaV1> Sources { get; set; } = new();
        public string? MainSourceFile { get; set; }
        public string? MainTableName { get; set; }
        public List<string> DbTables { get; set; } = new();
        public double? TableDiffRate { get; set; } // 0~1，缺失/多余表占比（粗略）
        public List<DerivedViewV1> DerivedViews { get; set; } = new();
        public DateTime? DerivedViewsUpdatedAt { get; set; }
        public RawCleanPolicyV1 RawCleanPolicy { get; set; } = new();
        public RevisionPolicyV1 RevisionPolicy { get; set; } = new();
        public List<RevisionPointV1> RevisionPoints { get; set; } = new();
        public ImportBatchV1? LastImport { get; set; }
        public List<AuditLogV1> AuditLogs { get; set; } = new();
        // 脱敏：跟随工程配置（可被 UI 修改并持久化到 *.project.json）
        public DesensitizationConfigV1 Desensitization { get; set; } = new();
    }

    private sealed class DerivedViewV1
    {
        public string Name { get; set; } = "";        // 强制 vw_ 前缀
        public string Sql { get; set; } = "";         // SELECT / WITH ... SELECT
        public int Version { get; set; } = 1;         // 更新 SQL 时自增
        public string Note { get; set; } = "";
        public List<string> DependsOn { get; set; } = new();
        public DateTime CreatedAt { get; set; } = DateTime.Now;
        public DateTime UpdatedAt { get; set; } = DateTime.Now;
        public bool Enabled { get; set; } = true;
    }

    private sealed class SourceFileMetaV1
    {
        public string FilePath { get; set; } = "";
        public long FileSizeBytes { get; set; }
        public long LastWriteUtcTicks { get; set; }
        public bool IsMain { get; set; }
        public List<string> Sheets { get; set; } = new();
        public Dictionary<string, string> TableNameMap { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, string> TableAliasMap { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    }

    private sealed class ImportBatchV1
    {
        public string BatchId { get; set; } = "";
        public DateTime Time { get; set; } = DateTime.Now;
        public string ImportMode { get; set; } = "text"; // text/smart
        public bool ClearedBeforeImport { get; set; } = true;
        public int ImportedTables { get; set; }
    }

    private sealed class RawCleanPolicyV1
    {
        public string Mode { get; set; } = "raw+clean"; // raw+clean / inplace
        public bool AllowArchiveRaw { get; set; } = true;
    }

    private sealed class RevisionPolicyV1
    {
        public bool EnableSqliteBackup { get; set; } = true;
        public int KeepLastN { get; set; } = 10;
        public long MaxTotalBytes { get; set; } = 1024L * 1024L * 1024L; // 1GB
    }

    private sealed class RevisionPointV1
    {
        public string Id { get; set; } = "";
        public DateTime Time { get; set; } = DateTime.Now;
        public string Type { get; set; } = "manual"; // import/clean/manual
        public string Note { get; set; } = "";
        public string? DbBackupPath { get; set; }
        public long? DbBackupBytes { get; set; }
    }

    private sealed class AuditLogV1
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public DateTime Time { get; set; } = DateTime.Now;
        public string Type { get; set; } = ""; // dbm-batch-delete / dbm-batch-rename / dbm-rename / etc
        public string Note { get; set; } = "";
        public string? RevisionPointId { get; set; }
        public List<string> RevisionPointIds { get; set; } = new();
        public List<string> ReportPaths { get; set; } = new();
        public string PayloadJson { get; set; } = ""; // 原始补充信息（JSON字符串）
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

    private string GetProjectMetaPath(string schemeId)
    {
        var db = GetSchemeDbPath(schemeId);
        var dir = Path.GetDirectoryName(db) ?? GetDbDir();
        return Path.Combine(dir, $"{schemeId}.project.json");
    }

    private ProjectMetaV1 LoadOrCreateProjectMeta(string schemeId, string? displayName = null, string? settingsJson = null)
    {
        try
        {
            var path = GetProjectMetaPath(schemeId);
            if (File.Exists(path))
            {
                var json = File.ReadAllText(path, Encoding.UTF8);
                var meta = System.Text.Json.JsonSerializer.Deserialize<ProjectMetaV1>(json) ?? new ProjectMetaV1();
                // 轻量纠偏
                meta.Id = string.IsNullOrWhiteSpace(meta.Id) ? schemeId : meta.Id;
                if (!string.IsNullOrWhiteSpace(displayName)) meta.DisplayName = displayName!;
                meta.DbPath = GetSchemeDbPath(schemeId);
                meta.IniPath = GetSchemeIniPath(schemeId);
                meta.DbSizeBytes = SafeFileSize(meta.DbPath);
                // 注意：不要在“读取”时更新 UpdatedAt，避免只浏览列表就污染修改时间
                return meta;
            }
        }
        catch { }

        // 不存在则创建最小元数据
        var m = new ProjectMetaV1
        {
            Id = schemeId,
            DisplayName = string.IsNullOrWhiteSpace(displayName) ? schemeId : displayName!,
            DbPath = GetSchemeDbPath(schemeId),
            IniPath = GetSchemeIniPath(schemeId),
            CreatedAt = DateTime.Now,
            UpdatedAt = DateTime.Now,
            LastOpenAt = DateTime.Now,
            DbSizeBytes = SafeFileSize(GetSchemeDbPath(schemeId)),
            RawCleanPolicy = new RawCleanPolicyV1 { Mode = "raw+clean", AllowArchiveRaw = true },
            RevisionPolicy = new RevisionPolicyV1 { EnableSqliteBackup = true, KeepLastN = 10, MaxTotalBytes = 1024L * 1024L * 1024L }
        };
        try
        {
            if (!string.IsNullOrWhiteSpace(settingsJson))
                m.Sources = ExtractSourcesFromSettings(settingsJson!);
        }
        catch { }
        try { SaveProjectMeta(m); } catch { }
        return m;
    }

    private void SaveProjectMeta(ProjectMetaV1 meta)
    {
        if (meta == null || string.IsNullOrWhiteSpace(meta.Id)) return;
        meta.UpdatedAt = DateTime.Now;
        meta.DbSizeBytes = SafeFileSize(meta.DbPath);
        var json = System.Text.Json.JsonSerializer.Serialize(meta, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
        var path = GetProjectMetaPath(meta.Id);
        Directory.CreateDirectory(Path.GetDirectoryName(path) ?? GetDbDir());
        File.WriteAllText(path, json, Encoding.UTF8);
    }

    // ==================== 脱敏：工程配置（路径/namespace） ====================

    private static string BuildNamespaceShort(string schemeId)
    {
        var t = (schemeId ?? "").Trim();
        if (t.Length == 0) return "DEFAULT";
        // 取字母数字，最多 4 位（P01/HR/FIN/SA01）
        var sb = new StringBuilder();
        foreach (var ch in t)
        {
            if (char.IsLetterOrDigit(ch)) sb.Append(char.ToUpperInvariant(ch));
            if (sb.Length >= 4) break;
        }
        return sb.Length == 0 ? "DEFAULT" : sb.ToString();
    }

    private DesensitizationConfigV1 EnsureDesensitizationConfigForActiveScheme()
    {
        if (string.IsNullOrWhiteSpace(_activeSchemeId))
            throw new InvalidOperationException("请先在项目中心打开一个项目");

        var schemeId = _activeSchemeId!;
        var meta = LoadOrCreateProjectMeta(schemeId, displayName: null, settingsJson: null);

        // 默认：跟随工程配置（跟随业务库目录）
        var dbDir = Path.GetDirectoryName(meta.DbPath) ?? GetDbDir();
        Directory.CreateDirectory(dbDir);
        var vaultDir = Path.Combine(dbDir, "vault");
        var policyDir = Path.Combine(dbDir, "policy");
        var maskedDir = Path.Combine(dbDir, "masked");
        Directory.CreateDirectory(vaultDir);
        Directory.CreateDirectory(policyDir);
        Directory.CreateDirectory(maskedDir);

        if (meta.Desensitization == null) meta.Desensitization = new DesensitizationConfigV1();
        if (string.IsNullOrWhiteSpace(meta.Desensitization.Namespace))
            meta.Desensitization.Namespace = BuildNamespaceShort(schemeId);
        if (string.IsNullOrWhiteSpace(meta.Desensitization.VaultDbPath))
            meta.Desensitization.VaultDbPath = Path.Combine(vaultDir, "vault.db");
        if (string.IsNullOrWhiteSpace(meta.Desensitization.PolicyDbPath))
            meta.Desensitization.PolicyDbPath = Path.Combine(policyDir, "policy.db");
        if (string.IsNullOrWhiteSpace(meta.Desensitization.MaskedDbPath))
        {
            var ns = meta.Desensitization.Namespace ?? BuildNamespaceShort(schemeId);
            meta.Desensitization.MaskedDbPath = Path.Combine(maskedDir, $"masked_{ns}.db");
        }
        if (string.IsNullOrWhiteSpace(meta.Desensitization.OutputMode))
            meta.Desensitization.OutputMode = "raw";

        // 若是首次初始化（字段为空），写回 meta（跟随工程配置）
        meta.UpdatedAt = DateTime.Now;
        SaveProjectMeta(meta);
        // 同步当前路由模式（注意：仅记录，不强制立即 Apply；由 UI 或导入完成后触发）
        _outputMode = string.Equals(meta.Desensitization.OutputMode, "masked", StringComparison.OrdinalIgnoreCase) ? "masked" : "raw";
        return meta.Desensitization;
    }

    private static void ExecSqlScript(SqliteConnection conn, string sqlScript)
    {
        if (conn.State != ConnectionState.Open) conn.Open();
        // 极简切分：按分号分割（对本项目 schema/seed 足够）
        var parts = (sqlScript ?? "")
            .Split(';')
            .Select(x => x.Trim())
            .Where(x => !string.IsNullOrWhiteSpace(x));
        foreach (var p in parts)
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = p;
            cmd.ExecuteNonQuery();
        }
    }

    private void EnsureVaultReady(DesensitizationConfigV1 cfg)
    {
        if (string.IsNullOrWhiteSpace(cfg.VaultDbPath)) throw new InvalidOperationException("VaultDbPath 为空");
        if (_vaultConn != null && string.Equals(_vaultDbPath, cfg.VaultDbPath, StringComparison.OrdinalIgnoreCase) && _vaultConn.State == ConnectionState.Open)
            return;

        try { _vaultConn?.Dispose(); } catch { }
        _vaultDbPath = cfg.VaultDbPath;
        Directory.CreateDirectory(Path.GetDirectoryName(_vaultDbPath) ?? AppContext.BaseDirectory);
        _vaultConn = new SqliteConnection($"Data Source={_vaultDbPath}");
        _vaultConn.Open();

        ExecSqlScript(_vaultConn, VaultSchemaSql);
        _vaultHmacSecret = EnsureVaultHmacSecret(Path.GetDirectoryName(_vaultDbPath) ?? AppContext.BaseDirectory);
    }

    private void EnsurePolicyRepoReady(DesensitizationConfigV1 cfg)
    {
        if (string.IsNullOrWhiteSpace(cfg.PolicyDbPath)) throw new InvalidOperationException("PolicyDbPath 为空");
        if (_policyConn != null && string.Equals(_policyDbPath, cfg.PolicyDbPath, StringComparison.OrdinalIgnoreCase) && _policyConn.State == ConnectionState.Open)
            return;

        try { _policyConn?.Dispose(); } catch { }
        _policyDbPath = cfg.PolicyDbPath;
        Directory.CreateDirectory(Path.GetDirectoryName(_policyDbPath) ?? AppContext.BaseDirectory);
        _policyConn = new SqliteConnection($"Data Source={_policyDbPath}");
        _policyConn.Open();

        ExecSqlScript(_policyConn, PolicyRepoSchemaSql);
        try { EnsurePolicyRepoMigrations(_policyConn); } catch { }
    }

    private static bool SqliteColumnExists(SqliteConnection conn, string table, string column)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"PRAGMA table_info({SqliteManager.QuoteIdent(table)});";
        using var rd = cmd.ExecuteReader();
        while (rd.Read())
        {
            var name = rd.IsDBNull(1) ? "" : rd.GetString(1);
            if (string.Equals(name, column, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    private static void EnsurePolicyRepoMigrations(SqliteConnection conn)
    {
        // v2：模板规则支持“字段匹配模式”（精确/包含/正则）
        if (!SqliteColumnExists(conn, "template_rule", "match_mode"))
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "ALTER TABLE template_rule ADD COLUMN match_mode TEXT NOT NULL DEFAULT 'exact';";
            cmd.ExecuteNonQuery();
        }
        if (!SqliteColumnExists(conn, "template_rule", "match_pattern"))
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "ALTER TABLE template_rule ADD COLUMN match_pattern TEXT;";
            cmd.ExecuteNonQuery();
        }
        // 预留：policy_rule 同样字段（后续把模板“发布为 policy”时复用）
        if (!SqliteColumnExists(conn, "policy_rule", "match_mode"))
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "ALTER TABLE policy_rule ADD COLUMN match_mode TEXT NOT NULL DEFAULT 'exact';";
            cmd.ExecuteNonQuery();
        }
        if (!SqliteColumnExists(conn, "policy_rule", "match_pattern"))
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "ALTER TABLE policy_rule ADD COLUMN match_pattern TEXT;";
            cmd.ExecuteNonQuery();
        }
    }

    private static byte[] EnsureVaultHmacSecret(string dir)
    {
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "vault_secret.dpapi.bin");
        if (!File.Exists(path))
        {
            var raw = RandomNumberGenerator.GetBytes(32);
            var protectedBytes = ProtectedData.Protect(raw, optionalEntropy: null, scope: DataProtectionScope.CurrentUser);
            File.WriteAllBytes(path, protectedBytes);
            return raw;
        }
        var enc = File.ReadAllBytes(path);
        return ProtectedData.Unprotect(enc, optionalEntropy: null, scope: DataProtectionScope.CurrentUser);
    }

    private static byte[] EncryptByDpapi(string plainText)
    {
        var bytes = Encoding.UTF8.GetBytes(plainText ?? "");
        return ProtectedData.Protect(bytes, optionalEntropy: null, scope: DataProtectionScope.CurrentUser);
    }

    private static string DecryptByDpapi(byte[] cipher)
    {
        var plain = ProtectedData.Unprotect(cipher, optionalEntropy: null, scope: DataProtectionScope.CurrentUser);
        return Encoding.UTF8.GetString(plain);
    }

    private static string HmacSha256Hex(byte[] secret, string text)
    {
        using var h = new HMACSHA256(secret);
        var bytes = Encoding.UTF8.GetBytes(text ?? "");
        var hash = h.ComputeHash(bytes);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }

    private static readonly char[] Base32 = "0123456789ABCDEFGHJKMNPQRSTVWXYZ".ToCharArray(); // Crockford Base32（去 I L O U）
    private static string RandomBase32(int len)
    {
        var b = RandomNumberGenerator.GetBytes(len);
        var sb = new StringBuilder(len);
        for (int i = 0; i < len; i++)
            sb.Append(Base32[b[i] % Base32.Length]);
        return sb.ToString();
    }

    private static char CheckCharCrc8(string s)
    {
        // 轻量 CRC8（多项式 0x07）
        byte crc = 0;
        foreach (var ch in Encoding.ASCII.GetBytes(s ?? ""))
        {
            crc ^= ch;
            for (int i = 0; i < 8; i++)
                crc = (byte)(((crc & 0x80) != 0) ? ((crc << 1) ^ 0x07) : (crc << 1));
        }
        return Base32[crc % Base32.Length];
    }

    private static string BuildReadableToken(string type, string ns)
    {
        var t = (type ?? "UNK").Trim().ToUpperInvariant();
        var n = (ns ?? "DEFAULT").Trim().ToUpperInvariant();
        var body = RandomBase32(12);
        var withoutChk = $"{t}_{n}_{body}";
        var chk = CheckCharCrc8(withoutChk);
        return $"{withoutChk}_{chk}";
    }

    private static string NormalizeValue(string type, string raw)
    {
        var t = (type ?? "").Trim().ToUpperInvariant();
        var s = (raw ?? "").Trim();
        if (s.Length == 0) return "";

        if (t == "PHONE")
        {
            var digits = new string(s.Where(char.IsDigit).ToArray());
            if (digits.StartsWith("86") && digits.Length > 11) digits = digits.Substring(digits.Length - 11);
            return digits;
        }
        if (t == "IDNO" || t == "TAX_ID" || t == "ORG_CODE" || t == "BANK_ACCOUNT" ||
            t == "LICENSE_NO" || t == "CERT_NO" || t == "INVOICE_NO" || t == "INVOICE_TAX_ID" ||
            t == "PAYMENT_CHANNEL" || t == "TRANSACTION_ID" || t == "CONTRACT_NO" || t == "SEAL_INFO")
        {
            var x = s.Replace(" ", "").Replace("\t", "");
            x = x.ToUpperInvariant();
            return x;
        }
        if (t == "EMAIL")
        {
            var x = s.Trim();
            var at = x.LastIndexOf('@');
            if (at > 0 && at < x.Length - 1)
            {
                var local = x.Substring(0, at);
                var dom = x.Substring(at + 1).ToLowerInvariant();
                return local + "@" + dom;
            }
            return x;
        }
        if (t == "NAME" || t == "ADDR" ||
            t == "ORG_NAME" || t == "ORG_ADDR" || t == "ORG_CITY" || t == "ORG_BRAND" || t == "ORG_UNIT" ||
            t == "CONTACT_NAME" || t == "BANK_NAME" || t == "BANK_BRANCH" || t == "ACCOUNT_NAME" ||
            t == "LEGAL_PERSON" || t == "SIGNATURE" || t == "SHAREHOLDER" || t == "ULTIMATE_BENEFICIARY" || t == "PARENT_ORG" ||
            t == "INVOICE_TITLE" || t == "INDUSTRY" || t == "BUSINESS_SCOPE" || t == "ORG_SCALE" || t == "ORG_TYPE" || t == "REGION")
        {
            // 多空格合一
            var x = System.Text.RegularExpressions.Regex.Replace(s, @"\s+", " ").Trim();
            return x;
        }
        return s;
    }

    private static string MaskValue(string type, string raw)
    {
        var t = (type ?? "").Trim().ToUpperInvariant();
        var s = (raw ?? "").Trim();
        if (s.Length == 0) return "";
        if (t == "PHONE")
        {
            var digits = new string(s.Where(char.IsDigit).ToArray());
            if (digits.Length >= 11) return digits.Substring(0, 3) + "****" + digits.Substring(digits.Length - 4);
            return s.Length <= 2 ? "**" : (s.Substring(0, 2) + "****");
        }
        if (t == "IDNO")
        {
            if (s.Length <= 8) return "****";
            return s.Substring(0, 6) + new string('*', Math.Max(4, s.Length - 10)) + s.Substring(s.Length - 4);
        }
        if (t == "TAX_ID" || t == "ORG_CODE")
        {
            if (s.Length <= 6) return "****";
            return s.Substring(0, 2) + new string('*', Math.Max(4, s.Length - 4)) + s.Substring(s.Length - 2);
        }
        if (t == "LICENSE_NO" || t == "CERT_NO" || t == "INVOICE_NO" || t == "INVOICE_TAX_ID" || t == "TRANSACTION_ID" || t == "PAYMENT_CHANNEL" || t == "CONTRACT_NO" || t == "SEAL_INFO")
        {
            if (s.Length <= 6) return "****";
            return s.Substring(0, 2) + "****" + s.Substring(s.Length - 2);
        }
        if (t == "BANK_ACCOUNT")
        {
            // 银行账号：只留后4位
            var digits = new string(s.Where(char.IsLetterOrDigit).ToArray());
            if (digits.Length <= 4) return "****";
            return "****" + digits.Substring(digits.Length - 4);
        }
        if (t == "NAME")
        {
            return s.Length <= 1 ? "*" : (s.Substring(0, 1) + new string('*', Math.Min(4, s.Length - 1)));
        }
        if (t == "ORG_NAME" || t == "CONTACT_NAME" || t == "ACCOUNT_NAME")
        {
            return s.Length <= 2 ? "*" : (s.Substring(0, 2) + new string('*', Math.Min(6, Math.Max(1, s.Length - 2))));
        }
        if (t == "LEGAL_PERSON" || t == "SIGNATURE" || t == "SHAREHOLDER" || t == "ULTIMATE_BENEFICIARY")
        {
            return s.Length <= 2 ? "*" : (s.Substring(0, 2) + new string('*', Math.Min(6, Math.Max(1, s.Length - 2))));
        }
        if (t == "ADDR")
        {
            return s.Length <= 6 ? (s.Substring(0, 1) + "****") : (s.Substring(0, 6) + "****");
        }
        if (t == "ORG_ADDR")
        {
            return s.Length <= 6 ? (s.Substring(0, 1) + "****") : (s.Substring(0, 6) + "****");
        }
        if (t == "ORG_PHONE" || t == "CONTACT_PHONE")
        {
            var digits = new string(s.Where(char.IsDigit).ToArray());
            if (digits.Length >= 11) return digits.Substring(0, 3) + "****" + digits.Substring(digits.Length - 4);
            return s.Length <= 2 ? "**" : (s.Substring(0, 2) + "****");
        }
        return s.Length <= 4 ? "****" : (s.Substring(0, 2) + "****" + s.Substring(s.Length - 2));
    }

    // Vault/Policy Repo 内置 schema（不依赖外部文件，方便发布为单 exe）
    private const string VaultSchemaSql = @"
CREATE TABLE IF NOT EXISTS namespace_registry (
  namespace TEXT PRIMARY KEY,
  display_name TEXT,
  created_at TEXT NOT NULL,
  created_by TEXT,
  note TEXT
);
CREATE TABLE IF NOT EXISTS token_map (
  id TEXT PRIMARY KEY,
  namespace TEXT NOT NULL,
  type TEXT NOT NULL,
  fingerprint TEXT NOT NULL,
  token TEXT NOT NULL,
  enc_value BLOB NOT NULL,
  norm_hint TEXT,
  policy_id TEXT,
  policy_version INTEGER,
  created_at TEXT NOT NULL,
  created_by TEXT,
  last_used_at TEXT,
  use_count INTEGER NOT NULL DEFAULT 0,
  UNIQUE(namespace, type, fingerprint),
  UNIQUE(token)
);
CREATE INDEX IF NOT EXISTS idx_token_map_ns_type ON token_map(namespace, type);
CREATE INDEX IF NOT EXISTS idx_token_map_token ON token_map(token);
CREATE INDEX IF NOT EXISTS idx_token_map_fpr ON token_map(fingerprint);
CREATE TABLE IF NOT EXISTS audit_log (
  id TEXT PRIMARY KEY,
  action TEXT NOT NULL,
  namespace TEXT,
  operator TEXT,
  role TEXT,
  reason_ticket TEXT,
  input_ref TEXT,
  output_ref TEXT,
  policy_id TEXT,
  policy_version INTEGER,
  row_count INTEGER,
  col_count INTEGER,
  started_at TEXT,
  finished_at TEXT,
  status TEXT NOT NULL DEFAULT 'OK',
  error_message TEXT,
  detail_json TEXT,
  created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_audit_action_time ON audit_log(action, created_at);
CREATE INDEX IF NOT EXISTS idx_audit_ns_time ON audit_log(namespace, created_at);
";

    private const string PolicyRepoSchemaSql = @"
CREATE TABLE IF NOT EXISTS policy (
  id TEXT PRIMARY KEY,
  namespace TEXT NOT NULL,
  name TEXT NOT NULL,
  description TEXT,
  current_version INTEGER NOT NULL DEFAULT 1,
  status TEXT NOT NULL DEFAULT 'draft',
  created_at TEXT NOT NULL,
  created_by TEXT,
  updated_at TEXT,
  updated_by TEXT
);
CREATE INDEX IF NOT EXISTS idx_policy_ns ON policy(namespace);
CREATE INDEX IF NOT EXISTS idx_policy_status ON policy(status);
CREATE TABLE IF NOT EXISTS policy_version (
  id TEXT PRIMARY KEY,
  policy_id TEXT NOT NULL,
  version INTEGER NOT NULL,
  note TEXT,
  created_at TEXT NOT NULL,
  created_by TEXT,
  UNIQUE(policy_id, version)
);
CREATE TABLE IF NOT EXISTS policy_rule (
  id TEXT PRIMARY KEY,
  policy_id TEXT NOT NULL,
  policy_version INTEGER NOT NULL,
  table_name TEXT,
  column_name TEXT NOT NULL,
  match_mode TEXT NOT NULL DEFAULT 'exact',
  match_pattern TEXT,
  column_alias TEXT,
  data_type TEXT NOT NULL,
  action TEXT NOT NULL,
  output_token_col TEXT,
  output_mask_col TEXT,
  keep_raw_col INTEGER NOT NULL DEFAULT 0,
  normalize_profile TEXT NOT NULL DEFAULT 'default',
  normalize_params TEXT,
  on_error TEXT NOT NULL DEFAULT 'fail',
  enabled INTEGER NOT NULL DEFAULT 1,
  sort_order INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_policy_rule_policy ON policy_rule(policy_id, policy_version);
CREATE TABLE IF NOT EXISTS template (
  id TEXT PRIMARY KEY,
  namespace TEXT NOT NULL,
  name TEXT NOT NULL,
  description TEXT,
  created_at TEXT NOT NULL,
  created_by TEXT,
  updated_at TEXT,
  updated_by TEXT,
  UNIQUE(namespace, name)
);
CREATE INDEX IF NOT EXISTS idx_template_ns ON template(namespace);
CREATE TABLE IF NOT EXISTS template_rule (
  id TEXT PRIMARY KEY,
  template_id TEXT NOT NULL,
  table_name TEXT,
  column_name TEXT NOT NULL,
  match_mode TEXT NOT NULL DEFAULT 'exact',
  match_pattern TEXT,
  data_type TEXT NOT NULL,
  action TEXT NOT NULL,
  output_token_col TEXT,
  output_mask_col TEXT,
  keep_raw_col INTEGER NOT NULL DEFAULT 0,
  normalize_profile TEXT NOT NULL DEFAULT 'default',
  normalize_params TEXT,
  on_error TEXT NOT NULL DEFAULT 'fail',
  enabled INTEGER NOT NULL DEFAULT 1,
  sort_order INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_template_rule_tpl ON template_rule(template_id);
";

    private const string DefaultTemplateSeedSql = @"
INSERT OR IGNORE INTO template(id, namespace, name, description, created_at, created_by, updated_at, updated_by)
VALUES ('tpl_pii_basic_tokenize_v1','DEFAULT','PII-基础脱敏(Tokenize)','基础 PII 脱敏模板：手机号/证件号/姓名/地址 -> *_token',datetime('now'),'system',datetime('now'),'system');
INSERT OR IGNORE INTO template_rule(id, template_id, table_name, column_name, data_type, action, output_token_col, output_mask_col, keep_raw_col, normalize_profile, normalize_params, on_error, enabled, sort_order)
VALUES
('tplr_phone_zh_手机号','tpl_pii_basic_tokenize_v1',NULL,'手机号','PHONE','TOKENIZE',NULL,NULL,0,'phone',NULL,'fail',1,10),
('tplr_phone_en_phone','tpl_pii_basic_tokenize_v1',NULL,'phone','PHONE','TOKENIZE',NULL,NULL,0,'phone',NULL,'fail',1,11),
('tplr_phone_en_mobile','tpl_pii_basic_tokenize_v1',NULL,'mobile','PHONE','TOKENIZE',NULL,NULL,0,'phone',NULL,'fail',1,12),
('tplr_idno_zh_身份证号','tpl_pii_basic_tokenize_v1',NULL,'身份证号','IDNO','TOKENIZE',NULL,NULL,0,'idno',NULL,'fail',1,20),
('tplr_idno_zh_身份证','tpl_pii_basic_tokenize_v1',NULL,'身份证','IDNO','TOKENIZE',NULL,NULL,0,'idno',NULL,'fail',1,21),
('tplr_name_zh_姓名','tpl_pii_basic_tokenize_v1',NULL,'姓名','NAME','TOKENIZE',NULL,NULL,0,'name',NULL,'fail',1,30),
('tplr_addr_zh_地址','tpl_pii_basic_tokenize_v1',NULL,'地址','ADDR','TOKENIZE',NULL,NULL,0,'addr',NULL,'fail',1,40),
('tplr_addr_zh_详细地址','tpl_pii_basic_tokenize_v1',NULL,'详细地址','ADDR','TOKENIZE',NULL,NULL,0,'addr',NULL,'fail',1,41);
";

    // 企业基础模板：企业自身/客户/供应商/银行（示例）
    private const string EnterpriseBasicTemplateSeedSql = @"
INSERT OR IGNORE INTO template(id, namespace, name, description, created_at, created_by, updated_at, updated_by)
VALUES ('tpl_enterprise_self_basic_v1','DEFAULT','企业自身-基础脱敏(Tokenize)','企业自身信息：公司名称/税号/电话/地址/联系人等',datetime('now'),'system',datetime('now'),'system');
INSERT OR IGNORE INTO template_rule(id, template_id, table_name, column_name, match_mode, match_pattern, data_type, action, output_token_col, output_mask_col, keep_raw_col, normalize_profile, normalize_params, on_error, enabled, sort_order)
VALUES
('tplr_self_org_name','tpl_enterprise_self_basic_v1',NULL,'公司名称','contains','公司名称','ORG_NAME','TOKENIZE',NULL,NULL,0,'default',NULL,'fail',1,10),
('tplr_self_tax_id','tpl_enterprise_self_basic_v1',NULL,'税号','contains','税号','TAX_ID','TOKENIZE',NULL,NULL,0,'default',NULL,'fail',1,20),
('tplr_self_org_code','tpl_enterprise_self_basic_v1',NULL,'组织机构代码','contains','组织机构代码','ORG_CODE','TOKENIZE',NULL,NULL,0,'default',NULL,'fail',1,21),
('tplr_self_org_phone','tpl_enterprise_self_basic_v1',NULL,'公司电话','contains','电话','ORG_PHONE','TOKENIZE',NULL,NULL,0,'phone',NULL,'fail',1,30),
('tplr_self_org_addr','tpl_enterprise_self_basic_v1',NULL,'经营地址','contains','地址','ORG_ADDR','TOKENIZE',NULL,NULL,0,'addr',NULL,'fail',1,40),
('tplr_self_contact','tpl_enterprise_self_basic_v1',NULL,'联系人','contains','联系人','CONTACT_NAME','TOKENIZE',NULL,NULL,0,'name',NULL,'fail',1,50),
('tplr_self_contact_ph','tpl_enterprise_self_basic_v1',NULL,'联系人电话','contains','联系人电话','CONTACT_PHONE','TOKENIZE',NULL,NULL,0,'phone',NULL,'fail',1,51);

INSERT OR IGNORE INTO template(id, namespace, name, description, created_at, created_by, updated_at, updated_by)
VALUES ('tpl_enterprise_customer_basic_v1','DEFAULT','客户-基础脱敏(Tokenize)','客户信息：客户名称/税号/电话/联系人/地址',datetime('now'),'system',datetime('now'),'system');
INSERT OR IGNORE INTO template_rule(id, template_id, table_name, column_name, match_mode, match_pattern, data_type, action, output_token_col, output_mask_col, keep_raw_col, normalize_profile, normalize_params, on_error, enabled, sort_order)
VALUES
('tplr_cust_name','tpl_enterprise_customer_basic_v1',NULL,'客户名称','contains','客户名称','ORG_NAME','TOKENIZE',NULL,NULL,0,'default',NULL,'fail',1,10),
('tplr_cust_tax','tpl_enterprise_customer_basic_v1',NULL,'客户税号','contains','税号','TAX_ID','TOKENIZE',NULL,NULL,0,'default',NULL,'fail',1,20),
('tplr_cust_phone','tpl_enterprise_customer_basic_v1',NULL,'客户电话','contains','电话','ORG_PHONE','TOKENIZE',NULL,NULL,0,'phone',NULL,'fail',1,30),
('tplr_cust_contact','tpl_enterprise_customer_basic_v1',NULL,'客户联系人','contains','联系人','CONTACT_NAME','TOKENIZE',NULL,NULL,0,'name',NULL,'fail',1,40),
('tplr_cust_addr','tpl_enterprise_customer_basic_v1',NULL,'客户地址','contains','地址','ORG_ADDR','TOKENIZE',NULL,NULL,0,'addr',NULL,'fail',1,50);

INSERT OR IGNORE INTO template(id, namespace, name, description, created_at, created_by, updated_at, updated_by)
VALUES ('tpl_enterprise_supplier_basic_v1','DEFAULT','供应商-基础脱敏(Tokenize)','供应商信息：供应商名称/税号/电话/联系人/地址',datetime('now'),'system',datetime('now'),'system');
INSERT OR IGNORE INTO template_rule(id, template_id, table_name, column_name, match_mode, match_pattern, data_type, action, output_token_col, output_mask_col, keep_raw_col, normalize_profile, normalize_params, on_error, enabled, sort_order)
VALUES
('tplr_supp_name','tpl_enterprise_supplier_basic_v1',NULL,'供应商名称','contains','供应商名称','ORG_NAME','TOKENIZE',NULL,NULL,0,'default',NULL,'fail',1,10),
('tplr_supp_tax','tpl_enterprise_supplier_basic_v1',NULL,'供应商税号','contains','税号','TAX_ID','TOKENIZE',NULL,NULL,0,'default',NULL,'fail',1,20),
('tplr_supp_phone','tpl_enterprise_supplier_basic_v1',NULL,'供应商电话','contains','电话','ORG_PHONE','TOKENIZE',NULL,NULL,0,'phone',NULL,'fail',1,30),
('tplr_supp_contact','tpl_enterprise_supplier_basic_v1',NULL,'供应商联系人','contains','联系人','CONTACT_NAME','TOKENIZE',NULL,NULL,0,'name',NULL,'fail',1,40),
('tplr_supp_addr','tpl_enterprise_supplier_basic_v1',NULL,'供应商地址','contains','地址','ORG_ADDR','TOKENIZE',NULL,NULL,0,'addr',NULL,'fail',1,50);

INSERT OR IGNORE INTO template(id, namespace, name, description, created_at, created_by, updated_at, updated_by)
VALUES ('tpl_enterprise_bank_basic_v1','DEFAULT','财务-银行信息(Tokenize)','银行账号/开户行/开户人等',datetime('now'),'system',datetime('now'),'system');
INSERT OR IGNORE INTO template_rule(id, template_id, table_name, column_name, match_mode, match_pattern, data_type, action, output_token_col, output_mask_col, keep_raw_col, normalize_profile, normalize_params, on_error, enabled, sort_order)
VALUES
('tplr_bank_account','tpl_enterprise_bank_basic_v1',NULL,'银行账号','contains','账号','BANK_ACCOUNT','TOKENIZE',NULL,NULL,0,'default',NULL,'fail',1,10),
('tplr_bank_name','tpl_enterprise_bank_basic_v1',NULL,'开户行','contains','开户行','BANK_NAME','MASK',NULL,NULL,1,'default',NULL,'fail',1,20),
('tplr_bank_holder','tpl_enterprise_bank_basic_v1',NULL,'开户人','contains','开户','ACCOUNT_NAME','TOKENIZE',NULL,NULL,0,'default',NULL,'fail',1,30);
";

    // 企业扩展模板：合规/法务/票据/流水/画像/资产（示例）
    private const string EnterpriseExtendedTemplateSeedSql = @"
INSERT OR IGNORE INTO template(id, namespace, name, description, created_at, created_by, updated_at, updated_by)
VALUES ('tpl_enterprise_compliance_cert_v1','DEFAULT','合规-法人/证照(Tokenize)','法人代表、营业执照/许可证/资质编号等',datetime('now'),'system',datetime('now'),'system');
INSERT OR IGNORE INTO template_rule(id, template_id, table_name, column_name, match_mode, match_pattern, data_type, action, output_token_col, output_mask_col, keep_raw_col, normalize_profile, normalize_params, on_error, enabled, sort_order)
VALUES
('tplr_legal_person','tpl_enterprise_compliance_cert_v1',NULL,'法定代表人','contains','法定代表人','LEGAL_PERSON','TOKENIZE',NULL,NULL,0,'name',NULL,'fail',1,10),
('tplr_license_no','tpl_enterprise_compliance_cert_v1',NULL,'营业执照','contains','营业执照','LICENSE_NO','TOKENIZE',NULL,NULL,0,'default',NULL,'fail',1,20),
('tplr_cert_no','tpl_enterprise_compliance_cert_v1',NULL,'资质','contains','资质','CERT_NO','TOKENIZE',NULL,NULL,0,'default',NULL,'fail',1,30);

INSERT OR IGNORE INTO template(id, namespace, name, description, created_at, created_by, updated_at, updated_by)
VALUES ('tpl_enterprise_finance_invoice_txn_v1','DEFAULT','财务-发票/流水(Tokenize)','发票号码/代码、交易流水号、商户号等',datetime('now'),'system',datetime('now'),'system');
INSERT OR IGNORE INTO template_rule(id, template_id, table_name, column_name, match_mode, match_pattern, data_type, action, output_token_col, output_mask_col, keep_raw_col, normalize_profile, normalize_params, on_error, enabled, sort_order)
VALUES
('tplr_invoice_no','tpl_enterprise_finance_invoice_txn_v1',NULL,'发票','contains','发票','INVOICE_NO','TOKENIZE',NULL,NULL,0,'default',NULL,'fail',1,10),
('tplr_invoice_tax','tpl_enterprise_finance_invoice_txn_v1',NULL,'开票税号','contains','税号','INVOICE_TAX_ID','TOKENIZE',NULL,NULL,0,'default',NULL,'fail',1,20),
('tplr_txn_id','tpl_enterprise_finance_invoice_txn_v1',NULL,'流水','contains','流水','TRANSACTION_ID','TOKENIZE',NULL,NULL,0,'default',NULL,'fail',1,30),
('tplr_mch_id','tpl_enterprise_finance_invoice_txn_v1',NULL,'商户','contains','商户','PAYMENT_CHANNEL','TOKENIZE',NULL,NULL,0,'default',NULL,'fail',1,40);

INSERT OR IGNORE INTO template(id, namespace, name, description, created_at, created_by, updated_at, updated_by)
VALUES ('tpl_enterprise_legal_contract_v1','DEFAULT','法务-合同/签章(Tokenize)','合同编号、签章标识、签署人等',datetime('now'),'system',datetime('now'),'system');
INSERT OR IGNORE INTO template_rule(id, template_id, table_name, column_name, match_mode, match_pattern, data_type, action, output_token_col, output_mask_col, keep_raw_col, normalize_profile, normalize_params, on_error, enabled, sort_order)
VALUES
('tplr_contract_no','tpl_enterprise_legal_contract_v1',NULL,'合同','contains','合同','CONTRACT_NO','TOKENIZE',NULL,NULL,0,'default',NULL,'fail',1,10),
('tplr_seal','tpl_enterprise_legal_contract_v1',NULL,'签章','contains','签章','SEAL_INFO','TOKENIZE',NULL,NULL,0,'default',NULL,'fail',1,20),
('tplr_signature','tpl_enterprise_legal_contract_v1',NULL,'签署','contains','签署','SIGNATURE','TOKENIZE',NULL,NULL,0,'name',NULL,'fail',1,30);

INSERT OR IGNORE INTO template(id, namespace, name, description, created_at, created_by, updated_at, updated_by)
VALUES ('tpl_enterprise_profile_mask_v1','DEFAULT','画像-行业/规模/品牌(Mask)','行业、经营范围、品牌、企业规模等（展示/共享需泛化）',datetime('now'),'system',datetime('now'),'system');
INSERT OR IGNORE INTO template_rule(id, template_id, table_name, column_name, match_mode, match_pattern, data_type, action, output_token_col, output_mask_col, keep_raw_col, normalize_profile, normalize_params, on_error, enabled, sort_order)
VALUES
('tplr_industry','tpl_enterprise_profile_mask_v1',NULL,'行业','contains','行业','INDUSTRY','MASK',NULL,NULL,1,'default',NULL,'fail',1,10),
('tplr_brand','tpl_enterprise_profile_mask_v1',NULL,'品牌','contains','品牌','ORG_BRAND','MASK',NULL,NULL,1,'default',NULL,'fail',1,20),
('tplr_scope','tpl_enterprise_profile_mask_v1',NULL,'经营范围','contains','经营范围','BUSINESS_SCOPE','MASK',NULL,NULL,1,'default',NULL,'fail',1,30),
('tplr_scale','tpl_enterprise_profile_mask_v1',NULL,'规模','contains','规模','ORG_SCALE','MASK',NULL,NULL,1,'default',NULL,'fail',1,40);
";

    private static long SafeFileSize(string path)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(path)) return 0;
            var fi = new FileInfo(path);
            return fi.Exists ? fi.Length : 0;
        }
        catch { return 0; }
    }

    private static List<SourceFileMetaV1> ExtractSourcesFromSettings(string settingsJson)
    {
        var list = new List<SourceFileMetaV1>();
        using var doc = System.Text.Json.JsonDocument.Parse(string.IsNullOrWhiteSpace(settingsJson) ? "{}" : settingsJson);
        if (!doc.RootElement.TryGetProperty("sources", out var sources) || sources.ValueKind != System.Text.Json.JsonValueKind.Array)
            return list;

        foreach (var s in sources.EnumerateArray())
        {
            var fp = (s.TryGetProperty("filePath", out var fpEl) ? fpEl.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(fp)) continue;
            var sheets = new List<string>();
            if (s.TryGetProperty("sheets", out var sh) && sh.ValueKind == System.Text.Json.JsonValueKind.Array)
                sheets = sh.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
            var tnMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                if (s.TryGetProperty("tableNameMap", out var tm) && tm.ValueKind == System.Text.Json.JsonValueKind.Object)
                {
                    foreach (var p in tm.EnumerateObject())
                    {
                        var k = p.Name ?? "";
                        var v = p.Value.ValueKind == System.Text.Json.JsonValueKind.String ? (p.Value.GetString() ?? "") : p.Value.ToString();
                        if (!string.IsNullOrWhiteSpace(k) && !string.IsNullOrWhiteSpace(v)) tnMap[k] = v;
                    }
                }
            }
            catch { }
            var taMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                if (s.TryGetProperty("tableAliasMap", out var am) && am.ValueKind == System.Text.Json.JsonValueKind.Object)
                {
                    foreach (var p in am.EnumerateObject())
                    {
                        var k = p.Name ?? "";
                        var v = p.Value.ValueKind == System.Text.Json.JsonValueKind.String ? (p.Value.GetString() ?? "") : p.Value.ToString();
                        if (!string.IsNullOrWhiteSpace(k) && !string.IsNullOrWhiteSpace(v)) taMap[k] = v;
                    }
                }
            }
            catch { }
            bool isMain = (s.TryGetProperty("isMain", out var im) && im.ValueKind == System.Text.Json.JsonValueKind.True) ||
                          (s.TryGetProperty("role", out var role) && string.Equals(role.GetString(), "main", StringComparison.OrdinalIgnoreCase));

            long size = 0, ticks = 0;
            try
            {
                var fi = new FileInfo(fp);
                if (fi.Exists)
                {
                    size = fi.Length;
                    ticks = fi.LastWriteTimeUtc.Ticks;
                }
            }
            catch { }

            list.Add(new SourceFileMetaV1
            {
                FilePath = fp.Replace('\\', '/'),
                FileSizeBytes = size,
                LastWriteUtcTicks = ticks,
                IsMain = isMain,
                Sheets = sheets,
                TableNameMap = tnMap,
                TableAliasMap = taMap
            });
        }
        return list;
    }

    private static string? ExtractMainTableNameFromSettings(string settingsJson)
    {
        try
        {
            using var doc = System.Text.Json.JsonDocument.Parse(string.IsNullOrWhiteSpace(settingsJson) ? "{}" : settingsJson);
            if (doc.RootElement.TryGetProperty("sqliteMainTableName", out var mt) && mt.ValueKind == System.Text.Json.JsonValueKind.String)
            {
                var v = mt.GetString();
                if (!string.IsNullOrWhiteSpace(v)) return v;
            }
        }
        catch { }
        return null;
    }

    private List<string> GetDbTablesForScheme(string schemeId)
    {
        try
        {
            var db = GetSchemeDbPath(schemeId);
            if (!File.Exists(db)) return new List<string>();
            using var conn = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={db}");
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT name FROM sqlite_master WHERE type IN ('table','view') AND name NOT LIKE 'sqlite_%' ORDER BY name;";
            using var rd = cmd.ExecuteReader();
            var list = new List<string>();
            while (rd.Read())
            {
                var n = rd.IsDBNull(0) ? "" : (rd.GetString(0) ?? "");
                if (!string.IsNullOrWhiteSpace(n)) list.Add(n);
            }
            return list;
        }
        catch { return new List<string>(); }
    }

    private static List<string> ExpectedTableNamesFromSources(List<SourceFileMetaV1> sources)
    {
        var list = new List<string>();
        foreach (var s in sources ?? new List<SourceFileMetaV1>())
        {
            var fp = s.FilePath ?? "";
            var baseName = Path.GetFileNameWithoutExtension(fp.Replace('/', Path.DirectorySeparatorChar));
            if (string.IsNullOrWhiteSpace(baseName)) continue;
            var sheets = s.Sheets ?? new List<string>();
            foreach (var sh in sheets)
            {
                if (string.IsNullOrWhiteSpace(sh)) continue;
                if (s.TableNameMap != null && s.TableNameMap.TryGetValue(sh, out var tn) && !string.IsNullOrWhiteSpace(tn))
                    list.Add(tn.Replace("[", "（").Replace("]", "）"));
                else
                    list.Add($"{baseName}|{sh}".Replace("[", "（").Replace("]", "）"));
            }
        }
        return list.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static bool IsSelectOnlySqlForView(string sql)
    {
        if (string.IsNullOrWhiteSpace(sql)) return false;
        string StripComments(string s)
        {
            var x = System.Text.RegularExpressions.Regex.Replace(s ?? "", @"--.*?$", "", System.Text.RegularExpressions.RegexOptions.Multiline);
            x = System.Text.RegularExpressions.Regex.Replace(x, @"/\*[\s\S]*?\*/", "");
            return x;
        }
        bool IsSingleStatement(string s)
        {
            var t = (s ?? "").Trim();
            t = System.Text.RegularExpressions.Regex.Replace(t, @";+\s*$", "");
            return !t.Contains(';');
        }
        var t0 = StripComments(sql).TrimStart();
        if (!IsSingleStatement(t0)) return false;
        return t0.StartsWith("select", StringComparison.OrdinalIgnoreCase)
               || t0.StartsWith("with", StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeSqlForView(string sql)
    {
        var t = (sql ?? "").Trim();
        // 去掉末尾分号
        t = System.Text.RegularExpressions.Regex.Replace(t, @";+\s*$", "");
        return t;
    }

    private static List<string> GuessDependsOnTables(string sql)
    {
        var list = new List<string>();
        try
        {
            var s = System.Text.RegularExpressions.Regex.Replace(sql ?? "", @"--.*?$", "", System.Text.RegularExpressions.RegexOptions.Multiline);
            s = System.Text.RegularExpressions.Regex.Replace(s, @"/\*[\s\S]*?\*/", "");
            // 支持：Table / base.Table / [base].[Table] / "base"."Table" / `base`.`Table`
            var rx = new System.Text.RegularExpressions.Regex(
                @"\b(from|join)\s+(\[[^\]]+\](?:\s*\.\s*\[[^\]]+\])?|" +
                @"""[^""]+""(?:\s*\.\s*""[^""]+"")?|" +
                @"`[^`]+`(?:\s*\.\s*`[^`]+`)?|" +
                @"[A-Za-z0-9_\u4e00-\u9fa5\$\._]+)",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            foreach (System.Text.RegularExpressions.Match m in rx.Matches(s))
            {
                var token = m.Groups[2].Value.Trim();
                // 规范化：去掉空格与两侧引用符
                token = System.Text.RegularExpressions.Regex.Replace(token, @"\s+", "");
                if (token.StartsWith("[") && token.EndsWith("]") && !token.Contains("].["))
                    token = token[1..^1].Replace("]]", "]");
                if (token.StartsWith("\"") && token.EndsWith("\"") && !token.Contains("\".\""))
                    token = token[1..^1].Replace("\"\"", "\"");
                if (token.StartsWith("`") && token.EndsWith("`") && !token.Contains("`.`"))
                    token = token[1..^1];
                // [a].[b] -> a.b, "a"."b" -> a.b, `a`.`b` -> a.b
                token = token.Replace("].[", ".").Replace("[", "").Replace("]", "");
                token = token.Replace("\".\"", ".").Replace("\"", "");
                token = token.Replace("`.`", ".").Replace("`", "");
                token = token.Trim('.'); // 容错
                if (!string.IsNullOrWhiteSpace(token)) list.Add(token);
            }
        }
        catch { }
        return list.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static string ReplaceSqlIdentifier(string sql, string oldName, string newName)
    {
        if (string.IsNullOrWhiteSpace(sql)) return sql;
        oldName = (oldName ?? "").Trim();
        newName = (newName ?? "").Trim();
        if (string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName)) return sql;

        string EscapeRegex(string s) => System.Text.RegularExpressions.Regex.Escape(s);

        // 1) [old] -> [new]（严格）
        var s1 = System.Text.RegularExpressions.Regex.Replace(
            sql,
            @"\[" + EscapeRegex(oldName) + @"\]",
            "[" + newName.Replace("]", "]]") + "]",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        // 2) "old" -> "new"
        s1 = System.Text.RegularExpressions.Regex.Replace(
            s1,
            "\"" + EscapeRegex(oldName) + "\"",
            "\"" + newName.Replace("\"", "\"\"") + "\"",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        // 3) `old` -> `new`
        s1 = System.Text.RegularExpressions.Regex.Replace(
            s1,
            "`" + EscapeRegex(oldName) + "`",
            "`" + newName + "`",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        // 4) 裸标识符 old -> new（按标识符边界）
        // 标识符字符：字母/数字/下划线/中文/$/点（点用于 schema.table 处理；但这里替换整体 oldName）
        // 通过负向前后断言避免替换到更长单词内部
        s1 = System.Text.RegularExpressions.Regex.Replace(
            s1,
            @"(?<![A-Za-z0-9_\u4e00-\u9fa5\$])" + EscapeRegex(oldName) + @"(?![A-Za-z0-9_\u4e00-\u9fa5\$])",
            newName,
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        return s1;
    }

    private static bool ContainsSqlIdentifier(string sql, string name)
    {
        if (string.IsNullOrWhiteSpace(sql)) return false;
        name = (name ?? "").Trim();
        if (string.IsNullOrWhiteSpace(name)) return false;
        string EscapeRegex(string s) => System.Text.RegularExpressions.Regex.Escape(s);

        // [name]
        if (System.Text.RegularExpressions.Regex.IsMatch(sql, @"\[" + EscapeRegex(name) + @"\]", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            return true;
        // "name"
        if (System.Text.RegularExpressions.Regex.IsMatch(sql, "\"" + EscapeRegex(name) + "\"", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            return true;
        // `name`
        if (System.Text.RegularExpressions.Regex.IsMatch(sql, "`" + EscapeRegex(name) + "`", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            return true;
        // bare
        if (System.Text.RegularExpressions.Regex.IsMatch(sql,
            @"(?<![A-Za-z0-9_\u4e00-\u9fa5\$])" + EscapeRegex(name) + @"(?![A-Za-z0-9_\u4e00-\u9fa5\$])",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            return true;
        return false;
    }

    private static string ExtractSelectSqlFromCreateView(string createViewSql)
    {
        try
        {
            var s = (createViewSql ?? "").Trim();
            if (string.IsNullOrWhiteSpace(s)) return "";
            // sqlite_master.sql 通常是：CREATE VIEW [vw_xxx] AS SELECT ...
            var m = System.Text.RegularExpressions.Regex.Match(
                s,
                @"(?is)^\s*CREATE\s+VIEW\s+(?:IF\s+NOT\s+EXISTS\s+)?(.+?)\s+AS\s+(.*)$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (m.Success)
            {
                var body = (m.Groups[2].Value ?? "").Trim();
                return NormalizeSqlForView(body);
            }
            return NormalizeSqlForView(s);
        }
        catch
        {
            return NormalizeSqlForView(createViewSql);
        }
    }

    private sealed class JsonUpdateItem
    {
        public string Path { get; set; } = "";
        public string Kind { get; set; } = ""; // exact/sql
        public string Before { get; set; } = "";
        public string After { get; set; } = "";
    }

    private sealed class IniUpdateItem
    {
        public string ExcelPath { get; set; } = "";
        public string Scope { get; set; } = "";
        public string Path { get; set; } = "";
        public string Kind { get; set; } = "";
        public string Before { get; set; } = "";
        public string After { get; set; } = "";
    }

    private static (string UpdatedJson, List<JsonUpdateItem> Report) UpdateSchemeSettingsJsonReferences(
        string settingsJson,
        string oldName,
        string newName)
    {
        var report = new List<JsonUpdateItem>();
        try
        {
            if (string.IsNullOrWhiteSpace(settingsJson)) return (settingsJson, report);
            oldName = (oldName ?? "").Trim();
            newName = (newName ?? "").Trim();
            if (string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName)) return (settingsJson, report);

            JsonNode? root = JsonNode.Parse(settingsJson);
            if (root == null) return (settingsJson, report);

            bool IsTableKey(string? k)
            {
                if (string.IsNullOrWhiteSpace(k)) return false;
                var x = k!;
                return x.Contains("table", StringComparison.OrdinalIgnoreCase)
                       || x.Contains("view", StringComparison.OrdinalIgnoreCase)
                       || x.Contains("sheet", StringComparison.OrdinalIgnoreCase)
                       || x.Contains("main", StringComparison.OrdinalIgnoreCase)
                       || x.Contains("after", StringComparison.OrdinalIgnoreCase)
                       || x.Contains("before", StringComparison.OrdinalIgnoreCase)
                       || x.Contains("target", StringComparison.OrdinalIgnoreCase)
                       || x.Contains("output", StringComparison.OrdinalIgnoreCase)
                       || x.Contains("source", StringComparison.OrdinalIgnoreCase)
                       || x.Contains("ref", StringComparison.OrdinalIgnoreCase);
            }

            bool IsSqlKey(string? k)
            {
                if (string.IsNullOrWhiteSpace(k)) return false;
                return k!.Contains("sql", StringComparison.OrdinalIgnoreCase) || k.Contains("query", StringComparison.OrdinalIgnoreCase);
            }

            string ReplaceSqlRef(string s)
            {
                if (string.IsNullOrWhiteSpace(s)) return s;
                return ReplaceSqlIdentifier(s, oldName, newName);
            }

            void Walk(JsonNode? node, string path, string? propName)
            {
                if (node == null) return;
                if (node is JsonObject obj)
                {
                    foreach (var kv in obj.ToList())
                    {
                        Walk(kv.Value, $"{path}.{kv.Key}", kv.Key);
                    }
                    return;
                }
                if (node is JsonArray arr)
                {
                    for (int i = 0; i < arr.Count; i++)
                    {
                        Walk(arr[i], $"{path}[{i}]", propName);
                    }
                    return;
                }
                if (node is JsonValue val)
                {
                    if (!val.TryGetValue<string>(out var s) || s == null) return;
                    var before = s;
                    var after = before;

                    // 1) 精确匹配：常见“表名字段”
                    if (IsTableKey(propName) && string.Equals(before, oldName, StringComparison.OrdinalIgnoreCase))
                    {
                        after = newName;
                        report.Add(new JsonUpdateItem { Path = path, Kind = "exact", Before = before, After = after });
                    }

                    // 2) SQL 字段：替换引用
                    if (IsSqlKey(propName) && before.IndexOf(oldName, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        var x = ReplaceSqlRef(before);
                        if (!string.Equals(x, before, StringComparison.Ordinal))
                        {
                            after = x;
                            report.Add(new JsonUpdateItem { Path = path, Kind = "sql", Before = before.Length > 200 ? before[..200] + "..." : before, After = after.Length > 200 ? after[..200] + "..." : after });
                        }
                    }

                    // 3) 关键路径兜底：sqliteMainTableName
                    if (path.EndsWith(".sqliteMainTableName", StringComparison.OrdinalIgnoreCase) && string.Equals(before, oldName, StringComparison.OrdinalIgnoreCase))
                    {
                        after = newName;
                        report.Add(new JsonUpdateItem { Path = path, Kind = "exact", Before = before, After = after });
                    }

                    // 4) tableNameMap：工作表 -> 物理表名（key 是工作表名，不含 table 字样）
                    if (path.IndexOf(".tableNameMap.", StringComparison.OrdinalIgnoreCase) >= 0 && string.Equals(before, oldName, StringComparison.OrdinalIgnoreCase))
                    {
                        after = newName;
                        report.Add(new JsonUpdateItem { Path = path, Kind = "exact", Before = before, After = after });
                    }

                    if (!string.Equals(after, before, StringComparison.Ordinal))
                    {
                        // 写回
                        if (propName != null && node.Parent is JsonObject pobj)
                        {
                            pobj[propName] = after;
                        }
                        else
                        {
                            // 数组元素：通过 path 定位很麻烦；这里用 Parent 直接替换
                            if (node.Parent is JsonArray parr)
                            {
                                for (int i = 0; i < parr.Count; i++)
                                {
                                    if (ReferenceEquals(parr[i], node))
                                    {
                                        parr[i] = after;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            Walk(root, "$", null);

            var updated = root.ToJsonString(new System.Text.Json.JsonSerializerOptions { WriteIndented = false });
            // 去重 + 限制大小（避免过大消息）
            report = report
                .GroupBy(x => x.Path + "|" + x.Kind, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .Take(200)
                .ToList();
            return (updated, report);
        }
        catch
        {
            return (settingsJson, report);
        }
    }

    private List<IniUpdateItem> UpdateExcelSettingsIniForActiveScheme(string oldName, string newName)
    {
        var report = new List<IniUpdateItem>();
        try
        {
            if (string.IsNullOrWhiteSpace(_activeSchemeId)) return report;
            var schemeId = _activeSchemeId!;
            var iniPath = GetSchemeIniPath(schemeId);
            if (!File.Exists(iniPath)) return report;

            // 从“最新项目配置”提取 sources 列表（优先用已更新后的 settingsJson）
            var ini = ReadIni(iniPath);
            string settingsJson = "";
            if (ini.TryGetValue("settings", out var st) && st.TryGetValue("settingsJsonB64", out var sj) && !string.IsNullOrWhiteSpace(sj))
                settingsJson = FromB64(sj);
            if (string.IsNullOrWhiteSpace(settingsJson)) return report;

            var sources = ExtractSourcesFromSettings(settingsJson);
            var excelPaths = sources.Select(s => (s.FilePath ?? "").Replace('\\', '/'))
                .Where(x => !string.IsNullOrWhiteSpace(x) && (File.Exists(x.Replace('/', '\\')) || File.Exists(x)))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var excelPath0 in excelPaths)
            {
                var excelPath = excelPath0.Replace('\\', '/');
                string nativePath = excelPath.Replace('/', '\\');
                var p = File.Exists(nativePath) ? nativePath : excelPath;
                if (!File.Exists(p)) continue;

                var settingsIni = GetSettingsIniPath(p);
                if (!File.Exists(settingsIni)) continue;

                var secIni = ReadIni(settingsIni);
                bool changedAny = false;

                foreach (var kv in secIni.ToList())
                {
                    var scope = kv.Key ?? "";
                    if (string.IsNullOrWhiteSpace(scope)) continue;
                    if (!kv.Value.TryGetValue("json", out var b64) || string.IsNullOrWhiteSpace(b64)) continue;
                    string rawJson;
                    try { rawJson = Encoding.UTF8.GetString(Convert.FromBase64String(b64)); }
                    catch { continue; }
                    if (string.IsNullOrWhiteSpace(rawJson)) continue;

                    var r = UpdateSchemeSettingsJsonReferences(rawJson, oldName, newName);
                    if (!string.Equals(r.UpdatedJson, rawJson, StringComparison.Ordinal))
                    {
                        kv.Value["json"] = Convert.ToBase64String(Encoding.UTF8.GetBytes(r.UpdatedJson));
                        kv.Value["updatedAt"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        changedAny = true;
                    }

                    foreach (var it in (r.Report ?? new List<JsonUpdateItem>()).Take(50))
                    {
                        report.Add(new IniUpdateItem
                        {
                            ExcelPath = excelPath,
                            Scope = scope,
                            Path = it.Path,
                            Kind = it.Kind,
                            Before = it.Before,
                            After = it.After
                        });
                    }
                }

                if (changedAny)
                {
                    WriteIni(settingsIni, secIni);
                }
            }
        }
        catch
        {
            // 不阻塞主流程
        }

        return report.Take(300).ToList();
    }

    private void SyncDerivedViewsMetaFromDb(string schemeId)
    {
        try
        {
            if (_sqliteManager == null) return;
            if (string.IsNullOrWhiteSpace(schemeId)) return;

            var meta = LoadOrCreateProjectMeta(schemeId, displayName: null, settingsJson: null);
            meta.DerivedViews = meta.DerivedViews ?? new List<DerivedViewV1>();
            var now = DateTime.Now;

            // 仅同步 main 库中的 vw_ 视图（派生视图约定）
            var rows = _sqliteManager.Query("SELECT name, sql FROM main.sqlite_master WHERE type='view' AND name LIKE 'vw\\_%' ESCAPE '\\' ORDER BY name;");
            var dbMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in rows)
            {
                r.TryGetValue("name", out var nv);
                r.TryGetValue("sql", out var sv);
                var n = Convert.ToString(nv) ?? "";
                var s = Convert.ToString(sv) ?? "";
                if (string.IsNullOrWhiteSpace(n)) continue;
                dbMap[n.Trim()] = ExtractSelectSqlFromCreateView(s);
            }

            // 删除：meta 中存在但 DB 中不存在
            meta.DerivedViews.RemoveAll(v => v == null || string.IsNullOrWhiteSpace(v.Name) || !dbMap.ContainsKey(v.Name));

            // 更新/新增
            foreach (var kv in dbMap)
            {
                var name = kv.Key;
                var sql = kv.Value;
                var dv = meta.DerivedViews.FirstOrDefault(x => string.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase));
                if (dv == null)
                {
                    dv = new DerivedViewV1
                    {
                        Name = name,
                        Sql = sql,
                        Note = "",
                        Version = 1,
                        CreatedAt = now,
                        UpdatedAt = now,
                        DependsOn = GuessDependsOnTables(sql),
                        Enabled = true
                    };
                    meta.DerivedViews.Insert(0, dv);
                }
                else
                {
                    var changed = !string.Equals((dv.Sql ?? "").Trim(), (sql ?? "").Trim(), StringComparison.Ordinal);
                    dv.Sql = sql;
                    dv.DependsOn = GuessDependsOnTables(sql);
                    dv.Enabled = true;
                    if (changed)
                    {
                        dv.Version = Math.Max(1, dv.Version) + 1;
                        dv.UpdatedAt = now;
                    }
                }
            }

            meta.DerivedViewsUpdatedAt = now;
            meta.DbTables = GetDbTablesForScheme(schemeId);
            meta.DbSizeBytes = SafeFileSize(meta.DbPath);
            meta.UpdatedAt = now;
            SaveProjectMeta(meta);
        }
        catch
        {
            // 同步失败不阻塞主流程
        }
    }

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
        if (string.IsNullOrWhiteSpace(s)) s = "未命名项目";
        // Windows 文件名非法字符：\/:*?"<>|
        foreach (var ch in Path.GetInvalidFileNameChars())
            s = s.Replace(ch, '_');
        s = s.Replace('/', '_').Replace('\\', '_').Replace(':', '_').Replace('*', '_').Replace('?', '_').Replace('"', '_')
             .Replace('<', '_').Replace('>', '_').Replace('|', '_');
        s = s.Trim().TrimEnd('.'); // Windows 不允许尾部 '.'
        if (s.Length > maxLen) s = s.Substring(0, maxLen).Trim();
        if (string.IsNullOrWhiteSpace(s)) s = "未命名项目";
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

            // 若未记录 LastSchemeId，则默认取最近一个项目（提升“第二次打开自动加载”成功率）
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

    private object BuildProjectListPayload()
    {
        // 当前实现：基于 recent.json 的“最多保留 50 个项目”
        // 后续可扩展为扫描 schemes 目录以构建“所有项目”视图
        return _schemes
            .OrderByDescending(x => x.LastOpenTime)
            .Select(x =>
            {
                var meta = LoadOrCreateProjectMeta(x.Id, x.Name);
                var dbExists = !string.IsNullOrWhiteSpace(meta.DbPath) && File.Exists(meta.DbPath);
                var iniExists = !string.IsNullOrWhiteSpace(meta.IniPath) && File.Exists(meta.IniPath);
                // 表信息（轻量）：仅读取名称清单
                try
                {
                    meta.DbTables = GetDbTablesForScheme(x.Id);
                    if (meta.Sources != null && meta.Sources.Count > 0)
                    {
                        meta.MainSourceFile = meta.Sources.FirstOrDefault(s => s.IsMain)?.FilePath ?? meta.Sources.FirstOrDefault()?.FilePath;
                        var expected = ExpectedTableNamesFromSources(meta.Sources);
                        var actual = meta.DbTables ?? new List<string>();
                        int miss = expected.Count(t => !actual.Any(a => string.Equals(a, t, StringComparison.OrdinalIgnoreCase)));
                        int extra = actual.Count(t => !expected.Any(e => string.Equals(e, t, StringComparison.OrdinalIgnoreCase)));
                        int denom = Math.Max(1, expected.Union(actual, StringComparer.OrdinalIgnoreCase).Count());
                        meta.TableDiffRate = (double)(miss + extra) / denom;
                    }
                }
                catch { }
                return new
                {
                    id = x.Id,
                    name = x.Name,
                    lastOpenTime = x.LastOpenTime.ToString("yyyy-MM-dd HH:mm"),
                    dbPath = meta.DbPath,
                    dbSizeBytes = meta.DbSizeBytes,
                    updatedAt = meta.UpdatedAt.ToString("yyyy-MM-dd HH:mm"),
                    createdAt = meta.CreatedAt.ToString("yyyy-MM-dd HH:mm"),
                    openCount = meta.OpenCount,
                    iniPath = meta.IniPath,
                    mainSource = meta.MainSourceFile,
                    mainTable = meta.MainTableName,
                    dbTableCount = (meta.DbTables?.Count ?? 0),
                    tableDiffRate = meta.TableDiffRate,
                    derivedViewCount = (meta.DerivedViews?.Count ?? 0),
                    derivedViewsUpdatedAt = (meta.DerivedViewsUpdatedAt.HasValue ? meta.DerivedViewsUpdatedAt.Value.ToString("yyyy-MM-dd HH:mm") : ""),
                    sourcesCount = meta.Sources?.Count ?? 0,
                    status = (dbExists && iniExists) ? "ok" : (iniExists ? "db_missing" : "ini_missing")
                };
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

            // 优先从磁盘加载（便于“就近覆盖”快速迭代）；若不存在再回退到嵌入式资源
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
            var entryFileName = GetEntryHtmlFileNameForMode(_webBootUserMode);
            if (string.IsNullOrWhiteSpace(htmlPath))
                htmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, entryFileName);
            if (!File.Exists(htmlPath))
            {
                // 兼容调试目录：向上三级
                var htmlPath2 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", entryFileName);
                if (File.Exists(htmlPath2)) htmlPath = htmlPath2;
            }

            if (File.Exists(htmlPath))
            {
                // 优先走“虚拟域名映射”加载：比 file:// 更稳定（不会出现部分交互/权限策略怪异）
                _webHtmlPathCache = htmlPath;
                _webBaseDirCache = Path.GetDirectoryName(htmlPath);
                _webHtmlContentCache = ""; // 磁盘模式下不依赖缓存
                try { EnsureVirtualHostMapping(); } catch { }
                try { await EnsureBootUserModeScriptAsync(); } catch { }
                try
                {
                    var ts = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
                    // 统一走“两套入口壳”（normal/expert），由入口页跳转到 index.html?mode=...
                    var url = $"https://{WebVirtualHost}/{entryFileName}?ts={ts}&mode={_webBootUserMode}";
                    webView21.Source = new Uri(url);
                    System.Diagnostics.Debug.WriteLine($"成功从磁盘加载HTML文件（virtual host）: {htmlPath} -> {url}");
                    try { this.Text = $"ExcelSQLite - WEB:disk - {Path.GetFileName(htmlPath)} - {_webBootUserMode}"; } catch { }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"从磁盘加载HTML失败，将回退到嵌入式资源: {ex.Message}");
                    htmlPath = "";
                }
            }

            // 若磁盘加载失败/不存在，则回退到嵌入式资源（仅用于兜底）
            if (!File.Exists(htmlPath))
            {
                _webHtmlPathCache = null;
                _webBaseDirCache = null;
                try
                {
                    htmlContent = GetEmbeddedResource("ExcelSQLiteWeb.index.html");
                    System.Diagnostics.Debug.WriteLine("成功从嵌入式资源加载HTML文件");
                    try { this.Text = $"ExcelSQLite - WEB:embedded - index.html - {_webBootUserMode}"; } catch { }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"HTML加载失败: {ex.Message}");
                    htmlContent = $"<html><body><h1>错误</h1><p>无法找到index.html文件</p><p>当前目录: {AppDomain.CurrentDomain.BaseDirectory}</p></body></html>";
                    try { this.Text = $"ExcelSQLite - WEB:error - {_webBootUserMode}"; } catch { }
                }

                _webHtmlContentCache = htmlContent ?? "";
                // 通过 WebView2 的 DocumentCreated 脚本注入启动模式（比拼接 HTML 更可靠）
                try { await EnsureBootUserModeScriptAsync(); } catch { }
                webView21.NavigateToString(_webHtmlContentCache);
                System.Diagnostics.Debug.WriteLine("HTML内容加载完成（NavigateToString兜底）");
            }

            webView21.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;
            System.Diagnostics.Debug.WriteLine("WebMessageReceived事件注册成功");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"WebView2初始化异常: {ex.Message}");
        }
    }

    private async System.Threading.Tasks.Task EnsureBootUserModeScriptAsync()
    {
        if (webView21?.CoreWebView2 == null) return;
        try
        {
            if (!string.IsNullOrWhiteSpace(_webBootUserModeScriptId))
            {
                try { webView21.CoreWebView2.RemoveScriptToExecuteOnDocumentCreated(_webBootUserModeScriptId); } catch { }
            }
            var m = string.Equals(_webBootUserMode, "expert", StringComparison.OrdinalIgnoreCase) ? "expert" : "normal";
            _webBootUserModeScriptId = await webView21.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync($"window.__bootUserMode='{m}';");
        }
        catch { }
    }

    private void EnsureVirtualHostMapping()
    {
        try
        {
            if (webView21?.CoreWebView2 == null) return;
            if (string.IsNullOrWhiteSpace(_webBaseDirCache) || !Directory.Exists(_webBaseDirCache)) return;
            // 反复设置可能抛异常：忽略即可（映射已存在）
            webView21.CoreWebView2.SetVirtualHostNameToFolderMapping(
                WebVirtualHost,
                _webBaseDirCache,
                Microsoft.Web.WebView2.Core.CoreWebView2HostResourceAccessKind.Allow);
        }
        catch { }
    }

    private static string InjectBootUserMode(string html, string mode)
    {
        var m = string.Equals(mode, "expert", StringComparison.OrdinalIgnoreCase) ? "expert" : "normal";
        const string marker = "<!--BOOT_USER_MODE-->";
        var snippet = $"{marker}<script>window.__bootUserMode='{m}';</script>";
        if (string.IsNullOrWhiteSpace(html)) return snippet;

        try
        {
            var idx = html.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (idx >= 0)
            {
                var end = html.IndexOf("</script>", idx, StringComparison.OrdinalIgnoreCase);
                if (end > idx)
                {
                    end += "</script>".Length;
                    return html.Substring(0, idx) + snippet + html.Substring(end);
                }
            }

            var bodyIdx = html.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
            if (bodyIdx >= 0)
            {
                var bodyEnd = html.IndexOf(">", bodyIdx);
                if (bodyEnd > bodyIdx)
                {
                    bodyEnd += 1;
                    return html.Substring(0, bodyEnd) + snippet + html.Substring(bodyEnd);
                }
            }
        }
        catch { }
        // fallback：直接拼到最前面
        return snippet + html;
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
                case "startMetadataScanSqlite":
                    StartMetadataScanSqlite(data);
                    break;
                case "cancelMetadataScan":
                    CancelMetadataScan();
                    break;
                case "detectRelations":
                    DetectRelations(data);
                    break;
                case "saveMtjJson":
                    SaveMtjJson(data);
                    break;
                case "loadMtjJson":
                    LoadMtjJson(data);
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
                case "getSplitPreview":
                    GetSplitPreview(data);
                    break;
                case "getSqlResultSchema":
                    GetSqlResultSchema(data);
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
                case "initVaultDb":
                    InitVaultDb(data);
                    break;
                case "initPolicyRepoDb":
                    InitPolicyRepoDb(data);
                    break;
                case "seedDefaultTemplates":
                    SeedDefaultTemplates(data);
                    break;
                case "dsListTemplates":
                    DsListTemplates(data);
                    break;
                case "dsGetTemplate":
                    DsGetTemplate(data);
                    break;
                case "dsUpsertTemplate":
                    DsUpsertTemplate(data);
                    break;
                case "dsDeleteTemplate":
                    DsDeleteTemplate(data);
                    break;
                case "dsGetTableColumns":
                    DsGetTableColumns(data);
                    break;
                case "dsListAuditLogs":
                    DsListAuditLogs(data);
                    break;
                case "setOutputMode":
                    SetOutputMode(data);
                    break;
                case "dsGetRoutingStatus":
                    DsGetRoutingStatus(data);
                    break;
                case "dsReviewCoverage":
                    DsReviewCoverage(data);
                    break;
                case "dsReviewStrength":
                    DsReviewStrength(data);
                    break;
                case "dsReviewSampleScan":
                    DsReviewSampleScan(data);
                    break;
                case "dsBackupCreate":
                    DsBackupCreate(data);
                    break;
                case "dsBackupList":
                    DsBackupList(data);
                    break;
                case "dsBackupRestore":
                    DsBackupRestore(data);
                    break;
                case "dsBackupDelete":
                    DsBackupDelete(data);
                    break;
                case "dsBackupOpenFolder":
                    DsBackupOpenFolder(data);
                    break;
                case "dsSeedEnterpriseTemplates":
                    DsSeedEnterpriseTemplates(data);
                    break;
                case "dsTemplateMatchPreview":
                    DsTemplateMatchPreview(data);
                    break;
                case "dsEnumArchiveStart":
                    DsEnumArchiveStart(data);
                    break;
                case "dsEnumArchiveCancel":
                    DsEnumArchiveCancel(data);
                    break;
                case "dsEnumArchiveList":
                    DsEnumArchiveList(data);
                    break;
                case "dsEnumArchiveLoad":
                    DsEnumArchiveLoad(data);
                    break;
                case "dsEnumArchiveDelete":
                    DsEnumArchiveDelete(data);
                    break;
                case "dsEnumArchiveOpenFolder":
                    DsEnumArchiveOpenFolder(data);
                    break;
                case "getDesensitizationStatus":
                    GetDesensitizationStatus(data);
                    break;
                case "executeMaskJob":
                    ExecuteMaskJob(data);
                    break;
                case "detokenizeTableExport":
                    DetokenizeTableExport(data);
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
                            SendMessageToWebView(new { action = "error", message = "项目ID为空" });
                            break;
                        }
                        LoadSchemeInternal(schemeId, autoLoad: false);
                        break;
                    }
                case "getProjectList":
                    SendMessageToWebView(new
                    {
                        action = "projectList",
                        projects = BuildProjectListPayload(),
                        activeSchemeId = _activeSchemeId,
                        activeSchemeDbPath = _activeSchemeDbPath
                    });
                    break;
                case "getProjectDetail":
                    {
                        string schemeId = (data.TryGetProperty("schemeId", out var sid) ? sid.GetString() : null) ?? string.Empty;
                        if (string.IsNullOrWhiteSpace(schemeId))
                        {
                            SendMessageToWebView(new { action = "error", message = "项目ID为空" });
                            break;
                        }
                        var sm = _schemes.FirstOrDefault(x => string.Equals(x.Id, schemeId, StringComparison.OrdinalIgnoreCase));
                        var meta = LoadOrCreateProjectMeta(schemeId, sm?.Name ?? schemeId);
                        try
                        {
                            meta.DbTables = GetDbTablesForScheme(schemeId);
                            meta.DbSizeBytes = SafeFileSize(meta.DbPath);
                            if (meta.Sources != null && meta.Sources.Count > 0)
                                meta.MainSourceFile = meta.Sources.FirstOrDefault(s => s.IsMain)?.FilePath ?? meta.Sources.FirstOrDefault()?.FilePath;
                        }
                        catch { }
                        SendMessageToWebView(new { action = "projectDetail", schemeId = schemeId, meta = meta });
                        break;
                    }
                case "deleteScheme":
                    {
                        string schemeId = (data.TryGetProperty("schemeId", out var sid) ? sid.GetString() : null) ?? string.Empty;
                        if (string.IsNullOrWhiteSpace(schemeId))
                        {
                        SendMessageToWebView(new { action = "error", message = "删除失败：项目ID为空" });
                            break;
                        }
                        DeleteSchemeInternal(schemeId);
                        break;
                    }
                case "createRevisionPoint":
                    CreateRevisionPoint(data);
                    break;
                case "recordImportBatch":
                    RecordImportBatch(data);
                    break;
                case "upsertDerivedView":
                    UpsertDerivedView(data);
                    break;
                case "deleteDerivedView":
                    DeleteDerivedView(data);
                    break;
                case "mergeTables":
                    MergeTables(data);
                    break;
                case "archiveRawTables":
                    ArchiveRawTables(data);
                    break;
                case "selectBaseDb":
                    SelectAndAttachBaseDb();
                    break;
                case "detachBaseDb":
                    DetachBaseDb();
                    break;
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
                case "getSqliteIndexes":
                    GetSqliteIndexes(data);
                    break;
                case "createSqliteIndex":
                    CreateSqliteIndex(data);
                    break;
                case "dropSqliteIndex":
                    DropSqliteIndex(data);
                    break;
                case "createSqliteView":
                    CreateSqliteViewFromSql(data);
                    break;
                case "createSqliteTempTable":
                    CreateSqliteTempTableFromSql(data);
                    break;
                case "executeSqlScript":
                    ExecuteSqlScript(data);
                    break;
                case "getActiveFields":
                    GetActiveFields(data);
                    break;
                case "getTableList":
                    GetTableList();
                    break;
                case "getDbObjects":
                    GetDbObjects();
                    break;
                case "dropDbObjects":
                    DropDbObjects(data);
                    break;
                case "getDbObjectDependencies":
                    GetDbObjectDependencies(data);
                    break;
                case "renameDbObject":
                    RenameDbObject(data);
                    break;
                case "restoreRevisionPoint":
                    RestoreRevisionPoint(data);
                    break;
                case "appendAuditLog":
                    AppendAuditLog(data);
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
                    case "setUserModeAndReload":
                        SetUserModeAndReloadAsync(data);
                        break;
                    case "toggleUserMode":
                        ToggleUserModeFromHost();
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
                case "resetSqlConnection":
                    ResetSqlConnection();
                    break;
                case "cancelExport":
                    CancelExport();
                    break;
                case "countQueryRows":
                    CountQueryRows(data);
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
                case "openPath":
                    OpenPath(data);
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

    private async void SetUserModeAndReloadAsync(System.Text.Json.JsonElement data)
    {
        try
        {
            // 兼容：mode 可能在 payload.mode 或根节点 mode
            var payload = data.TryGetProperty("payload", out var pl) ? pl : data;
            var mode =
                (payload.TryGetProperty("mode", out var md) ? md.GetString() : null)
                ?? (data.TryGetProperty("mode", out var md2) ? md2.GetString() : null)
                ?? "normal";

            _webBootUserMode = string.Equals(mode, "expert", StringComparison.OrdinalIgnoreCase) ? "expert" : "normal";

            // 通过 DocumentCreated 脚本注入启动模式，然后整体 reload
            try { await EnsureBootUserModeScriptAsync(); } catch { }

            // 先给前端回执（避免“点了没反应”的错觉；同时便于定位宿主是否收到消息）
            try { SendMessageToWebView(new { action = "userModeReloading", mode = _webBootUserMode }); } catch { }

            // 优先走“虚拟域名映射”重新加载（避免 file:// 导致的交互/资源问题）
            if (!string.IsNullOrWhiteSpace(_webHtmlPathCache) && File.Exists(_webHtmlPathCache))
            {
                try { _webBaseDirCache = Path.GetDirectoryName(_webHtmlPathCache); } catch { }
                try { EnsureVirtualHostMapping(); } catch { }
                try
                {
                    var ts = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
                    var entryFileName = GetEntryHtmlFileNameForMode(_webBootUserMode);
                    webView21.Source = new Uri($"https://{WebVirtualHost}/{entryFileName}?ts={ts}&mode={_webBootUserMode}");
                }
                catch { }
            }
            else
            {
                var html = string.IsNullOrWhiteSpace(_webHtmlContentCache) ? "" : _webHtmlContentCache;
                try { webView21.NavigateToString(html); } catch { }
            }
        }
        catch (Exception ex)
        {
            try { SendMessageToWebView(new { action = "error", message = "切换用户模式失败：" + ex.Message }); } catch { }
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
            string schemeName = (data.TryGetProperty("schemeName", out var sn) ? sn.GetString() : null) ?? "未命名项目";
            var settingsEl = (data.TryGetProperty("settings", out var st) ? st : default);
            string settingsJson = settingsEl.ValueKind != System.Text.Json.JsonValueKind.Undefined ? settingsEl.GetRawText() : "{}";
            schemeName = string.IsNullOrWhiteSpace(schemeName) ? "未命名项目" : schemeName.Trim();

            // 项目文件名（安全化）：优先用传入 schemeId，否则由 schemeName 推导
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

            // 写入项目 ini（DataConfig/schemes/{项目名}.ini）
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

            // 同步 Project Hub 元数据（跟随业务库目录）
            try
            {
                var pm = LoadOrCreateProjectMeta(schemeId, schemeName, settingsJson);
                pm.DisplayName = schemeName;
                pm.DbPath = GetSchemeDbPath(schemeId);
                pm.IniPath = schemeIniPath;
                pm.LastOpenAt = DateTime.Now;
                pm.Sources = ExtractSourcesFromSettings(settingsJson);
                pm.MainSourceFile = pm.Sources.FirstOrDefault(s => s.IsMain)?.FilePath ?? pm.Sources.FirstOrDefault()?.FilePath;
                pm.MainTableName = ExtractMainTableNameFromSettings(settingsJson) ?? pm.MainTableName;
                SaveProjectMeta(pm);
            }
            catch { }

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

            // 保存项目后，立即切换到项目库（文件 SQLite），并尽量将当前内存库内容备份过去（便于“保存后继续用”）
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
            WriteErrorLog("保存项目失败", ex);
            SendMessageToWebView(new { action = "error", message = $"保存项目失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void LoadSchemeInternal(string schemeId, bool autoLoad)
    {
        // 读取项目 ini
        var schemeIniPath = GetSchemeIniPath(schemeId);
        if (!File.Exists(schemeIniPath))
        {
            SendMessageToWebView(new { action = "error", message = $"项目未找到: {schemeId}" });
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

        // 更新 recent.json（最近项目）
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

        // 同步 Project Hub 元数据（LastOpen/源文件/DB大小）
        try
        {
            var pm = LoadOrCreateProjectMeta(schemeId, displayName, settingsJson);
            pm.DisplayName = displayName;
            pm.LastOpenAt = DateTime.Now;
            pm.OpenCount = Math.Max(0, pm.OpenCount) + 1;
            pm.DbPath = _activeSchemeDbPath ?? GetSchemeDbPath(schemeId);
            pm.IniPath = schemeIniPath;
            pm.Sources = ExtractSourcesFromSettings(settingsJson);
            pm.MainTableName = ExtractMainTableNameFromSettings(settingsJson) ?? pm.MainTableName;
            SaveProjectMeta(pm);
        }
        catch { }

        // 用项目里的“主表文件”回填当前文件（便于后续字段/导入/打开原生文件等功能正常工作）
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
                    notice = "提示：项目中的部分源文件已发生变化（" + string.Join("，", changed.Take(5)) + (changed.Count > 5 ? "..." : "") + "），建议重新导入SQLite以保证一致性。";
                }
            }
        }
        catch { }

        // 打开/切换项目临时库（文件数据库）
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

    private void DeleteSchemeInternal(string schemeId)
    {
        try
        {
            // 不允许删除当前激活项目（避免误删正在用的库）
            if (!string.IsNullOrWhiteSpace(_activeSchemeId) && string.Equals(_activeSchemeId, schemeId, StringComparison.OrdinalIgnoreCase))
            {
                SendMessageToWebView(new { action = "error", message = "不允许删除当前正在使用的项目。请先切换到其他项目。"});
                return;
            }

            string iniPath = GetSchemeIniPath(schemeId);
            string dbPath = GetSchemeDbPath(schemeId);
            string metaPath = GetProjectMetaPath(schemeId);

            bool iniOk = true, dbOk = true, metaOk = true;
            try { if (File.Exists(iniPath)) File.Delete(iniPath); } catch { iniOk = false; }
            try { if (File.Exists(dbPath)) File.Delete(dbPath); } catch { dbOk = false; }
            try { if (File.Exists(metaPath)) File.Delete(metaPath); } catch { metaOk = false; }

            _schemes.RemoveAll(s => string.Equals(s.Id, schemeId, StringComparison.OrdinalIgnoreCase));
            SaveRecentSchemesToDisk();

            SendMessageToWebView(new
            {
                action = "schemeDeleted",
                schemeId = schemeId,
                ok = iniOk && dbOk && metaOk,
                recentSchemes = BuildRecentSchemesPayload(),
                projects = BuildProjectListPayload()
            });
        }
        catch (Exception ex)
        {
            WriteErrorLog("删除项目失败", ex);
            SendMessageToWebView(new { action = "error", message = $"删除项目失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void CreateRevisionPoint(System.Text.Json.JsonElement data)
    {
        try
        {
            string type = (data.TryGetProperty("type", out var tp) ? tp.GetString() : null) ?? "manual";
            string note = (data.TryGetProperty("note", out var nt) ? nt.GetString() : null) ?? "";
            bool enableBackup = true;
            if (data.TryGetProperty("enableBackup", out var eb) && (eb.ValueKind == System.Text.Json.JsonValueKind.True || eb.ValueKind == System.Text.Json.JsonValueKind.False))
                enableBackup = eb.ValueKind == System.Text.Json.JsonValueKind.True;

            var r = CreateRevisionPointInternal(type, note, enableBackup, sendMessage: true);
            if (!r.Ok)
                SendMessageToWebView(new { action = "revisionPointCreated", ok = false, message = r.Message });
        }
        catch (Exception ex)
        {
            WriteErrorLog("创建修订点失败", ex);
            SendMessageToWebView(new { action = "revisionPointCreated", ok = false, message = $"创建修订点失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private (bool Ok, RevisionPointV1? Rp, ProjectMetaV1? Meta, string Message) CreateRevisionPointInternal(string type, string note, bool enableBackup, bool sendMessage)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(_activeSchemeId) || string.IsNullOrWhiteSpace(_activeSchemeDbPath))
                return (false, null, null, "未绑定项目：无法创建修订点");
            if (_sqliteManager?.Connection == null)
                return (false, null, null, "SQLite未就绪：无法创建修订点");

            var schemeId = _activeSchemeId!;
            var meta = LoadOrCreateProjectMeta(schemeId, displayName: null, settingsJson: null);
            meta.DbPath = _activeSchemeDbPath!;
            meta.IniPath = GetSchemeIniPath(schemeId);
            meta.DbSizeBytes = SafeFileSize(meta.DbPath);

            var rp = new RevisionPointV1
            {
                Id = Guid.NewGuid().ToString("N"),
                Time = DateTime.Now,
                Type = string.IsNullOrWhiteSpace(type) ? "manual" : type,
                Note = note ?? ""
            };

            bool doBackup = meta.RevisionPolicy?.EnableSqliteBackup ?? true;
            doBackup = doBackup && enableBackup;

            if (doBackup)
            {
                var dir = Path.GetDirectoryName(_activeSchemeDbPath) ?? GetDbDir();
                var backupDir = Path.Combine(dir, "backups", schemeId);
                Directory.CreateDirectory(backupDir);
                var file = $"{schemeId}-rev-{DateTime.Now:yyyyMMdd-HHmmss}.db";
                var backupPath = Path.Combine(backupDir, file);
                using var dest = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={backupPath}");
                dest.Open();
                _sqliteManager.Connection.BackupDatabase(dest);
                dest.Close();
                rp.DbBackupPath = backupPath;
                rp.DbBackupBytes = SafeFileSize(backupPath);
            }

            meta.RevisionPoints = meta.RevisionPoints ?? new List<RevisionPointV1>();
            meta.RevisionPoints.Insert(0, rp);
            meta.UpdatedAt = DateTime.Now;

            try { CleanupRevisionBackups(meta); } catch { }
            SaveProjectMeta(meta);

            if (sendMessage)
            {
                try { SendMessageToWebView(new { action = "revisionPointCreated", ok = true, revisionPoint = rp, meta = meta }); } catch { }
            }

            return (true, rp, meta, "");
        }
        catch (Exception ex)
        {
            return (false, null, null, ex.Message);
        }
    }

    private void CleanupRevisionBackups(ProjectMetaV1 meta)
    {
        if (meta == null || string.IsNullOrWhiteSpace(meta.Id)) return;
        var keep = meta.RevisionPolicy?.KeepLastN ?? 10;
        keep = Math.Max(1, Math.Min(200, keep));
        var maxBytes = meta.RevisionPolicy?.MaxTotalBytes ?? (1024L * 1024L * 1024L);
        maxBytes = Math.Max(50L * 1024L * 1024L, maxBytes); // 最低 50MB

        var points = (meta.RevisionPoints ?? new List<RevisionPointV1>())
            .Where(p => p != null)
            .OrderByDescending(p => p.Time)
            .ToList();

        // 先按数量裁剪（删除多余备份文件）
        if (points.Count > keep)
        {
            foreach (var p in points.Skip(keep).ToList())
            {
                try { if (!string.IsNullOrWhiteSpace(p.DbBackupPath) && File.Exists(p.DbBackupPath)) File.Delete(p.DbBackupPath); } catch { }
            }
            points = points.Take(keep).ToList();
        }

        // 再按总大小裁剪
        long total = 0;
        foreach (var p in points)
        {
            if (!string.IsNullOrWhiteSpace(p.DbBackupPath) && File.Exists(p.DbBackupPath))
            {
                var sz = SafeFileSize(p.DbBackupPath);
                p.DbBackupBytes = sz;
                total += sz;
            }
        }
        if (total > maxBytes)
        {
            foreach (var p in points.OrderBy(p => p.Time).ToList()) // 从最旧开始删
            {
                if (total <= maxBytes) break;
                try
                {
                    if (!string.IsNullOrWhiteSpace(p.DbBackupPath) && File.Exists(p.DbBackupPath))
                    {
                        var sz = SafeFileSize(p.DbBackupPath);
                        File.Delete(p.DbBackupPath);
                        total -= sz;
                        p.DbBackupPath = null;
                        p.DbBackupBytes = null;
                    }
                }
                catch { }
            }
            // 移除“已被清理且无备份路径”的历史点（保留纯记录意义不大）
            points = points.Where(p => !string.IsNullOrWhiteSpace(p.DbBackupPath)).OrderByDescending(p => p.Time).ToList();
        }

        meta.RevisionPoints = points;
    }

    private void RecordImportBatch(System.Text.Json.JsonElement data)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(_activeSchemeId) || string.IsNullOrWhiteSpace(_activeSchemeDbPath))
            {
                SendMessageToWebView(new { action = "error", message = "未绑定项目：无法记录导入批次" });
                return;
            }
            var schemeId = _activeSchemeId!;
            var meta = LoadOrCreateProjectMeta(schemeId, displayName: null, settingsJson: null);
            meta.DbPath = _activeSchemeDbPath!;
            meta.IniPath = GetSchemeIniPath(schemeId);

            string batchId = (data.TryGetProperty("batchId", out var bid) ? bid.GetString() : null) ?? "";
            string importMode = (data.TryGetProperty("importMode", out var im) ? im.GetString() : null) ?? "text";
            bool cleared = (data.TryGetProperty("clearedBeforeImport", out var cb) && cb.ValueKind == System.Text.Json.JsonValueKind.True);
            int tables = 0;
            if (data.TryGetProperty("importedTables", out var it) && it.ValueKind == System.Text.Json.JsonValueKind.Number)
                tables = it.GetInt32();

            meta.LastImport = new ImportBatchV1
            {
                BatchId = string.IsNullOrWhiteSpace(batchId) ? $"imp-{DateTime.Now:yyyyMMdd-HHmmss}" : batchId,
                Time = DateTime.Now,
                ImportMode = importMode,
                ClearedBeforeImport = cleared,
                ImportedTables = Math.Max(0, tables)
            };

            meta.MainTableName = MainTableNameOrDefault();
            meta.DbTables = GetDbTablesForScheme(schemeId);
            meta.DbSizeBytes = SafeFileSize(meta.DbPath);
            meta.UpdatedAt = DateTime.Now;
            SaveProjectMeta(meta);

            SendMessageToWebView(new { action = "importBatchRecorded", meta = meta });
        }
        catch (Exception ex)
        {
            WriteErrorLog("记录导入批次失败", ex);
            SendMessageToWebView(new { action = "error", message = $"记录导入批次失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void UpsertDerivedView(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }
            if (string.IsNullOrWhiteSpace(_activeSchemeId))
            {
                SendMessageToWebView(new { action = "error", message = "请先保存/加载项目后再创建派生视图" });
                return;
            }

            string name = (data.TryGetProperty("name", out var nm) ? nm.GetString() : null) ?? "";
            string sql = (data.TryGetProperty("sql", out var sq) ? sq.GetString() : null) ?? "";
            string note = (data.TryGetProperty("note", out var nt) ? nt.GetString() : null) ?? "";
            name = (name ?? "").Trim();
            sql = (sql ?? "").Trim();

            if (string.IsNullOrWhiteSpace(name) || !name.StartsWith("vw_", StringComparison.OrdinalIgnoreCase))
            {
                SendMessageToWebView(new { action = "error", message = "视图名称必须以 vw_ 开头" });
                return;
            }
            // 不允许危险字符（保持简单）
            foreach (var ch in Path.GetInvalidFileNameChars())
            {
                if (name.Contains(ch)) { SendMessageToWebView(new { action = "error", message = "视图名称包含非法字符" }); return; }
            }
            if (!System.Text.RegularExpressions.Regex.IsMatch(name, @"^[A-Za-z0-9_\u4e00-\u9fa5]+$"))
            {
                SendMessageToWebView(new { action = "error", message = "视图名称仅允许：中文/英文/数字/下划线" });
                return;
            }
            if (!IsSelectOnlySqlForView(sql))
            {
                SendMessageToWebView(new { action = "error", message = "派生视图仅允许 SELECT 查询（包含 WITH ... SELECT），且不允许多语句" });
                return;
            }
            sql = NormalizeSqlForView(sql);

            // 写入 SQLite（CREATE VIEW）
            _sqliteManager.Execute($"DROP VIEW IF EXISTS [{name}];");
            _sqliteManager.Execute($"CREATE VIEW [{name}] AS {sql};");

            // 写入 project.json
            var schemeId = _activeSchemeId!;
            var meta = LoadOrCreateProjectMeta(schemeId, displayName: null, settingsJson: null);
            meta.DerivedViews = meta.DerivedViews ?? new List<DerivedViewV1>();
            var now = DateTime.Now;
            var dv = meta.DerivedViews.FirstOrDefault(x => string.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase));
            if (dv == null)
            {
                dv = new DerivedViewV1
                {
                    Name = name,
                    Sql = sql,
                    Note = note ?? "",
                    Version = 1,
                    CreatedAt = now,
                    UpdatedAt = now,
                    DependsOn = GuessDependsOnTables(sql),
                    Enabled = true
                };
                meta.DerivedViews.Insert(0, dv);
            }
            else
            {
                dv.Sql = sql;
                dv.Note = note ?? "";
                dv.Version = Math.Max(1, dv.Version) + 1;
                dv.UpdatedAt = now;
                dv.DependsOn = GuessDependsOnTables(sql);
                dv.Enabled = true;
            }
            meta.DerivedViewsUpdatedAt = now;
            meta.DbTables = GetDbTablesForScheme(schemeId);
            meta.DbSizeBytes = SafeFileSize(meta.DbPath);
            meta.UpdatedAt = now;
            SaveProjectMeta(meta);

            SendMessageToWebView(new { action = "derivedViewUpserted", ok = true, name = name, meta = meta });
        }
        catch (Exception ex)
        {
            WriteErrorLog("创建/更新派生视图失败", ex);
            SendMessageToWebView(new { action = "error", message = $"创建/更新派生视图失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void DeleteDerivedView(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }
            if (string.IsNullOrWhiteSpace(_activeSchemeId))
            {
                SendMessageToWebView(new { action = "error", message = "请先保存/加载项目后再删除派生视图" });
                return;
            }

            string name = (data.TryGetProperty("name", out var nm) ? nm.GetString() : null) ?? "";
            name = (name ?? "").Trim();
            if (string.IsNullOrWhiteSpace(name) || !name.StartsWith("vw_", StringComparison.OrdinalIgnoreCase))
            {
                SendMessageToWebView(new { action = "error", message = "视图名称必须以 vw_ 开头" });
                return;
            }

            _sqliteManager.Execute($"DROP VIEW IF EXISTS [{name}];");

            var schemeId = _activeSchemeId!;
            var meta = LoadOrCreateProjectMeta(schemeId, displayName: null, settingsJson: null);
            meta.DerivedViews = meta.DerivedViews ?? new List<DerivedViewV1>();
            meta.DerivedViews.RemoveAll(x => string.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase));
            meta.DerivedViewsUpdatedAt = DateTime.Now;
            meta.DbTables = GetDbTablesForScheme(schemeId);
            meta.DbSizeBytes = SafeFileSize(meta.DbPath);
            meta.UpdatedAt = DateTime.Now;
            SaveProjectMeta(meta);

            SendMessageToWebView(new { action = "derivedViewDeleted", ok = true, name = name, meta = meta });
        }
        catch (Exception ex)
        {
            WriteErrorLog("删除派生视图失败", ex);
            SendMessageToWebView(new { action = "error", message = $"删除派生视图失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void MergeTables(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "mergeTablesComplete", ok = false, message = "SQLite管理器未初始化" });
                return;
            }
            string dest = (data.TryGetProperty("destTable", out var dt) ? dt.GetString() : null) ?? "";
            bool overwrite = !(data.TryGetProperty("overwrite", out var ow) && ow.ValueKind == System.Text.Json.JsonValueKind.False);
            dest = dest.Trim();
            if (string.IsNullOrWhiteSpace(dest))
            {
                SendMessageToWebView(new { action = "mergeTablesComplete", ok = false, message = "输出表名为空" });
                return;
            }
            if (dest.StartsWith("vw_", StringComparison.OrdinalIgnoreCase))
            {
                SendMessageToWebView(new { action = "mergeTablesComplete", ok = false, message = "输出表名不允许以 vw_ 开头（该前缀保留给派生视图）" });
                return;
            }

            var src = new List<string>();
            if (data.TryGetProperty("sourceTables", out var st) && st.ValueKind == System.Text.Json.JsonValueKind.Array)
            {
                foreach (var x in st.EnumerateArray())
                {
                    var s = x.GetString() ?? "";
                    if (!string.IsNullOrWhiteSpace(s)) src.Add(s.Trim());
                }
            }
            src = src.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (src.Count == 0)
            {
                SendMessageToWebView(new { action = "mergeTablesComplete", ok = false, message = "来源表为空" });
                return;
            }

            // 1) 计算列并集（全部用 TEXT，避免类型冲突）
            var unionCols = new List<string>();
            var unionSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var colsOf = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (var t in src)
            {
                var cols = _sqliteManager.GetTableColumns(t);
                colsOf[t] = cols;
                foreach (var c in cols)
                {
                    if (unionSet.Add(c)) unionCols.Add(c);
                }
            }
            if (unionCols.Count == 0)
            {
                SendMessageToWebView(new { action = "mergeTablesComplete", ok = false, message = "无法读取来源表字段" });
                return;
            }

            using var txn = _sqliteManager.Connection!.BeginTransaction();
            try
            {
                if (overwrite)
                    _sqliteManager.Execute($"DROP TABLE IF EXISTS {SqliteManager.QuoteIdent(dest)};", null, txn);

                // 创建目标表（若已存在且 overwrite=false，则保证缺失列也补齐）
                if (!_sqliteManager.TableExists(dest))
                {
                    var defs = string.Join(", ", unionCols.Select(c => $"{SqliteManager.QuoteIdent(c)} TEXT"));
                    _sqliteManager.Execute($"CREATE TABLE {SqliteManager.QuoteIdent(dest)} ({defs});", null, txn);
                }
                else
                {
                    // 补齐缺失列
                    var existing = _sqliteManager.GetTableColumns(dest);
                    var ex = new HashSet<string>(existing, StringComparer.OrdinalIgnoreCase);
                    foreach (var c in unionCols)
                    {
                        if (ex.Contains(c)) continue;
                        _sqliteManager.Execute($"ALTER TABLE {SqliteManager.QuoteIdent(dest)} ADD COLUMN {SqliteManager.QuoteIdent(c)} TEXT;", null, txn);
                    }
                }

                // 2) 逐表 INSERT，缺失列填 NULL
                foreach (var t in src)
                {
                    var cols = colsOf[t];
                    var set = new HashSet<string>(cols, StringComparer.OrdinalIgnoreCase);
                    var selectList = string.Join(", ", unionCols.Select(c => set.Contains(c) ? SqliteManager.QuoteIdent(c) : $"NULL AS {SqliteManager.QuoteIdent(c)}"));
                    var colList = string.Join(", ", unionCols.Select(c => SqliteManager.QuoteIdent(c)));
                    _sqliteManager.Execute($"INSERT INTO {SqliteManager.QuoteIdent(dest)} ({colList}) SELECT {selectList} FROM {SqliteManager.QuoteIdent(t)};", null, txn);
                }

                txn.Commit();
            }
            catch
            {
                txn.Rollback();
                throw;
            }

            // 刷新元数据（可选）
            try
            {
                if (!string.IsNullOrWhiteSpace(_activeSchemeId))
                {
                    var meta = LoadOrCreateProjectMeta(_activeSchemeId!, displayName: null, settingsJson: null);
                    meta.DbTables = GetDbTablesForScheme(_activeSchemeId!);
                    meta.DbSizeBytes = SafeFileSize(meta.DbPath);
                    meta.UpdatedAt = DateTime.Now;
                    SaveProjectMeta(meta);
                }
            }
            catch { }

            SendMessageToWebView(new { action = "mergeTablesComplete", ok = true, destTable = dest });
        }
        catch (Exception ex)
        {
            WriteErrorLog("合并生成表失败", ex);
            SendMessageToWebView(new { action = "mergeTablesComplete", ok = false, message = ex.Message, hasErrorLog = true });
        }
    }

    private void ArchiveRawTables(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "rawArchiveComplete", ok = false, message = "SQLite管理器未初始化" });
                return;
            }
            if (string.IsNullOrWhiteSpace(_activeSchemeId) || string.IsNullOrWhiteSpace(_activeSchemeDbPath))
            {
                SendMessageToWebView(new { action = "rawArchiveComplete", ok = false, message = "请先保存/加载项目后再归档 raw 表" });
                return;
            }

            bool dropAfter = !(data.TryGetProperty("dropAfter", out var da) && da.ValueKind == System.Text.Json.JsonValueKind.False);

            // raw 表识别规则：raw__ 前缀（后续可扩展）
            var rawTables = new List<string>();
            try
            {
                var rows = _sqliteManager.Query("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'raw__%' ORDER BY name;");
                foreach (var r in rows)
                {
                    if (r.TryGetValue("name", out var v) && v != null && v != DBNull.Value)
                    {
                        var n = Convert.ToString(v) ?? "";
                        if (!string.IsNullOrWhiteSpace(n)) rawTables.Add(n);
                    }
                }
            }
            catch { }

            if (rawTables.Count == 0)
            {
                SendMessageToWebView(new { action = "rawArchiveComplete", ok = false, message = "未找到 raw__ 前缀的 raw 表（无需归档）" });
                return;
            }

            var schemeId = _activeSchemeId!;
            var dir = Path.GetDirectoryName(_activeSchemeDbPath) ?? GetDbDir();
            var backupDir = Path.Combine(dir, "backups", "raw-archive", schemeId);
            Directory.CreateDirectory(backupDir);
            var archivePath = Path.Combine(backupDir, $"raw-archive-{DateTime.Now:yyyyMMdd-HHmmss}.db");

            // 创建空库文件
            using (var conn = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={archivePath}"))
            {
                conn.Open();
                conn.Close();
            }

            string Esc(string s) => (s ?? "").Replace("'", "''");

            // 使用 ATTACH 复制 raw 表（结构简化：CREATE TABLE AS SELECT）
            _sqliteManager.Execute($"ATTACH DATABASE '{Esc(archivePath)}' AS arch;");
            foreach (var t in rawTables)
            {
                _sqliteManager.Execute($"DROP TABLE IF EXISTS arch.{SqliteManager.QuoteIdent(t)};");
                _sqliteManager.Execute($"CREATE TABLE arch.{SqliteManager.QuoteIdent(t)} AS SELECT * FROM {SqliteManager.QuoteIdent(t)};");
            }
            _sqliteManager.Execute("DETACH DATABASE arch;");

            int dropped = 0;
            if (dropAfter)
            {
                foreach (var t in rawTables)
                {
                    try
                    {
                        _sqliteManager.Execute($"DROP TABLE IF EXISTS {SqliteManager.QuoteIdent(t)};");
                        dropped++;
                    }
                    catch { }
                }
            }

            // 更新项目元数据（DbTables）
            try
            {
                var meta = LoadOrCreateProjectMeta(schemeId, displayName: null, settingsJson: null);
                meta.DbTables = GetDbTablesForScheme(schemeId);
                meta.DbSizeBytes = SafeFileSize(meta.DbPath);
                meta.UpdatedAt = DateTime.Now;
                SaveProjectMeta(meta);
            }
            catch { }

            SendMessageToWebView(new
            {
                action = "rawArchiveComplete",
                ok = true,
                archivePath = archivePath.Replace('\\', '/'),
                tables = rawTables,
                dropped = dropped
            });
        }
        catch (Exception ex)
        {
            WriteErrorLog("归档 raw 表失败", ex);
            SendMessageToWebView(new { action = "rawArchiveComplete", ok = false, message = ex.Message, hasErrorLog = true });
        }
    }

    private void SelectAndAttachBaseDb()
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }
            if (string.IsNullOrWhiteSpace(_activeSchemeId))
            {
                SendMessageToWebView(new { action = "error", message = "请先保存/加载项目后再挂载基础库" });
                return;
            }

            using var ofd = new OpenFileDialog();
            ofd.Filter = "SQLite 数据库 (*.db;*.sqlite)|*.db;*.sqlite|All files (*.*)|*.*";
            ofd.FilterIndex = 1;
            ofd.RestoreDirectory = true;
            ofd.Title = "选择基础库（共享维表/口径表）";
            if (ofd.ShowDialog() != DialogResult.OK) return;

            AttachBaseDbInternal(ofd.FileName);
        }
        catch (Exception ex)
        {
            WriteErrorLog("选择并挂载基础库失败", ex);
            SendMessageToWebView(new { action = "error", message = $"挂载基础库失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void AttachBaseDbInternal(string filePath)
    {
        if (_sqliteManager == null) return;
        if (string.IsNullOrWhiteSpace(_activeSchemeId)) return;
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            SendMessageToWebView(new { action = "error", message = "基础库文件不存在" });
            return;
        }

        var schemeId = _activeSchemeId!;
        var meta = LoadOrCreateProjectMeta(schemeId, displayName: null, settingsJson: null);
        var alias = string.IsNullOrWhiteSpace(meta.BaseDbAlias) ? "base" : meta.BaseDbAlias!;

        // 若已挂载过同名 alias，先 detach（忽略失败）
        try { _sqliteManager.DetachDatabase(alias); } catch { }

        _sqliteManager.AttachDatabase(filePath, alias);
        meta.BaseDbPath = filePath;
        meta.BaseDbAlias = alias;
        // 读取基础库表清单（不含 sqlite_ 系统表；含 view）
        try
        {
            var list = new List<string>();
            using var conn = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={filePath}");
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT name FROM sqlite_master WHERE type IN ('table','view') AND name NOT LIKE 'sqlite_%' ORDER BY name;";
            using var rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                var n = rd.IsDBNull(0) ? "" : (rd.GetString(0) ?? "");
                if (!string.IsNullOrWhiteSpace(n)) list.Add(n);
            }
            meta.BaseDbTables = list;
        }
        catch { meta.BaseDbTables = meta.BaseDbTables ?? new List<string>(); }
        meta.DbTables = GetDbTablesForScheme(schemeId);
        meta.DbSizeBytes = SafeFileSize(meta.DbPath);
        meta.UpdatedAt = DateTime.Now;
        SaveProjectMeta(meta);

        SendMessageToWebView(new { action = "baseDbAttached", ok = true, baseDbPath = filePath.Replace('\\', '/'), alias = alias, meta = meta });
        try { GetTableList(); } catch { }
    }

    private void DetachBaseDb()
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }
            if (string.IsNullOrWhiteSpace(_activeSchemeId))
            {
                SendMessageToWebView(new { action = "error", message = "请先保存/加载项目后再解除基础库" });
                return;
            }

            var schemeId = _activeSchemeId!;
            var meta = LoadOrCreateProjectMeta(schemeId, displayName: null, settingsJson: null);
            var alias = string.IsNullOrWhiteSpace(meta.BaseDbAlias) ? "base" : meta.BaseDbAlias!;
            try { _sqliteManager.DetachDatabase(alias); } catch { }
            meta.BaseDbPath = null;
            meta.BaseDbTables = new List<string>();
            meta.DbTables = GetDbTablesForScheme(schemeId);
            meta.DbSizeBytes = SafeFileSize(meta.DbPath);
            meta.UpdatedAt = DateTime.Now;
            SaveProjectMeta(meta);

            SendMessageToWebView(new { action = "baseDbAttached", ok = true, baseDbPath = "", alias = alias, meta = meta });
            try { GetTableList(); } catch { }
        }
        catch (Exception ex)
        {
            WriteErrorLog("解除基础库失败", ex);
            SendMessageToWebView(new { action = "error", message = $"解除基础库失败: {ex.Message}", hasErrorLog = true });
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

        // 自动挂载基础库（若 project meta 里配置了 BaseDbPath）
        try
        {
            if (!string.IsNullOrWhiteSpace(_activeSchemeId))
            {
                var meta = LoadOrCreateProjectMeta(_activeSchemeId!, displayName: null, settingsJson: null);
                if (!string.IsNullOrWhiteSpace(meta.BaseDbPath) && File.Exists(meta.BaseDbPath))
                {
                    var alias = string.IsNullOrWhiteSpace(meta.BaseDbAlias) ? "base" : meta.BaseDbAlias!;
                    try { _sqliteManager.DetachDatabase(alias); } catch { }
                    _sqliteManager.AttachDatabase(meta.BaseDbPath!, alias);
                    // 刷新基础库表清单（用于项目中心展示）
                    try
                    {
                        var list = new List<string>();
                        using var conn = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={meta.BaseDbPath}");
                        conn.Open();
                        using var cmd = conn.CreateCommand();
                        cmd.CommandText = "SELECT name FROM sqlite_master WHERE type IN ('table','view') AND name NOT LIKE 'sqlite_%' ORDER BY name;";
                        using var rd = cmd.ExecuteReader();
                        while (rd.Read())
                        {
                            var n = rd.IsDBNull(0) ? "" : (rd.GetString(0) ?? "");
                            if (!string.IsNullOrWhiteSpace(n)) list.Add(n);
                        }
                        meta.BaseDbTables = list;
                        SaveProjectMeta(meta);
                    }
                    catch { }
                }
            }
        }
        catch { }
    }

    private void CleanupTempDb()
    {
        try
        {
            // 说明：用户主动触发的“清理”，仅清理“孤儿库”（没有对应项目 ini 的 db）
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
                if (File.Exists(iniPath)) continue; // 有项目，不清理

                var ageDays = (DateTime.UtcNow - fi.LastWriteTimeUtc).TotalDays;
                if (ageDays >= orphanRetentionDays)
                {
                    try { fi.Delete(); deleted++; } catch { }
                }
            }

            SendMessageToWebView(new { action = "status", message = $"清理完成：扫描 {total} 个项目库，删除 {deleted} 个孤儿库（无对应项目且超期≥{orphanRetentionDays}天）" });
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
                SendMessageToWebView(new { action = "error", message = "未绑定项目：请先保存/加载项目后再重建数据库。" });
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
                message = $"项目数据库已重建：{dbPath}"
            });
        }
        catch (Exception ex)
        {
            WriteErrorLog("重建项目数据库失败", ex);
            SendMessageToWebView(new { action = "error", message = $"重建项目数据库失败: {ex.Message}", hasErrorLog = true });
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
            bool includeFields = !(data.TryGetProperty("includeFields", out var iff) && iff.ValueKind == System.Text.Json.JsonValueKind.False);

            // 取消上一次扫描
            try { _metadataScanCts?.Cancel(); } catch { }
            _metadataScanCts = new CancellationTokenSource();
            var token = _metadataScanCts.Token;

            var sw = Stopwatch.StartNew();
            var items = await Task.Run(() =>
            {
                var list = new List<object>();

                IEnumerable<string> files = Enumerable.Empty<string>();
                if (string.Equals(scope, "folder", StringComparison.OrdinalIgnoreCase))
                {
                    if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
                        throw new DirectoryNotFoundException("请选择有效的文件夹路径");
                    // 支持 xlsx / xls
                    files = Directory.EnumerateFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly)
                        .Concat(Directory.EnumerateFiles(folderPath, "*.xls", SearchOption.TopDirectoryOnly));
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
                    token.ThrowIfCancellationRequested();
                    idx++;
                    try
                    {
                        BeginInvoke(new Action(() =>
                        {
                            SendMessageToWebView(new { action = "metadataScanProgress", current = idx, total = total, fileName = Path.GetFileName(f), stage = "file" });
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
                                    // 字段列表（可选）
                                    List<string> fields = new();
                                    if (includeFields && _excelAnalyzer != null)
                                    {
                                        try
                                        {
                                            BeginInvoke(new Action(() =>
                                            {
                                                SendMessageToWebView(new { action = "metadataScanProgress", current = idx, total = total, fileName = Path.GetFileName(f), stage = "sheet", sheetName = s.Name });
                                            }));
                                        }
                                        catch { }
                                        try { fields = _excelAnalyzer.GetWorksheetFields(f, s.Name) ?? new List<string>(); } catch { fields = new List<string>(); }
                                    }

                                    sheets.Add(new { name = s.Name, rowCount = s.RowCount, colCount = s.ColCount, fields = fields });
                                    totalRows += s.RowCount;
                                }
                            }
                            else
                            {
                                foreach (var s in sheetNames)
                                {
                                    List<string> fields = new();
                                    if (includeFields && _excelAnalyzer != null)
                                    {
                                        try
                                        {
                                            BeginInvoke(new Action(() =>
                                            {
                                                SendMessageToWebView(new { action = "metadataScanProgress", current = idx, total = total, fileName = Path.GetFileName(f), stage = "sheet", sheetName = s });
                                            }));
                                        }
                                        catch { }
                                        try { fields = _excelAnalyzer.GetWorksheetFields(f, s) ?? new List<string>(); } catch { fields = new List<string>(); }
                                    }
                                    sheets.Add(new { name = s, fields = fields });
                                }
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
            }, token);
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
                includeRelations,
                includeFields,
                canceled = false
            };

            SendMessageToWebView(new { action = "metadataScanComplete", results = result });
        }
        catch (OperationCanceledException)
        {
            try
            {
                SendMessageToWebView(new { action = "metadataScanComplete", results = new { canceled = true, scopeLabel = "已取消", scanTime = 0, items = new List<object>() } });
            }
            catch { }
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"元数据扫描失败: {ex.Message}" });
        }
    }

    /// <summary>
    /// 元数据扫描（SQLite 数据源）：扫描当前已导入的表/视图，输出字段与样本统计
    /// 目标：让“元数据扫描”不是摆设（不依赖 ExcelAnalyzer）
    /// </summary>
    private async void StartMetadataScanSqlite(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null || !_sqliteManager.IsConnected)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite 未就绪：请先导入" });
                return;
            }

            bool includeStats = !(data.TryGetProperty("includeStats", out var ist) && ist.ValueKind == System.Text.Json.JsonValueKind.False);
            bool includeFields = !(data.TryGetProperty("includeFields", out var iff) && iff.ValueKind == System.Text.Json.JsonValueKind.False);
            int sampleRowsPerField = 1000;
            try
            {
                if (data.TryGetProperty("sampleRowsPerField", out var sr) && sr.TryGetInt32(out var n) && n > 0) sampleRowsPerField = n;
            }
            catch { sampleRowsPerField = 1000; }

            // 取消上一次扫描
            try { _metadataScanCts?.Cancel(); } catch { }
            _metadataScanCts = new CancellationTokenSource();
            var token = _metadataScanCts.Token;

            var sw = Stopwatch.StartNew();
            var items = await Task.Run(() =>
            {
                var list = new List<object>();
                var tables = _sqliteManager.GetTables() ?? new List<string>();
                int total = tables.Count;
                int idx = 0;

                // 采样：每表一次 SELECT * LIMIT N，避免“每列一条 SQL”带来的爆炸式耗时
                foreach (var t in tables)
                {
                    token.ThrowIfCancellationRequested();
                    idx++;
                    try
                    {
                        BeginInvoke(new Action(() =>
                        {
                            SendMessageToWebView(new { action = "metadataScanProgress", current = idx, total = total, fileName = t, stage = "table" });
                        }));
                    }
                    catch { }

                    try
                    {
                        // schema.table 支持：PRAGMA schema.table_info(table)
                        string schemaName = "main";
                        string tableNameOnly = t;
                        if (t.Contains("."))
                        {
                            var parts = t.Split('.', 2);
                            if (parts.Length == 2)
                            {
                                schemaName = parts[0];
                                tableNameOnly = parts[1];
                            }
                        }

                        var schemaSql = $"PRAGMA {SqliteManager.QuoteIdent(schemaName)}.table_info({SqliteManager.QuoteIdent(tableNameOnly)});";
                        var schemaRows = _sqliteManager.Query(schemaSql);
                        static string GetDictStr(Dictionary<string, object> d, string k)
                        {
                            try { return (d != null && d.TryGetValue(k, out var v) && v != null && v != DBNull.Value) ? Convert.ToString(v) ?? "" : ""; }
                            catch { return ""; }
                        }
                        var cols = schemaRows
                            .Select(r => new { name = GetDictStr(r, "name"), type = GetDictStr(r, "type") })
                            .Where(x => !string.IsNullOrWhiteSpace(x.name))
                            .ToList();

                        int colCount = cols.Count;
                        int rowCount = includeStats ? _sqliteManager.GetRowCount(t) : 0;

                        // 样本统计：非空率 / 样本唯一值 / 示例
                        Dictionary<string, (int nonNull, HashSet<string> uniq, string sample)> stat = new(StringComparer.OrdinalIgnoreCase);
                        int sampledRows = 0;
                        if (includeFields && sampleRowsPerField > 0 && colCount > 0)
                        {
                            foreach (var c in cols) stat[c.name] = (nonNull: 0, uniq: new HashSet<string>(), sample: "");
                            try
                            {
                                var sampleSql = $"SELECT * FROM {SqliteManager.QuoteIdent(t)} LIMIT {sampleRowsPerField};";
                                var sample = _sqliteManager.Query(sampleSql);
                                sampledRows = sample.Count;
                                foreach (var r in sample)
                                {
                                    foreach (var c in cols)
                                    {
                                        if (!r.ContainsKey(c.name)) continue;
                                        var v = r[c.name];
                                        if (v == null || v == DBNull.Value) continue;
                                        var s = Convert.ToString(v) ?? "";
                                        var tup = stat[c.name];
                                        tup.nonNull += 1;
                                        if (tup.uniq.Count < 2000) tup.uniq.Add(s);
                                        if (string.IsNullOrEmpty(tup.sample)) tup.sample = s;
                                        stat[c.name] = tup;
                                    }
                                }
                                // 若无行，保持 0
                            }
                            catch
                            {
                                // 某些 view 可能无法 SELECT *（例如依赖缺失），忽略样本
                            }
                        }

                        var fields = includeFields
                            ? cols.Select(c =>
                            {
                                // 注意：避免 ?: 导致命名元组退化为匿名元组（进而无法访问 .nonNull/.uniq/.sample）
                                (int nonNull, HashSet<string> uniq, string sample) s;
                                if (!stat.TryGetValue(c.name, out s))
                                    s = (nonNull: 0, uniq: new HashSet<string>(), sample: "");
                                double rate = 0;
                                try
                                {
                                    var denom = Math.Max(1, sampledRows);
                                    rate = Math.Round(s.nonNull * 100.0 / denom, 1);
                                }
                                catch { rate = 0; }
                                return new
                                {
                                    name = c.name,
                                    type = c.type,
                                    nonNullRate = rate,
                                    sampleDistinct = s.uniq.Count,
                                    sampleValue = s.sample
                                };
                            }).ToList<object>()
                            : new List<object>();

                        var sheets = new List<object>
                        {
                            new { name = t, rowCount = rowCount, colCount = colCount, fields = fields }
                        };

                        list.Add(new
                        {
                            fileName = t,
                            filePath = "",
                            sheetCount = 1,
                            totalRows = rowCount,
                            fileSize = "",
                            lastModified = "",
                            sheets = sheets
                        });
                    }
                    catch (Exception ex)
                    {
                        list.Add(new
                        {
                            fileName = t,
                            filePath = "",
                            sheetCount = 0,
                            totalRows = 0,
                            fileSize = "",
                            lastModified = "",
                            error = ex.Message
                        });
                    }
                }

                return list;
            }, token);
            sw.Stop();

            var result = new
            {
                scope = "sqlite",
                scopeLabel = "当前SQLite",
                scanTime = Math.Round(sw.Elapsed.TotalSeconds, 2),
                items = items,
                includeBasic = false,
                includeSheets = true,
                includeStats = includeStats,
                includeRelations = false,
                includeFields = includeFields,
                canceled = false
            };

            SendMessageToWebView(new { action = "metadataScanComplete", results = result });
        }
        catch (OperationCanceledException)
        {
            try
            {
                SendMessageToWebView(new { action = "metadataScanComplete", results = new { canceled = true, scope = "sqlite", scopeLabel = "已取消", scanTime = 0, items = new List<object>() } });
            }
            catch { }
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"SQLite元数据扫描失败: {ex.Message}" });
        }
    }

    private void CancelMetadataScan()
    {
        try
        {
            _metadataScanCts?.Cancel();
            SendMessageToWebView(new { action = "status", message = "已请求取消元数据扫描..." });
        }
        catch { }
    }

    // ==================== MTJ JSON（方案落盘） ====================
    private string GetSchemeMtjJsonPath(string schemeId)
        => Path.Combine(GetSchemesDir(), $"{schemeId}.mtj.json");

    private void SaveMtjJson(System.Text.Json.JsonElement data)
    {
        try
        {
            var schemeId = (data.TryGetProperty("schemeId", out var sid) ? sid.GetString() : null) ?? (_activeSchemeId ?? "");
            if (string.IsNullOrWhiteSpace(schemeId))
            {
                SendMessageToWebView(new { action = "error", message = "保存MTJ失败：未选择项目" });
                return;
            }
            if (!data.TryGetProperty("data", out var payload))
            {
                SendMessageToWebView(new { action = "error", message = "保存MTJ失败：data为空" });
                return;
            }
            EnsureDataConfigDirs();
            var path = GetSchemeMtjJsonPath(schemeId);
            var json = payload.GetRawText();
            File.WriteAllText(path, json, Encoding.UTF8);
            SendMessageToWebView(new { action = "mtjJsonSaved", message = $"已保存：{path}" });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"保存MTJ失败: {ex.Message}" });
        }
    }

    private void LoadMtjJson(System.Text.Json.JsonElement data)
    {
        try
        {
            var schemeId = (data.TryGetProperty("schemeId", out var sid) ? sid.GetString() : null) ?? (_activeSchemeId ?? "");
            if (string.IsNullOrWhiteSpace(schemeId)) return;
            EnsureDataConfigDirs();
            var path = GetSchemeMtjJsonPath(schemeId);
            if (!File.Exists(path))
            {
                SendMessageToWebView(new { action = "mtjJsonLoaded", data = (object?)null });
                return;
            }
            var json = File.ReadAllText(path, Encoding.UTF8);
            var obj = System.Text.Json.JsonSerializer.Deserialize<object>(json);
            SendMessageToWebView(new { action = "mtjJsonLoaded", data = obj });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"加载MTJ失败: {ex.Message}" });
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

                // ==================== 1) 命名相似性（候选筛选） ====================
                static string NormName(string s)
                {
                    s = (s ?? "").Trim().ToLowerInvariant();
                    s = s.Replace("_", " ").Replace("-", " ").Replace(".", " ").Replace("  ", " ");
                    return s;
                }
                static string CollapseName(string s)
                {
                    s = (s ?? "").Trim().ToLowerInvariant();
                    s = s.Replace("_", "").Replace("-", "").Replace(" ", "").Replace(".", "");
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

                static IEnumerable<string> Tokenize(string s)
                {
                    // 粗粒度：按非字母数字/中文切分；再对驼峰/数字边界切分
                    s = (s ?? "").Trim();
                    if (s.Length == 0) yield break;
                    var sb = new StringBuilder();
                    char prev = '\0';
                    List<string> tokens = new();
                    void Flush()
                    {
                        if (sb.Length == 0) return;
                        var t = sb.ToString().Trim();
                        sb.Clear();
                        if (!string.IsNullOrWhiteSpace(t)) tokens.Add(t);
                    }
                    foreach (var ch in s)
                    {
                        bool isSep = char.IsWhiteSpace(ch) || ch == '_' || ch == '-' || ch == '.' || ch == '/' || ch == '\\';
                        if (isSep)
                        {
                            Flush();
                            prev = ch;
                            continue;
                        }
                        bool boundary =
                            (prev != '\0')
                            && ((char.IsLetter(prev) && char.IsDigit(ch)) || (char.IsDigit(prev) && char.IsLetter(ch)))
                            || (char.IsLower(prev) && char.IsUpper(ch));
                        if (boundary)
                        {
                            Flush();
                        }
                        sb.Append(ch);
                        prev = ch;
                    }
                    Flush();
                    foreach (var t in tokens) yield return t;
                }

                static string CanonToken(string t)
                {
                    t = (t ?? "").Trim().ToLowerInvariant();
                    if (t.Length == 0) return "";
                    // 同义词归一（中英文）
                    return t switch
                    {
                        "id" or "no" or "code" or "key" or "pk" or "uuid" or "guid" or "编号" or "编码" or "代码" or "标识" => "id",
                        "name" or "nm" or "名称" or "姓名" => "name",
                        "date" or "dt" or "time" or "日期" or "时间" => "date",
                        "type" or "类别" or "类型" => "type",
                        "status" or "state" or "flag" or "状态" or "标志" => "status",
                        _ => t
                    };
                }

                static double TokenJaccard(string a, string b)
                {
                    var sa = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    var sb2 = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var t in Tokenize(a)) { var x = CanonToken(t); if (!string.IsNullOrWhiteSpace(x)) sa.Add(x); }
                    foreach (var t in Tokenize(b)) { var x = CanonToken(t); if (!string.IsNullOrWhiteSpace(x)) sb2.Add(x); }
                    if (sa.Count == 0 || sb2.Count == 0) return 0;
                    int inter = 0;
                    foreach (var x in sa) if (sb2.Contains(x)) inter++;
                    int uni = sa.Count + sb2.Count - inter;
                    return uni <= 0 ? 0 : (double)inter / uni;
                }

                static double NameScore(string colA, string colB)
                {
                    var jw = JaroWinkler(CollapseName(colA), CollapseName(colB));
                    var tj = TokenJaccard(NormName(colA), NormName(colB));
                    return Math.Max(jw, tj);
                }

                static bool IsKeyish(string col)
                {
                    var s = (col ?? "").Trim().ToLowerInvariant();
                    if (string.IsNullOrWhiteSpace(s)) return false;
                    if (s == "id" || s.EndsWith("id") || s.EndsWith("_id")) return true;
                    if (s.Contains("编号") || s.Contains("编码") || s.Contains("代码") || s.Contains("标识")) return true;
                    if (s.Contains("code") || s.Contains("no") || s.Contains("number")) return true;
                    return false;
                }

                var schemaMap = new Dictionary<string, List<(string Col, string Type)>>(StringComparer.OrdinalIgnoreCase);
                foreach (var t in tables)
                {
                    try { schemaMap[t] = mgr.GetTableSchema(t); }
                    catch { schemaMap[t] = new List<(string, string)>(); }
                }

                var colTypeMap = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
                foreach (var t in tables)
                {
                    var m0 = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var (col, typ) in schemaMap[t])
                        if (!string.IsNullOrWhiteSpace(col)) m0[col] = typ ?? "";
                    colTypeMap[t] = m0;
                }

                static string TypeGroup(string typ)
                {
                    var t = (typ ?? "").Trim().ToUpperInvariant();
                    if (t.Length == 0) return "unknown";
                    if (t.Contains("INT") || t.Contains("REAL") || t.Contains("NUM") || t.Contains("DEC") || t.Contains("DOUB") || t.Contains("FLOA"))
                        return "num";
                    if (t.Contains("DATE") || t.Contains("TIME"))
                        return "date";
                    if (t.Contains("CHAR") || t.Contains("TEXT") || t.Contains("CLOB") || t.Contains("VARCHAR"))
                        return "text";
                    if (t.Contains("BLOB")) return "blob";
                    return "other";
                }

                static bool IsTypeCompatible(string aType, string bType)
                {
                    var ga = TypeGroup(aType);
                    var gb = TypeGroup(bType);
                    if (ga == "unknown" || gb == "unknown") return true;
                    if (ga == gb) return true;
                    // 允许 text 与 other 的弱兼容（SQLite 类型宽松）
                    if ((ga == "text" && gb == "other") || (ga == "other" && gb == "text")) return true;
                    return false;
                }

                // ==================== 2) 值域同质性（枚举）分类：Exact / Left / Right / Contains ====================
                const double nameCandidateThreshold = 0.55;  // 命名相似性进入值域计算的阈值
                const int minDistinctForContains = 500;       // contains/left/right 的最大 distinct（超过则不算，防 N^2）
                const int lowCardinalityMax = 5;              // 低基数：常见 status/flag 噪音
                const double pkUniqThreshold = 0.98;          // 主键候选唯一性
                const double fkCoverThreshold = 0.95;         // 外键覆盖率（采样）
                const double exactStrongThreshold = 0.90;     // 完全匹配“强”阈值：达到就不再提 contains/left/right

                static string NormValSql(string colExpr)
                    => $"lower(replace(trim(CAST({colExpr} AS TEXT)),' ',''))";

                // profile cache
                var profileCache = new Dictionary<string, (long rows, long nn, long distinct, double uniq, double nullRate, int maxLen)>(StringComparer.OrdinalIgnoreCase);
                (long rows, long nn, long distinct, double uniq, double nullRate, int maxLen) GetProfile(string table, string col)
                {
                    var key = table + "||" + col;
                    if (profileCache.TryGetValue(key, out var p)) return p;
                    long rows = 0, nn = 0, dist = 0;
                    int maxLen = 0;
                    try
                    {
                        var qt = SqliteManager.QuoteIdent(table);
                        var qc = SqliteManager.QuoteIdent(col);
                        var norm = NormValSql(qc);
                        var sql = $@"
SELECT
  COUNT(*) AS rows,
  COUNT({qc}) AS nn,
  COUNT(DISTINCT {norm}) AS d,
  MAX(LENGTH({norm})) AS maxLen
FROM {qt};";
                        var r = mgr.Query(sql);
                        if (r.Count > 0)
                        {
                            var row = r[0];
                            row.TryGetValue("rows", out var vRows);
                            row.TryGetValue("nn", out var vNn);
                            row.TryGetValue("d", out var vD);
                            row.TryGetValue("maxLen", out var vMaxLen);
                            rows = Convert.ToInt64(vRows ?? 0);
                            nn = Convert.ToInt64(vNn ?? 0);
                            dist = Convert.ToInt64(vD ?? 0);
                            maxLen = Convert.ToInt32(vMaxLen ?? 0);
                        }
                    }
                    catch { }
                    double uniq = (nn > 0 && dist > 0) ? (double)dist / nn : 0;
                    double nullRate = rows > 0 ? 1.0 - (double)nn / rows : 1.0;
                    p = (rows, nn, dist, uniq, nullRate, maxLen);
                    profileCache[key] = p;
                    return p;
                }

                // distinct list cache（仅用于 contains/left/right；并且只在 distinct<=500 时加载全量）
                var distinctListCache = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
                List<string> GetDistinctList(string table, string col, int limit)
                {
                    var key = table + "||" + col + "||" + limit;
                    if (distinctListCache.TryGetValue(key, out var list)) return list;
                    list = new List<string>();
                    try
                    {
                        var qt = SqliteManager.QuoteIdent(table);
                        var qc = SqliteManager.QuoteIdent(col);
                        var norm = NormValSql(qc);
                        var rows = mgr.Query($"SELECT DISTINCT {norm} AS v FROM {qt} WHERE {qc} IS NOT NULL AND TRIM(CAST({qc} AS TEXT))<>'' LIMIT {Math.Max(1, limit)}");
                        foreach (var r in rows)
                        {
                            if (!r.TryGetValue("v", out var o) || o == null) continue;
                            var s = o.ToString();
                            if (string.IsNullOrWhiteSpace(s)) continue;
                            list.Add(s!.Trim().ToLowerInvariant().Replace(" ", ""));
                        }
                    }
                    catch { }
                    distinctListCache[key] = list;
                    return list;
                }

                (long inter, double coverA, double coverB) ExactDistinctOverlap(string aTable, string aCol, string bTable, string bCol)
                {
                    try
                    {
                        var qaT = SqliteManager.QuoteIdent(aTable);
                        var qaC = SqliteManager.QuoteIdent(aCol);
                        var qbT = SqliteManager.QuoteIdent(bTable);
                        var qbC = SqliteManager.QuoteIdent(bCol);
                        var na = NormValSql(qaC);
                        var nb = NormValSql(qbC);
                        var sql = $@"
WITH A AS (
  SELECT DISTINCT {na} AS v FROM {qaT} WHERE {qaC} IS NOT NULL AND TRIM(CAST({qaC} AS TEXT))<>'' 
),
B AS (
  SELECT DISTINCT {nb} AS v FROM {qbT} WHERE {qbC} IS NOT NULL AND TRIM(CAST({qbC} AS TEXT))<>'' 
)
SELECT COUNT(*) AS inter
FROM A INNER JOIN B ON A.v = B.v;";
                        var r = mgr.Query(sql);
                        long inter = 0;
                        if (r.Count > 0)
                        {
                            r[0].TryGetValue("inter", out var vInter);
                            inter = Convert.ToInt64(vInter ?? 0);
                        }
                        var pa = GetProfile(aTable, aCol);
                        var pb = GetProfile(bTable, bCol);
                        double coverA = pa.distinct > 0 ? (double)inter / pa.distinct : 0;
                        double coverB = pb.distinct > 0 ? (double)inter / pb.distinct : 0;
                        return (inter, coverA, coverB);
                    }
                    catch { return (0, 0, 0); }
                }

                double FkCoverSample(string pkTable, string pkCol, string fkTable, string fkCol, int sampleLimit)
                {
                    // fkCover = fk(非空) 中有多少能命中 pk distinct
                    try
                    {
                        var qPkT = SqliteManager.QuoteIdent(pkTable);
                        var qPkC = SqliteManager.QuoteIdent(pkCol);
                        var qFkT = SqliteManager.QuoteIdent(fkTable);
                        var qFkC = SqliteManager.QuoteIdent(fkCol);
                        var npk = NormValSql(qPkC);
                        var nfk = NormValSql(qFkC);
                        var lim = sampleLimit > 0 ? $" LIMIT {sampleLimit} " : "";
                        var sql = $@"
WITH PK AS (
  SELECT DISTINCT {npk} AS v
  FROM {qPkT}
  WHERE {qPkC} IS NOT NULL AND TRIM(CAST({qPkC} AS TEXT))<>'' 
),
FK AS (
  SELECT {nfk} AS v
  FROM {qFkT}
  WHERE {qFkC} IS NOT NULL AND TRIM(CAST({qFkC} AS TEXT))<>'' {lim}
)
SELECT
  COUNT(*) AS nn,
  SUM(CASE WHEN PK.v IS NOT NULL THEN 1 ELSE 0 END) AS hit
FROM FK
LEFT JOIN PK ON PK.v = FK.v;";
                        var r = mgr.Query(sql);
                        if (r.Count == 0) return 0;
                        var row = r[0];
                        row.TryGetValue("nn", out var vNn);
                        row.TryGetValue("hit", out var vHit);
                        double nn = Convert.ToDouble(vNn ?? 0);
                        if (nn <= 0) return 0;
                        double hit = Convert.ToDouble(vHit ?? 0);
                        return hit / nn;
                    }
                    catch { return 0; }
                }

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

                // 1) 候选字段对：命名相似 + 关键列兜底
                var candidates = new List<(string A, string aCol, string B, string bCol, double NameSim)>();
                if (useName)
                {
                    foreach (var (A, B) in pairs)
                    {
                        var aCols = schemaMap[A].Select(x => x.Col).ToList();
                        var bCols = schemaMap[B].Select(x => x.Col).ToList();
                        foreach (var ac in aCols)
                        foreach (var bc in bCols)
                        {
                            colTypeMap[A].TryGetValue(ac, out var at);
                            colTypeMap[B].TryGetValue(bc, out var bt);
                            if (!IsTypeCompatible(at ?? "", bt ?? ""))
                                continue;
                            var sim = NameScore(ac, bc);
                            if (sim >= nameCandidateThreshold)
                                candidates.Add((A, ac, B, bc, sim));
                        }
                    }
                }

                if (useValues)
                {
                    // 兜底：ID/CODE/编号/编码 两两配（限制规模）
                    foreach (var (A, B) in pairs.Take(30))
                    {
                        var aCols = schemaMap[A].Select(x => x.Col).Where(IsKeyish).Take(10).ToList();
                        var bCols = schemaMap[B].Select(x => x.Col).Where(IsKeyish).Take(10).ToList();
                        foreach (var ac in aCols)
                        foreach (var bc in bCols)
                            candidates.Add((A, ac, B, bc, NameScore(ac, bc)));
                    }
                }

                void AddResult(string leftTable, IEnumerable<string> leftCols, string rightTable, IEnumerable<string> rightCols,
                    string method, double score, double nameSim, double valueSim, double coverage, string? onSql,
                    string? enumCategory = null, double enumStrength = 0, double matchAtoB = 0, double matchBtoA = 0,
                    double uniqLeft = 0, double uniqRight = 0, double fkCover = 0, string? relationType = null)
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
                        onSql = onSql,
                        enumCategory,
                        enumStrength,
                        matchAtoB,
                        matchBtoA,
                        uniqLeft,
                        uniqRight,
                        fkCover,
                        relationType
                    };
                    resultList.Add((score, obj));
                }

                // 候选去重（同表对/同字段对只保留最大 nameSim）
                candidates = candidates
                    .GroupBy(x => $"{x.A}||{x.aCol}=>{x.B}||{x.bCol}", StringComparer.OrdinalIgnoreCase)
                    .Select(g => g.OrderByDescending(x => x.NameSim).First())
                    .OrderByDescending(x => x.NameSim)
                    .Take(Math.Max(200, maxPairs * 10))
                    .ToList();

                static string PairKey(string aT, string aC, string bT, string bC)
                {
                    // 忽略方向的签名：用于“同一对字段不重复出现”
                    var k1 = $"{aT}.{aC}";
                    var k2 = $"{bT}.{bC}";
                    return string.Compare(k1, k2, StringComparison.OrdinalIgnoreCase) <= 0 ? $"{k1}||{k2}" : $"{k2}||{k1}";
                }

                var reportedPairs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // ==================== 3) 计算值域相似性（分类去重：只输出最强类别） ====================
                if (useValues)
                {
                    foreach (var c in candidates)
                    {
                        // 先算完全匹配（distinct 交集覆盖率）
                        var (inter, coverA_exact, coverB_exact) = ExactDistinctOverlap(c.A, c.aCol, c.B, c.bCol);
                        var exactStrength = Math.Max(coverA_exact, coverB_exact);

                        // 唯一性/外键覆盖（方向性）
                        var pa = GetProfile(c.A, c.aCol);
                        var pb = GetProfile(c.B, c.bCol);
                        var fkBA = FkCoverSample(c.A, c.aCol, c.B, c.bCol, sampleRows); // B 引用 A
                        var fkAB = FkCoverSample(c.B, c.bCol, c.A, c.aCol, sampleRows); // A 引用 B

                        // 外键判定（强约束）
                        bool isFkBA = pa.uniq >= pkUniqThreshold && fkBA >= fkCoverThreshold && exactStrength >= minCoverage;
                        bool isFkAB = pb.uniq >= pkUniqThreshold && fkAB >= fkCoverThreshold && exactStrength >= minCoverage;

                        // 低基数降噪（status/flag）
                        bool lowCard = Math.Min(pa.distinct, pb.distinct) > 0 && Math.Min(pa.distinct, pb.distinct) <= lowCardinalityMax;
                        if (lowCard && c.NameSim < 0.80)
                        {
                            // 若命名证据不强，直接不报（避免 Status/Flag 噪音）
                            continue;
                        }

                        string enumCategory = "Exact";
                        double enumStrength = exactStrength;
                        double matchAtoB = coverA_exact;
                        double matchBtoA = coverB_exact;
                        string method = "枚举值-完全匹配";

                        // 如果完全匹配已经足够强，则不再算 contains/left/right（避免报告重复、也省性能）
                        if (exactStrength < exactStrongThreshold)
                        {
                            // contains/left/right：仅对小 distinct 的列计算（<=500）
                            if (pa.distinct > 0 && pb.distinct > 0
                                && pa.distinct <= minDistinctForContains && pb.distinct <= minDistinctForContains
                                && pa.maxLen <= 128 && pb.maxLen <= 128)
                            {
                                var aList = GetDistinctList(c.A, c.aCol, minDistinctForContains + 1);
                                var bList = GetDistinctList(c.B, c.bCol, minDistinctForContains + 1);
                                if (aList.Count <= minDistinctForContains && bList.Count <= minDistinctForContains && aList.Count > 0 && bList.Count > 0)
                                {
                                    static (double AtoB, double BtoA) Calc(List<string> A, List<string> B, Func<string, string, bool> pred)
                                    {
                                        int hitA = 0;
                                        foreach (var x in A)
                                        {
                                            bool ok = false;
                                            foreach (var y in B) { if (pred(x, y)) { ok = true; break; } }
                                            if (ok) hitA++;
                                        }
                                        int hitB = 0;
                                        foreach (var y in B)
                                        {
                                            bool ok = false;
                                            foreach (var x in A) { if (pred(y, x)) { ok = true; break; } }
                                            if (ok) hitB++;
                                        }
                                        return (A.Count > 0 ? (double)hitA / A.Count : 0, B.Count > 0 ? (double)hitB / B.Count : 0);
                                    }

                                    var (leftA, leftB) = Calc(aList, bList, (x, y) => y.StartsWith(x, StringComparison.OrdinalIgnoreCase));
                                    var (rightA, rightB) = Calc(aList, bList, (x, y) => y.EndsWith(x, StringComparison.OrdinalIgnoreCase));
                                    var (contA, contB) = Calc(aList, bList, (x, y) => y.Contains(x, StringComparison.OrdinalIgnoreCase));

                                    double leftStrength = Math.Max(leftA, leftB);
                                    double rightStrength = Math.Max(rightA, rightB);
                                    double contStrength = Math.Max(contA, contB);

                                    // Winner：强到弱（Exact > Left/Right > Contains），同强度时按优先级
                                    if (leftStrength >= rightStrength && leftStrength >= contStrength && leftStrength > enumStrength)
                                    {
                                        enumCategory = "Left";
                                        enumStrength = leftStrength;
                                        matchAtoB = leftA; matchBtoA = leftB;
                                        method = "枚举值-左包含";
                                    }
                                    else if (rightStrength >= leftStrength && rightStrength >= contStrength && rightStrength > enumStrength)
                                    {
                                        enumCategory = "Right";
                                        enumStrength = rightStrength;
                                        matchAtoB = rightA; matchBtoA = rightB;
                                        method = "枚举值-右包含";
                                    }
                                    else if (contStrength > enumStrength)
                                    {
                                        enumCategory = "Contains";
                                        enumStrength = contStrength;
                                        matchAtoB = contA; matchBtoA = contB;
                                        method = "枚举值-包含";
                                    }
                                }
                            }
                        }

                        // 最终阈值：按 Winner 强度过滤
                        if (enumStrength < minCoverage) continue;

                        // 判定类型 + 方向：外键优先（用于一键生成多表关联草稿）
                        string leftTable = c.A, rightTable = c.B;
                        string leftCol = c.aCol, rightCol = c.bCol;
                        double uniqLeft = pa.uniq, uniqRight = pb.uniq;
                        double fkCover = 0;
                        string relationType = "可能关联";
                        if (isFkBA && !isFkAB)
                        {
                            // B 引用 A：A 为主键端
                            relationType = "外键候选（B→A）";
                            leftTable = c.A; leftCol = c.aCol;
                            rightTable = c.B; rightCol = c.bCol;
                            uniqLeft = pa.uniq; uniqRight = pb.uniq;
                            fkCover = fkBA;
                        }
                        else if (isFkAB && !isFkBA)
                        {
                            relationType = "外键候选（A→B）";
                            // 调整输出方向：左=主键端（原 B），右=外键端（原 A）
                            leftTable = c.B; leftCol = c.bCol;
                            rightTable = c.A; rightCol = c.aCol;
                            uniqLeft = pb.uniq; uniqRight = pa.uniq;
                            fkCover = fkAB;
                        }
                        else if (lowCard)
                        {
                            relationType = "同域枚举（低基数）";
                        }

                        var score = Math.Min(100.0, enumStrength * 100 * 0.75 + c.NameSim * 100 * 0.15 + (fkCover > 0 ? 20 : 0));
                        var onSql = $"A.[{leftCol}] = B.[{rightCol}]";
                        AddResult(leftTable, new[] { leftCol }, rightTable, new[] { rightCol }, method, score, c.NameSim, enumStrength,
                            fkCover > 0 ? fkCover : enumStrength,
                            onSql,
                            enumCategory, enumStrength, matchAtoB, matchBtoA, uniqLeft, uniqRight, fkCover, relationType);
                        reportedPairs.Add(PairKey(leftTable, leftCol, rightTable, rightCol));
                    }
                }

                // ==================== 4) 仅命名相似（弱候选）：只有在“没有值域结果”或“该字段对未被值域覆盖”时输出 ====================
                if (useName)
                {
                    foreach (var c in candidates.Where(x => x.NameSim >= 0.80).OrderByDescending(x => x.NameSim).Take(Math.Max(20, maxPairs)))
                    {
                        var pk = PairKey(c.A, c.aCol, c.B, c.bCol);
                        if (reportedPairs.Contains(pk)) continue;
                        var score = c.NameSim * 100 * 0.60;
                        var onSql = $"A.[{c.aCol}] = B.[{c.bCol}]";
                        AddResult(c.A, new[] { c.aCol }, c.B, new[] { c.bCol }, "命名相似（仅元数据）", score, c.NameSim, 0, 0, onSql,
                            "Name", c.NameSim, 0, 0, 0, 0, 0, "可能关联（需数据验证）");
                        reportedPairs.Add(pk);
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
            ApplyExcelReportDefaults(ws);
            var headers = new[]
            {
                "置信度","分类","匹配强度","左表","左字段","右表","右字段","命名相似","A→B覆盖","B→A覆盖","FK覆盖","唯一性(左)","唯一性(右)","关系类型","匹配方式","推荐ON"
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
                string enumCategory = (el.TryGetProperty("enumCategory", out var ec) ? ec.GetString() : null) ?? "";
                double enumStrength = (el.TryGetProperty("enumStrength", out var es) && es.TryGetDouble(out var esv)) ? esv : 0;
                double matchAtoB = (el.TryGetProperty("matchAtoB", out var ab) && ab.TryGetDouble(out var abv)) ? abv : 0;
                double matchBtoA = (el.TryGetProperty("matchBtoA", out var ba) && ba.TryGetDouble(out var bav)) ? bav : 0;
                double fkCover = (el.TryGetProperty("fkCover", out var fk) && fk.TryGetDouble(out var fkv)) ? fkv : 0;
                double uniqLeft = (el.TryGetProperty("uniqLeft", out var ul) && ul.TryGetDouble(out var ulv)) ? ulv : 0;
                double uniqRight = (el.TryGetProperty("uniqRight", out var ur) && ur.TryGetDouble(out var urv)) ? urv : 0;
                string relationType = (el.TryGetProperty("relationType", out var rt0) ? rt0.GetString() : null) ?? "";

                ws.Cell(rowIdx, 1).Value = score;
                ws.Cell(rowIdx, 2).Value = enumCategory;
                ws.Cell(rowIdx, 3).Value = enumStrength;
                ws.Cell(rowIdx, 4).Value = leftTable;
                ws.Cell(rowIdx, 5).Value = leftCols;
                ws.Cell(rowIdx, 6).Value = rightTable;
                ws.Cell(rowIdx, 7).Value = rightCols;
                ws.Cell(rowIdx, 8).Value = nameSim;
                ws.Cell(rowIdx, 9).Value = matchAtoB;
                ws.Cell(rowIdx, 10).Value = matchBtoA;
                ws.Cell(rowIdx, 11).Value = fkCover;
                ws.Cell(rowIdx, 12).Value = uniqLeft;
                ws.Cell(rowIdx, 13).Value = uniqRight;
                ws.Cell(rowIdx, 14).Value = relationType;
                ws.Cell(rowIdx, 15).Value = method;
                ws.Cell(rowIdx, 16).Value = onSql;
                rowIdx++;
            }

            try { ApplyTextLeftNumberRightAlignment(ws, headerRow: 1); } catch { }
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
                ApplyExcelReportDefaults(ws);
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
                try { ApplyTextLeftNumberRightAlignment(ws, headerRow: 1); } catch { }
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
            ApplyExcelReportDefaults(ws1);
            ws1.Cell(1, 1).Value = "value";
            ws1.Cell(1, 1).Style.Font.Bold = true;
            ws1.Cell(1, 1).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#343a40");
            ws1.Cell(1, 1).Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
            for (int i = 0; i < lVals.Count && i < 1048575; i++) ws1.Cell(i + 2, 1).Value = lVals[i];
            try { ApplyTextLeftNumberRightAlignment(ws1, headerRow: 1); } catch { }
            ws1.Columns().AdjustToContents();

            var ws2 = wb.AddWorksheet(MakeSheetName($"{rTable}.{rCol}"));
            ApplyExcelReportDefaults(ws2);
            ws2.Cell(1, 1).Value = "value";
            ws2.Cell(1, 1).Style.Font.Bold = true;
            ws2.Cell(1, 1).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#343a40");
            ws2.Cell(1, 1).Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
            for (int i = 0; i < rVals.Count && i < 1048575; i++) ws2.Cell(i + 2, 1).Value = rVals[i];
            try { ApplyTextLeftNumberRightAlignment(ws2, headerRow: 1); } catch { }
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
            bool append = data.TryGetProperty("append", out var ap) && ap.ValueKind == System.Text.Json.JsonValueKind.True;
            string appendMode = (data.TryGetProperty("appendMode", out var am) ? am.GetString() : null) ?? "append"; // append/dedupe/upsert
            string keyColumnsRaw = (data.TryGetProperty("keyColumns", out var kc) ? kc.GetString() : null) ?? "";

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

            ImportResult result;

            // 追加策略：dedupe/upsert 需要“先导入到临时表，再合并到目标表”
            bool needStagingMerge =
                append
                && !resetDb
                && !dropExisting
                && !string.Equals(appendMode, "append", StringComparison.OrdinalIgnoreCase)
                && _sqliteManager.TableExists(tableName);

            static List<string> ParseKeyColumns(string raw)
            {
                var parts = (raw ?? "")
                    .Split(new[] { ',', '，', ';', '；', '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(x => (x ?? "").Trim())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                return parts;
            }

            static string ShortHash(string s)
            {
                try
                {
                    using var md5 = System.Security.Cryptography.MD5.Create();
                    var bs = md5.ComputeHash(System.Text.Encoding.UTF8.GetBytes(s ?? ""));
                    return string.Concat(bs.Take(6).Select(b => b.ToString("x2")));
                }
                catch { return Guid.NewGuid().ToString("N").Substring(0, 8); }
            }

            if (!needStagingMerge)
            {
                result = await _dataImporter.ImportWorksheetAsync(
                    filePath,
                    worksheetName,
                    tableName: tableName,
                    importMode: importMode,
                    progress: progress,
                    cancellationToken: CancellationToken.None,
                    append: (append && !resetDb && !dropExisting && string.Equals(appendMode, "append", StringComparison.OrdinalIgnoreCase)));
            }
            else
            {
                var keys = ParseKeyColumns(keyColumnsRaw);
                if (keys.Count == 0)
                {
                    SendMessageToWebView(new { action = "error", message = "去重/Upsert 需要指定键字段（逗号分隔）" });
                    return;
                }

                var stg = $"__stg_{ShortHash(tableName)}_{DateTime.Now:HHmmss}";
                // 1) 导入到临时表（覆盖写入）
                var stgRes = await _dataImporter.ImportWorksheetAsync(
                    filePath,
                    worksheetName,
                    tableName: stg,
                    importMode: importMode,
                    progress: progress,
                    cancellationToken: CancellationToken.None,
                    append: false);

                if (!stgRes.Success)
                {
                    result = stgRes;
                }
                else
                {
                    // 2) schema 演进：目标表缺失列 -> 自动加列（按你确认的策略 C）
                    var stgCols = _sqliteManager.GetTableColumns(stg);
                    var dstCols = _sqliteManager.GetTableColumns(tableName);
                    var dstSet = new HashSet<string>(dstCols, StringComparer.OrdinalIgnoreCase);
                    foreach (var c in stgCols)
                    {
                        if (dstSet.Contains(c)) continue;
                        try { _sqliteManager.Execute($"ALTER TABLE {SqliteManager.QuoteIdent(tableName)} ADD COLUMN {SqliteManager.QuoteIdent(c)} TEXT;"); } catch { }
                    }

                    // 3) 校验 key 列存在
                    var stgSet = new HashSet<string>(stgCols, StringComparer.OrdinalIgnoreCase);
                    foreach (var k in keys)
                    {
                        if (!stgSet.Contains(k))
                        {
                            _sqliteManager.Execute($"DROP TABLE IF EXISTS {SqliteManager.QuoteIdent(stg)};");
                            SendMessageToWebView(new { action = "error", message = $"去重/Upsert 失败：键字段不存在于源数据列中：{k}" });
                            return;
                        }
                    }

                    // 4) 确保唯一索引（用于 OR IGNORE / ON CONFLICT）
                    var uxName = $"ux_{ShortHash(tableName + "|" + string.Join(",", keys))}";
                    var keyExpr = string.Join(", ", keys.Select(k => SqliteManager.QuoteIdent(k)));
                    _sqliteManager.Execute($"CREATE UNIQUE INDEX IF NOT EXISTS {SqliteManager.QuoteIdent(uxName)} ON {SqliteManager.QuoteIdent(tableName)} ({keyExpr});");

                    // 5) 合并
                    var colList = string.Join(", ", stgCols.Select(c => SqliteManager.QuoteIdent(c)));
                    var selectCols = string.Join(", ", stgCols.Select(c => SqliteManager.QuoteIdent(c)));
                    int changed = 0;
                    if (string.Equals(appendMode, "dedupe", StringComparison.OrdinalIgnoreCase))
                    {
                        _sqliteManager.Execute($"INSERT OR IGNORE INTO {SqliteManager.QuoteIdent(tableName)} ({colList}) SELECT {selectCols} FROM {SqliteManager.QuoteIdent(stg)};");
                    }
                    else
                    {
                        var nonKey = stgCols.Where(c => !keys.Contains(c, StringComparer.OrdinalIgnoreCase)).ToList();
                        if (nonKey.Count == 0)
                        {
                            _sqliteManager.Execute($"INSERT OR IGNORE INTO {SqliteManager.QuoteIdent(tableName)} ({colList}) SELECT {selectCols} FROM {SqliteManager.QuoteIdent(stg)};");
                        }
                        else
                        {
                            var setSql = string.Join(", ", nonKey.Select(c => $"{SqliteManager.QuoteIdent(c)}=excluded.{SqliteManager.QuoteIdent(c)}"));
                            _sqliteManager.Execute($"INSERT INTO {SqliteManager.QuoteIdent(tableName)} ({colList}) SELECT {selectCols} FROM {SqliteManager.QuoteIdent(stg)} ON CONFLICT({keyExpr}) DO UPDATE SET {setSql};");
                        }
                    }
                    try
                    {
                        var r = _sqliteManager.Query("SELECT changes() AS c;");
                        if (r.Count > 0 && r[0].TryGetValue("c", out var v) && v != null && v != DBNull.Value)
                            changed = Convert.ToInt32(v);
                    }
                    catch { }

                    // 6) 清理临时表
                    try { _sqliteManager.Execute($"DROP TABLE IF EXISTS {SqliteManager.QuoteIdent(stg)};"); } catch { }

                    result = new ImportResult
                    {
                        FilePath = filePath,
                        WorksheetName = worksheetName,
                        TableName = tableName,
                        ImportMode = importMode,
                        Success = true,
                        RowCount = stgRes.RowCount,
                        ColumnCount = stgRes.ColumnCount,
                        ImportTime = stgRes.ImportTime,
                        ConversionStats = stgRes.ConversionStats,
                        Message = string.Equals(appendMode, "dedupe", StringComparison.OrdinalIgnoreCase)
                            ? $"去重合并完成：变更 {changed} 行（源 {stgRes.RowCount} 行）"
                            : $"Upsert 完成：变更 {changed} 行（源 {stgRes.RowCount} 行）"
                    };
                }
            }

            // 记录“当前主表”的真实表名（长期项目：不再强依赖 Main）
            if (string.Equals(role, "main", StringComparison.OrdinalIgnoreCase))
            {
                _currentMainTableName = tableName; // 目标表名
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
            int pageSize = (data.TryGetProperty("limit", out var lim) && lim.ValueKind == System.Text.Json.JsonValueKind.Number) ? lim.GetInt32() : 0;
            int offset = (data.TryGetProperty("offset", out var off) && off.ValueKind == System.Text.Json.JsonValueKind.Number) ? off.GetInt32() : 0;

            // SQL 参数（绑定参数）：支持 :p / @p / $p；也支持 {type,value} 显式类型
            object? sqlParams = null;
            try { sqlParams = ReadSqlParams(data); } catch { sqlParams = null; }

            // SQL性能保护（只影响 SQL 实验室/执行SQL页等“手工SQL”来源）：
            // 当用户没有显式传 limit/offset 时，默认只预览前 200 行，避免一次性加载几十万行导致 UI 假死。
            if (pageSize <= 0 && !string.IsNullOrWhiteSpace(source))
            {
                // 约定：所有 SQL 实验室相关 source 都以 sql- 开头（如 sql-editor/sql-generator/sql-editor-page）
                if (source.StartsWith("sql-", StringComparison.OrdinalIgnoreCase))
                {
                    pageSize = 200;
                    offset = 0;
                }
            }

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
                // 注意：当前项目使用文件 SQLite（{项目名}.db）以便复现与继承清洗结果；
                // SQL实验室写入被视为“实验动作”，由前端强硬风险提示+项目管理中的“重建/重导”兜底。

                var sw = Stopwatch.StartNew();
                _sqlLabTxn ??= _sqliteManager.Connection.BeginTransaction();
                _sqlExecCts?.Cancel();
                _sqlExecCts?.Dispose();
                _sqlExecCts = new CancellationTokenSource();
                int affected = await _sqliteManager.ExecuteAsync(sql, parameters: sqlParams, txn: _sqlLabTxn, timeoutSeconds: timeoutSeconds, cancellationToken: _sqlExecCts.Token);
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
            // 性能保护：允许前端传 limit/offset 做分页预览，避免一次性加载超大结果导致 WebView 假死
            // 仅对 SELECT/WITH/EXPLAIN 生效；PRAGMA 不做包裹
            string runSql = sql;
            try
            {
                var trimmed0 = (sql ?? string.Empty).TrimStart();
                bool isPageable =
                    trimmed0.StartsWith("select", StringComparison.OrdinalIgnoreCase)
                    || trimmed0.StartsWith("with", StringComparison.OrdinalIgnoreCase)
                    || trimmed0.StartsWith("explain", StringComparison.OrdinalIgnoreCase);
                if (isPageable && pageSize > 0)
                {
                    var sqlNoSemi = sql.Trim().TrimEnd(';');
                    var safeOffset = Math.Max(0, offset);
                    var safeLimit = Math.Max(1, pageSize);
                    runSql = $"SELECT * FROM ({sqlNoSemi}) t LIMIT {safeLimit} OFFSET {safeOffset}";
                }
            }
            catch { runSql = sql; }

            var result = await _queryEngine.ExecuteQueryAsync(runSql, txn: txn, parameters: sqlParams, timeoutSeconds: timeoutSeconds, cancellationToken: _sqlExecCts.Token);

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
                    sql = sql, // 返回原始 SQL（更符合用户认知）
                    executedSql = result.Sql, // 实际执行 SQL（可能带 LIMIT/OFFSET 包裹）
                    limit = pageSize,
                    offset = offset,
                    txnOpen = _sqlLabTxn != null
                }
            });
        }
        catch (OperationCanceledException)
        {
            SendMessageToWebView(new { action = "sqlCancelled", message = "已取消执行", requestId = requestId, source = source });
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
            try { _sqliteManager?.Interrupt(); } catch { }
            SendMessageToWebView(new { action = "sqlCancelAck" });
        }
        catch { }
    }

    private void ResetSqlConnection()
    {
        try
        {
            // 1) 先尽最大努力中断当前执行
            try { _sqlExecCts?.Cancel(); } catch { }
            try { _sqliteManager?.Interrupt(); } catch { }

            // 2) 回滚并关闭“实验事务”（避免连接处于 write txn 状态导致后续异常/锁）
            try
            {
                _sqlLabTxn?.Rollback();
            }
            catch { }
            try
            {
                _sqlLabTxn?.Dispose();
            }
            catch { }
            _sqlLabTxn = null;
            try { SendMessageToWebView(new { action = "sqlLabTxnState", txnOpen = false }); } catch { }

            // 3) 重建连接（文件库安全；内存库会丢数据——当前项目默认文件库）
            try { _sqliteManager?.Reopen(); } catch { }

            // 4) 刷新表列表（避免 UI 仍显示旧状态）
            try { GetTableList(); } catch { }

            SendMessageToWebView(new { action = "sqlResetDone", message = "已强制恢复：事务已回滚，连接已重建" });
        }
        catch (Exception ex)
        {
            try { SendMessageToWebView(new { action = "sqlResetDone", message = $"强制恢复失败: {ex.Message}" }); } catch { }
        }
    }

    private void CancelExport()
    {
        try
        {
            _exportCts?.Cancel();
        }
        catch { }
    }

    private static Dictionary<string, object?>? ReadSqlParams(System.Text.Json.JsonElement data)
    {
        try
        {
            if (!data.TryGetProperty("params", out var ps) || ps.ValueKind != System.Text.Json.JsonValueKind.Object)
                return null;

            object? ParseValue(System.Text.Json.JsonElement el)
            {
                try
                {
                    // 支持：{ type: "text|int|real|null|auto", value: ... }（前端参数面板可显式指定类型）
                    if (el.ValueKind == System.Text.Json.JsonValueKind.Object)
                    {
                        string tp = "";
                        if (el.TryGetProperty("type", out var t1) && t1.ValueKind == System.Text.Json.JsonValueKind.String) tp = (t1.GetString() ?? "");
                        else if (el.TryGetProperty("t", out var t2) && t2.ValueKind == System.Text.Json.JsonValueKind.String) tp = (t2.GetString() ?? "");
                        tp = tp.Trim().ToLowerInvariant();

                        System.Text.Json.JsonElement vv;
                        bool hasV =
                            (el.TryGetProperty("value", out vv))
                            || (el.TryGetProperty("v", out vv));

                        if (!hasV) return el.ToString();

                        if (tp == "null") return DBNull.Value;
                        if (tp == "text")
                        {
                            if (vv.ValueKind == System.Text.Json.JsonValueKind.String) return vv.GetString() ?? "";
                            return vv.ToString();
                        }
                        if (tp == "int")
                        {
                            if (vv.ValueKind == System.Text.Json.JsonValueKind.Number && vv.TryGetInt64(out var li)) return li;
                            if (long.TryParse(vv.ToString(), out var li2)) return li2;
                            return 0L;
                        }
                        if (tp == "real")
                        {
                            if (vv.ValueKind == System.Text.Json.JsonValueKind.Number) return vv.GetDouble();
                            if (double.TryParse(vv.ToString(), out var dd)) return dd;
                            return 0.0;
                        }
                        // auto：继续走默认推断
                        el = vv;
                    }

                    if (el.ValueKind == System.Text.Json.JsonValueKind.String)
                    {
                        var s0 = el.GetString();
                        if (string.IsNullOrEmpty(s0)) return DBNull.Value;
                        if (long.TryParse(s0, out var li)) return li;
                        if (double.TryParse(s0, out var dd)) return dd;
                        return s0;
                    }
                    if (el.ValueKind == System.Text.Json.JsonValueKind.Number)
                    {
                        if (el.TryGetInt64(out var li)) return li;
                        return el.GetDouble();
                    }
                    if (el.ValueKind == System.Text.Json.JsonValueKind.True || el.ValueKind == System.Text.Json.JsonValueKind.False)
                        return el.GetBoolean();
                    if (el.ValueKind == System.Text.Json.JsonValueKind.Null)
                        return DBNull.Value;

                    return el.ToString();
                }
                catch
                {
                    return el.ToString();
                }
            }

            var dict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in ps.EnumerateObject())
            {
                var key = (p.Name ?? "").Trim();
                if (string.IsNullOrWhiteSpace(key)) continue;
                dict[key] = ParseValue(p.Value);
            }
            return dict;
        }
        catch
        {
            return null;
        }
    }

    private void CountQueryRows(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }
            string sql = (data.TryGetProperty("sql", out var s) ? s.GetString() : null) ?? string.Empty;
            string requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
            string source = (data.TryGetProperty("source", out var src) ? src.GetString() : null) ?? "";
            object? sqlParams = null;
            try { sqlParams = ReadSqlParams(data); } catch { sqlParams = null; }
            if (string.IsNullOrWhiteSpace(sql))
            {
                SendMessageToWebView(new { action = "sqlCountComplete", requestId, source, count = 0L });
                return;
            }

            var sqlNoSemi = sql.Trim().TrimEnd(';');
            var countSql = $"SELECT COUNT(*) AS cnt FROM ({sqlNoSemi}) t";
            long cnt = 0;
            try
            {
                var rows = _sqliteManager.Query(countSql, sqlParams);
                if (rows.Count > 0 && rows[0].TryGetValue("cnt", out var v) && v != null) cnt = Convert.ToInt64(v);
            }
            catch { cnt = 0; }
            SendMessageToWebView(new { action = "sqlCountComplete", requestId, source, count = cnt });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"统计行数失败: {ex.Message}" });
        }
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
            ApplyExcelReportDefaults(wsMap);
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
            ApplyExcelReportDefaults(wsRules);
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

            try { ApplyTextLeftNumberRightAlignment(wsMap, headerRow: 1); } catch { }
            try { ApplyTextLeftNumberRightAlignment(wsRules, headerRow: 1); } catch { }
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
            ApplyExcelReportDefaults(wsSum);
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
                var v = sum[i].Item2;
                if (v is int iv) wsSum.Cell(i + 2, 2).Value = iv;
                else if (v is long lv) wsSum.Cell(i + 2, 2).Value = lv;
                else if (v is double dv) wsSum.Cell(i + 2, 2).Value = dv;
                else wsSum.Cell(i + 2, 2).Value = v?.ToString() ?? "";
            }
            wsSum.Columns().AdjustToContents();

            var wsF = wb.Worksheets.Add("字段差异");
            ApplyExcelReportDefaults(wsF);
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
            ApplyExcelReportDefaults(wsD);
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

            try { ApplyTextLeftNumberRightAlignment(wsSum, headerRow: 1); } catch { }
            try { ApplyTextLeftNumberRightAlignment(wsF, headerRow: 1); } catch { }
            try { ApplyTextLeftNumberRightAlignment(wsD, headerRow: 1); } catch { }
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

    private void GetSqliteIndexes(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }

            string tableName = (data.TryGetProperty("tableName", out var tn) ? tn.GetString() : null) ?? MainTableNameOrDefault();
            if (string.IsNullOrWhiteSpace(tableName))
            {
                SendMessageToWebView(new { action = "sqliteIndexesLoaded", tableName = "", indexes = Array.Empty<object>() });
                return;
            }

            var list = _sqliteManager.Query($"PRAGMA index_list({SqliteManager.QuoteIdent(tableName)});");
            var indexes = new List<object>();
            foreach (var r in list)
            {
                var name = Convert.ToString(r.TryGetValue("name", out var v) ? v : null) ?? "";
                if (string.IsNullOrWhiteSpace(name)) continue;
                bool unique = false;
                try { unique = Convert.ToInt32(r.TryGetValue("unique", out var u) ? u : 0) == 1; } catch { }

                List<string> cols = new();
                try
                {
                    var info = _sqliteManager.Query($"PRAGMA index_info({SqliteManager.QuoteIdent(name)});");
                    foreach (var rr in info)
                    {
                        var cn = Convert.ToString(rr.TryGetValue("name", out var nv) ? nv : null) ?? "";
                        if (!string.IsNullOrWhiteSpace(cn)) cols.Add(cn);
                    }
                }
                catch { }

                indexes.Add(new { name, unique, columns = cols });
            }

            SendMessageToWebView(new { action = "sqliteIndexesLoaded", tableName, indexes });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"获取索引失败: {ex.Message}" });
        }
    }

    private void CreateSqliteIndex(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }

            string tableName = (data.TryGetProperty("tableName", out var tn) ? tn.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(tableName))
            {
                SendMessageToWebView(new { action = "error", message = "创建索引失败：表名为空" });
                return;
            }

            bool unique = (data.TryGetProperty("unique", out var uq) && uq.ValueKind == System.Text.Json.JsonValueKind.True);
            string indexName = (data.TryGetProperty("indexName", out var ixn) ? ixn.GetString() : null) ?? "";

            List<string> columns = new();
            if (data.TryGetProperty("columns", out var cs) && cs.ValueKind == System.Text.Json.JsonValueKind.Array)
            {
                foreach (var c in cs.EnumerateArray())
                {
                    var s = c.GetString() ?? "";
                    if (!string.IsNullOrWhiteSpace(s)) columns.Add(s);
                }
            }
            if (columns.Count == 0)
            {
                SendMessageToWebView(new { action = "error", message = "创建索引失败：未选择字段" });
                return;
            }

            if (string.IsNullOrWhiteSpace(indexName))
            {
                static string San(string x)
                {
                    var s = new string((x ?? "").Select(ch => char.IsLetterOrDigit(ch) ? ch : '_').ToArray());
                    while (s.Contains("__")) s = s.Replace("__", "_");
                    return s.Trim('_');
                }
                var t0 = San(tableName);
                var c0 = string.Join("_", columns.Take(3).Select(San));
                indexName = $"idx_{t0}_{c0}";
                if (indexName.Length > 60) indexName = indexName.Substring(0, 60);
            }

            var colsSql = string.Join(", ", columns.Select(SqliteManager.QuoteIdent));
            var sql = $"CREATE {(unique ? "UNIQUE " : "")}INDEX IF NOT EXISTS {SqliteManager.QuoteIdent(indexName)} ON {SqliteManager.QuoteIdent(tableName)} ({colsSql});";
            _sqliteManager.Execute(sql);
            try { _sqliteManager.Execute($"ANALYZE {SqliteManager.QuoteIdent(tableName)};"); } catch { }

            SendMessageToWebView(new { action = "sqliteIndexCreated", tableName, indexName });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"创建索引失败: {ex.Message}" });
        }
    }

    private void DropSqliteIndex(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }
            string tableName = (data.TryGetProperty("tableName", out var tn) ? tn.GetString() : null) ?? "";
            string indexName = (data.TryGetProperty("indexName", out var ix) ? ix.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(indexName))
            {
                SendMessageToWebView(new { action = "error", message = "删除索引失败：索引名为空" });
                return;
            }

            _sqliteManager.Execute($"DROP INDEX IF EXISTS {SqliteManager.QuoteIdent(indexName)};");
            try { if (!string.IsNullOrWhiteSpace(tableName)) _sqliteManager.Execute($"ANALYZE {SqliteManager.QuoteIdent(tableName)};"); } catch { }
            SendMessageToWebView(new { action = "sqliteIndexDropped", tableName, indexName });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"删除索引失败: {ex.Message}" });
        }
    }

    private void CreateSqliteViewFromSql(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }
            string name = (data.TryGetProperty("name", out var n) ? n.GetString() : null) ?? "";
            string sql = (data.TryGetProperty("sql", out var s) ? s.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(sql))
            {
                SendMessageToWebView(new { action = "error", message = "创建视图失败：参数缺失" });
                return;
            }
            var sqlNoSemi = sql.Trim().TrimEnd(';');
            _sqliteManager.Execute($"CREATE VIEW {SqliteManager.QuoteIdent(name)} AS {sqlNoSemi};");
            try { GetTableList(); } catch { }
            SendMessageToWebView(new { action = "sqliteObjectCreated", objectType = "view", name });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"创建视图失败: {ex.Message}" });
        }
    }

    private void CreateSqliteTempTableFromSql(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }
            string name = (data.TryGetProperty("name", out var n) ? n.GetString() : null) ?? "";
            string sql = (data.TryGetProperty("sql", out var s) ? s.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(sql))
            {
                SendMessageToWebView(new { action = "error", message = "创建临时表失败：参数缺失" });
                return;
            }
            var sqlNoSemi = sql.Trim().TrimEnd(';');
            _sqliteManager.Execute($"CREATE TEMP TABLE {SqliteManager.QuoteIdent(name)} AS {sqlNoSemi};");
            try { GetTableList(); } catch { }
            SendMessageToWebView(new { action = "sqliteObjectCreated", objectType = "temp", name });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"创建临时表失败: {ex.Message}" });
        }
    }

    private static List<string> SplitSqlStatements(string sql)
    {
        var list = new List<string>();
        if (string.IsNullOrWhiteSpace(sql)) return list;
        var sb = new StringBuilder();
        bool inS = false, inD = false;
        bool inLine = false, inBlock = false;
        for (int i = 0; i < sql.Length; i++)
        {
            char ch = sql[i];
            char nx = (i + 1 < sql.Length) ? sql[i + 1] : '\0';

            if (inLine)
            {
                sb.Append(ch);
                if (ch == '\n') inLine = false;
                continue;
            }
            if (inBlock)
            {
                sb.Append(ch);
                if (ch == '*' && nx == '/')
                {
                    sb.Append(nx);
                    i++;
                    inBlock = false;
                }
                continue;
            }

            if (!inS && !inD)
            {
                if (ch == '-' && nx == '-')
                {
                    sb.Append(ch); sb.Append(nx); i++;
                    inLine = true;
                    continue;
                }
                if (ch == '/' && nx == '*')
                {
                    sb.Append(ch); sb.Append(nx); i++;
                    inBlock = true;
                    continue;
                }
            }

            if (ch == '\'' && !inD)
            {
                sb.Append(ch);
                if (inS && nx == '\'') { sb.Append(nx); i++; continue; } // '' escape
                inS = !inS;
                continue;
            }
            if (ch == '"' && !inS)
            {
                sb.Append(ch);
                if (inD && nx == '"') { sb.Append(nx); i++; continue; } // "" escape
                inD = !inD;
                continue;
            }

            if (ch == ';' && !inS && !inD)
            {
                var one = sb.ToString().Trim();
                if (!string.IsNullOrWhiteSpace(one)) list.Add(one);
                sb.Clear();
                continue;
            }

            sb.Append(ch);
        }
        var last = sb.ToString().Trim();
        if (!string.IsNullOrWhiteSpace(last)) list.Add(last);
        return list;
    }

    private async void ExecuteSqlScript(System.Text.Json.JsonElement data)
    {
        string requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        string source = (data.TryGetProperty("source", out var src) ? src.GetString() : null) ?? "sql-editor";
        try
        {
            if (_sqliteManager == null || _queryEngine == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite/查询引擎未初始化", requestId, source });
                return;
            }

            string sql = (data.TryGetProperty("sql", out var s) ? s.GetString() : null) ?? "";
            bool expertMode = !(data.TryGetProperty("expertMode", out var em) && em.ValueKind == System.Text.Json.JsonValueKind.False);
            bool allowDangerDdl = (data.TryGetProperty("allowDangerDdl", out var addl) && addl.ValueKind == System.Text.Json.JsonValueKind.True);
            bool confirmedDangerous = (data.TryGetProperty("confirmedDangerous", out var cd) && cd.ValueKind == System.Text.Json.JsonValueKind.True);
            int timeoutSeconds = (data.TryGetProperty("timeoutSeconds", out var ts) && ts.ValueKind == System.Text.Json.JsonValueKind.Number) ? ts.GetInt32() : 30;
            bool continueOnError = (data.TryGetProperty("continueOnError", out var coe) && coe.ValueKind == System.Text.Json.JsonValueKind.True);
            int previewLimit = (data.TryGetProperty("previewLimit", out var pl) && pl.ValueKind == System.Text.Json.JsonValueKind.Number) ? pl.GetInt32() : 200;
            previewLimit = Math.Max(0, Math.Min(5000, previewLimit));
            object? sqlParams = null;
            try { sqlParams = ReadSqlParams(data); } catch { sqlParams = null; }

            var stmts = SplitSqlStatements(sql);
            if (stmts.Count == 0)
            {
                SendMessageToWebView(new { action = "error", message = "脚本为空", requestId, source });
                return;
            }

            static bool IsSelectLike(string s0)
            {
                var t = (s0 ?? string.Empty).TrimStart();
                return t.StartsWith("select", StringComparison.OrdinalIgnoreCase)
                       || t.StartsWith("with", StringComparison.OrdinalIgnoreCase)
                       || t.StartsWith("pragma", StringComparison.OrdinalIgnoreCase)
                       || t.StartsWith("explain", StringComparison.OrdinalIgnoreCase);
            }
            static string StripComments(string s0)
            {
                if (string.IsNullOrEmpty(s0)) return string.Empty;
                var x = System.Text.RegularExpressions.Regex.Replace(s0, @"--.*?$", "", System.Text.RegularExpressions.RegexOptions.Multiline);
                x = System.Text.RegularExpressions.Regex.Replace(x, @"/\*[\s\S]*?\*/", "");
                return x;
            }

            // 对包含写入/DDL 的脚本：使用事务（同 SQL 查询分析一致）
            bool anyWrite = stmts.Any(x => !IsSelectLike(x));
            if (anyWrite && !expertMode)
            {
                SendMessageToWebView(new { action = "error", message = "脚本包含写入/DDL：请先开启【专家模式】", requestId, source });
                return;
            }

            // 风险二次校验（与单条一致）
            foreach (var one in stmts)
            {
                var s0 = StripComments(one).Trim().ToLowerInvariant();
                bool isDangerDdl = System.Text.RegularExpressions.Regex.IsMatch(s0, @"\b(drop|alter)\b");
                bool isUpdateNoWhere = System.Text.RegularExpressions.Regex.IsMatch(s0, @"\bupdate\b") && !System.Text.RegularExpressions.Regex.IsMatch(s0, @"\bwhere\b");
                bool isDeleteNoWhere = System.Text.RegularExpressions.Regex.IsMatch(s0, @"\bdelete\b") && !System.Text.RegularExpressions.Regex.IsMatch(s0, @"\bwhere\b");
                if (isDangerDdl && !allowDangerDdl)
                {
                    SendMessageToWebView(new { action = "error", message = "脚本含危险DDL（DROP/ALTER）被策略拦截：请勾选【允许危险DDL】", requestId, source });
                    return;
                }
                if ((isDangerDdl || isUpdateNoWhere || isDeleteNoWhere) && !confirmedDangerous)
                {
                    SendMessageToWebView(new { action = "error", message = "脚本含高风险语句未确认：请在弹窗确认后再次执行", requestId, source });
                    return;
                }
            }

            _sqlExecCts?.Cancel();
            _sqlExecCts?.Dispose();
            _sqlExecCts = new CancellationTokenSource();

            if (anyWrite && _sqliteManager.Connection != null)
            {
                _sqlLabTxn ??= _sqliteManager.Connection.BeginTransaction();
                SendMessageToWebView(new { action = "sqlLabTxnState", txnOpen = true });
            }

            Microsoft.Data.Sqlite.SqliteTransaction? txn = anyWrite ? _sqlLabTxn : null;
            QueryResult? lastQuery = null;
            int lastAffected = 0;
            bool hadError = false;
            string lastErrorMessage = "";
            for (int i = 0; i < stmts.Count; i++)
            {
                var one = stmts[i];
                var brief = one.Length > 120 ? one.Substring(0, 120) + "…" : one;
                SendMessageToWebView(new { action = "sqlScriptLog", requestId, source, index = i + 1, total = stmts.Count, message = brief });

                try
                {
                    var sw = Stopwatch.StartNew();
                    if (IsSelectLike(one))
                    {
                        // 脚本模式：默认只做结果预览（可配置 previewLimit），避免脚本里误跑超大 SELECT
                        string runSql = one;
                        try
                        {
                            if (previewLimit > 0)
                            {
                                var t0 = (one ?? string.Empty).TrimStart();
                                bool isPageable =
                                    t0.StartsWith("select", StringComparison.OrdinalIgnoreCase)
                                    || t0.StartsWith("with", StringComparison.OrdinalIgnoreCase)
                                    || t0.StartsWith("explain", StringComparison.OrdinalIgnoreCase);
                                bool hasLimit = one.IndexOf(" limit ", StringComparison.OrdinalIgnoreCase) >= 0;
                                if (isPageable && !hasLimit)
                                {
                                    var sqlNoSemi = one.Trim().TrimEnd(';');
                                    runSql = $"SELECT * FROM ({sqlNoSemi}) t LIMIT {previewLimit}";
                                }
                            }
                        }
                        catch { runSql = one; }

                        lastQuery = await _queryEngine.ExecuteQueryAsync(runSql, txn: txn, parameters: sqlParams, timeoutSeconds: timeoutSeconds, cancellationToken: _sqlExecCts.Token);
                        sw.Stop();
                        lastAffected = 0;
                        SendMessageToWebView(new
                        {
                            action = "sqlScriptStepResult",
                            requestId,
                            source,
                            index = i + 1,
                            total = stmts.Count,
                            kind = "query",
                            queryTime = Math.Round(sw.Elapsed.TotalSeconds, 2),
                            sql = one,
                            executedSql = lastQuery.Sql,
                            columns = lastQuery.Columns,
                            rows = lastQuery.Rows,
                            totalRows = lastQuery.TotalRows
                        });
                    }
                    else
                    {
                        lastQuery = null;
                        lastAffected = await _sqliteManager.ExecuteAsync(one, parameters: sqlParams, txn: txn, timeoutSeconds: timeoutSeconds, cancellationToken: _sqlExecCts.Token);
                        sw.Stop();
                        SendMessageToWebView(new
                        {
                            action = "sqlScriptStepResult",
                            requestId,
                            source,
                            index = i + 1,
                            total = stmts.Count,
                            kind = "exec",
                            queryTime = Math.Round(sw.Elapsed.TotalSeconds, 2),
                            sql = one,
                            affected = lastAffected
                        });
                    }
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception exStep)
                {
                    hadError = true;
                    lastErrorMessage = exStep.Message;
                    SendMessageToWebView(new
                    {
                        action = "sqlScriptStepError",
                        requestId,
                        source,
                        index = i + 1,
                        total = stmts.Count,
                        sql = one,
                        message = exStep.Message
                    });
                    if (!continueOnError) break;
                }
            }

            if (lastQuery != null)
            {
                SendMessageToWebView(new
                {
                    action = "queryComplete",
                    requestId,
                    source,
                    result = new
                    {
                        columns = lastQuery.Columns,
                        rows = lastQuery.Rows,
                        totalRows = lastQuery.TotalRows,
                        queryTime = lastQuery.QueryTime,
                        sql = sql,
                        executedSql = lastQuery.Sql,
                        txnOpen = _sqlLabTxn != null
                    }
                });
            }
            else
            {
                SendMessageToWebView(new
                {
                    action = "queryComplete",
                    requestId,
                    source,
                    result = new
                    {
                        columns = Array.Empty<string>(),
                        rows = Array.Empty<object>(),
                        totalRows = lastAffected,
                        queryTime = 0,
                        sql = sql,
                        executedSql = "",
                        txnOpen = _sqlLabTxn != null
                    }
                });
            }

            SendMessageToWebView(new
            {
                action = "sqlScriptDone",
                requestId,
                source,
                ok = !hadError,
                message = hadError ? ("脚本执行结束（部分失败）：" + lastErrorMessage) : "脚本执行完成"
            });
        }
        catch (OperationCanceledException)
        {
            SendMessageToWebView(new { action = "sqlCancelled", message = "已取消执行", requestId, source });
        }
        catch (Exception ex)
        {
            WriteErrorLog("执行SQL脚本失败", ex);
            SendMessageToWebView(new { action = "error", message = $"脚本执行失败: {ex.Message}", hasErrorLog = true, requestId, source });
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
            string outputOption = (data.TryGetProperty("outputOption", out var oo) ? oo.GetString() : null) ?? "files";
            string fileNamePrefix = (data.TryGetProperty("fileNamePrefix", out var fnp) ? fnp.GetString() : null) ?? string.Empty;
            bool overwriteExisting = data.TryGetProperty("overwriteExisting", out var ow) && ow.ValueKind == JsonValueKind.True;
            bool includeNullGroup = !(data.TryGetProperty("includeNullGroup", out var ing) && ing.ValueKind == JsonValueKind.False);
            string csvDelimiter = (data.TryGetProperty("csvDelimiter", out var cd) ? cd.GetString() : null) ?? ",";
            csvDelimiter = string.IsNullOrEmpty(csvDelimiter) ? "," : csvDelimiter[..1];
            string sql = (data.TryGetProperty("sql", out var sq) ? sq.GetString() : null) ?? string.Empty;

            if (_splitEngine == null)
            {
                SendMessageToWebView(new { action = "error", message = "分拆引擎未初始化" });
                return;
            }
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }

            SplitResult result;
            string tempTable = "";
            try
            {
                string actualSource = sourceTable;
                if (!string.IsNullOrWhiteSpace(sql))
                {
                    // SQL 结果集分拆：先落到临时表，再复用 SplitEngine
                    tempTable = "TempSplit_" + Guid.NewGuid().ToString("N");
                    _sqliteManager.Execute($"DROP TABLE IF EXISTS [{tempTable}];");
                    var sqlNoSemi = sql.Trim().TrimEnd(';');
                    _sqliteManager.Execute($"CREATE TABLE [{tempTable}] AS {sqlNoSemi};");
                    actualSource = tempTable;
                }

                if (splitType == "byField")
                {
                    result = _splitEngine.SplitByField(
                        actualSource,
                        splitField,
                        outputDirectory,
                        outputOption: outputOption,
                        fileNamePrefix: string.IsNullOrWhiteSpace(fileNamePrefix) ? null : fileNamePrefix,
                        includeNullGroup: includeNullGroup,
                        overwriteExisting: overwriteExisting,
                        csvDelimiter: csvDelimiter,
                        progress: (pct, stage) => SendMessageToWebView(new { action = "splitProgressUpdate", target = "stspt", percent = pct, stage = stage, running = true }));
                }
                else
                {
                    int rowsPerFile = data.GetProperty("rowsPerFile").GetInt32();
                    result = _splitEngine.SplitByRowCount(
                        actualSource,
                        rowsPerFile,
                        outputDirectory,
                        outputOption: outputOption,
                        fileNamePrefix: string.IsNullOrWhiteSpace(fileNamePrefix) ? null : fileNamePrefix,
                        overwriteExisting: overwriteExisting,
                        csvDelimiter: csvDelimiter,
                        progress: (pct, stage) => SendMessageToWebView(new { action = "splitProgressUpdate", target = "stspt", percent = pct, stage = stage, running = true }));
                }
            }
            finally
            {
                try { if (!string.IsNullOrWhiteSpace(tempTable)) _sqliteManager.Execute($"DROP TABLE IF EXISTS [{tempTable}];"); } catch { }
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
            string outputOption = (data.TryGetProperty("outputOption", out var oo) ? oo.GetString() : null) ?? "files";
            string fileNamePrefix = (data.TryGetProperty("fileNamePrefix", out var fnp) ? fnp.GetString() : null) ?? string.Empty;
            bool overwriteExisting = data.TryGetProperty("overwriteExisting", out var ow) && ow.ValueKind == JsonValueKind.True;
            bool includeNullGroup = !(data.TryGetProperty("includeNullGroup", out var ing) && ing.ValueKind == JsonValueKind.False);
            string csvDelimiter = (data.TryGetProperty("csvDelimiter", out var cd) ? cd.GetString() : null) ?? ",";
            csvDelimiter = string.IsNullOrEmpty(csvDelimiter) ? "," : csvDelimiter[..1];

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
                result = _splitEngine.SplitByField(
                    tempTable,
                    splitField,
                    outputDirectory,
                    outputOption: outputOption,
                    fileNamePrefix: string.IsNullOrWhiteSpace(fileNamePrefix) ? null : fileNamePrefix,
                    includeNullGroup: includeNullGroup,
                    overwriteExisting: overwriteExisting,
                    csvDelimiter: csvDelimiter,
                    progress: (pct, stage) => SendMessageToWebView(new { action = "splitProgressUpdate", target = "mtspt", percent = pct, stage = stage, running = true }));
            }
            else
            {
                int rowsPerFile = data.GetProperty("rowsPerFile").GetInt32();
                result = _splitEngine.SplitByRowCount(
                    tempTable,
                    rowsPerFile,
                    outputDirectory,
                    outputOption: outputOption,
                    fileNamePrefix: string.IsNullOrWhiteSpace(fileNamePrefix) ? null : fileNamePrefix,
                    overwriteExisting: overwriteExisting,
                    csvDelimiter: csvDelimiter,
                    progress: (pct, stage) => SendMessageToWebView(new { action = "splitProgressUpdate", target = "mtspt", percent = pct, stage = stage, running = true }));
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

    private void GetSplitPreview(System.Text.Json.JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }

            var target = (data.TryGetProperty("target", out var tg) ? tg.GetString() : null) ?? "stspt";
            var splitType = (data.TryGetProperty("splitType", out var st) ? st.GetString() : null) ?? "byField";
            var splitField = (data.TryGetProperty("splitField", out var sf) ? sf.GetString() : null) ?? "";
            var sourceTable = (data.TryGetProperty("sourceTable", out var tb) ? tb.GetString() : null) ?? "";
            var sql = (data.TryGetProperty("sql", out var sq) ? sq.GetString() : null) ?? "";
            var rowsPerFile = (data.TryGetProperty("rowsPerFile", out var rpf) && rpf.TryGetInt32(out var n)) ? n : 50000;
            if (rowsPerFile <= 0) rowsPerFile = 50000;

            if (string.IsNullOrWhiteSpace(splitField))
            {
                SendMessageToWebView(new { action = "splitPreviewLoaded", target = target, stats = new { totalRows = 0, distinctCount = 0, nullRows = 0, topValues = Array.Empty<object>() } });
                return;
            }

            string fromSql;
            if (!string.IsNullOrWhiteSpace(sql))
            {
                var sqlNoSemi = sql.Trim().TrimEnd(';');
                fromSql = $"({sqlNoSemi}) t";
            }
            else
            {
                if (string.IsNullOrWhiteSpace(sourceTable))
                    sourceTable = MainTableNameOrDefault();
                fromSql = $"{SqliteManager.QuoteIdent(sourceTable)}";
            }

            var col = SqliteManager.QuoteIdent(splitField);

            long ScalarLong(string sqlText)
            {
                var rows = _sqliteManager.Query(sqlText);
                if (rows.Count == 0) return 0;
                if (rows[0].TryGetValue("c", out var v) && v != null && v != DBNull.Value) return Convert.ToInt64(v);
                // fallback：取第一列
                var first = rows[0].Values.FirstOrDefault();
                return first == null || first == DBNull.Value ? 0 : Convert.ToInt64(first);
            }

            // total rows
            long totalRows = ScalarLong($"SELECT COUNT(1) AS c FROM {fromSql};");

            if (string.Equals(splitType, "byRowCount", StringComparison.OrdinalIgnoreCase))
            {
                var est = (int)Math.Ceiling(totalRows / (double)rowsPerFile);
                SendMessageToWebView(new
                {
                    action = "splitPreviewLoaded",
                    target = target,
                    stats = new
                    {
                        totalRows = totalRows,
                        distinctCount = 0,
                        nullRows = 0,
                        estimatedFiles = est,
                        topValues = Array.Empty<object>()
                    }
                });
                return;
            }

            long distinctCount = ScalarLong($"SELECT COUNT(DISTINCT {col}) AS c FROM {fromSql};");
            long nullRows = ScalarLong($"SELECT COUNT(1) AS c FROM {fromSql} WHERE {col} IS NULL;");

            var topSql = $"SELECT {col} AS v, COUNT(1) AS c FROM {fromSql} WHERE {col} IS NOT NULL GROUP BY {col} ORDER BY c DESC LIMIT 20;";
            var topRows = _sqliteManager.Query(topSql);
            var topValues = topRows.Select(r => new
            {
                value = Convert.ToString(r["v"]) ?? "",
                count = Convert.ToInt64(r["c"] ?? 0)
            }).ToList();

            SendMessageToWebView(new
            {
                action = "splitPreviewLoaded",
                target = target,
                stats = new
                {
                    totalRows = totalRows,
                    distinctCount = distinctCount,
                    nullRows = nullRows,
                    topValues = topValues
                }
            });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"分拆预览失败: {ex.Message}" });
        }
    }

    private void GetSqlResultSchema(System.Text.Json.JsonElement data)
    {
        string tempTable = "TempSchema_" + Guid.NewGuid().ToString("N");
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }
            var target = (data.TryGetProperty("target", out var tg) ? tg.GetString() : null) ?? "unknown";
            var sql = (data.TryGetProperty("sql", out var sq) ? sq.GetString() : null) ?? "";
            sql = sql.Trim();
            if (string.IsNullOrWhiteSpace(sql))
            {
                SendMessageToWebView(new { action = "sqlResultSchemaLoaded", target = target, fields = Array.Empty<string>() });
                return;
            }

            var sqlNoSemi = sql.Trim().TrimEnd(';');
            // 通过“落临时表 + PRAGMA table_info”获取列名，兼容复杂 join/select 表达式
            _sqliteManager.Execute($"DROP TABLE IF EXISTS [{tempTable}];");
            _sqliteManager.Execute($"CREATE TABLE [{tempTable}] AS SELECT * FROM ({sqlNoSemi}) t LIMIT 0;");
            var schema = _sqliteManager.GetTableSchema(tempTable);
            var fields = schema.Select(s => s.ColumnName).Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
            SendMessageToWebView(new { action = "sqlResultSchemaLoaded", target = target, fields = fields });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"获取SQL结果字段失败: {ex.Message}" });
        }
        finally
        {
            try { _sqliteManager?.Execute($"DROP TABLE IF EXISTS [{tempTable}];"); } catch { }
        }
    }

    private void OpenPath(System.Text.Json.JsonElement data)
    {
        try
        {
            var path = (data.TryGetProperty("path", out var p) ? p.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(path))
            {
                SendMessageToWebView(new { action = "error", message = "打开失败：路径为空" });
                return;
            }
            // 允许目录或文件
            if (!Directory.Exists(path) && !File.Exists(path))
            {
                SendMessageToWebView(new { action = "error", message = $"打开失败：路径不存在 {path}" });
                return;
            }

            try
            {
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = path,
                    UseShellExecute = true
                };
                System.Diagnostics.Process.Start(psi);
                SendMessageToWebView(new { action = "status", message = "已打开： " + path });
            }
            catch (Exception ex)
            {
                SendMessageToWebView(new { action = "error", message = $"打开失败: {ex.Message}" });
            }
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"打开失败: {ex.Message}" });
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

    private void GetDbObjects()
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }

            // main + attached DB 统一返回（Name 统一为：mainName / alias.Name）
            var objects = new List<object>();
            var dbs = _sqliteManager.Query("PRAGMA database_list;")
                .Select(r => Convert.ToString(r["name"]) ?? "")
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var db in dbs)
            {
                if (string.Equals(db, "temp", StringComparison.OrdinalIgnoreCase)) continue;
                var sql = $"SELECT name, type FROM {SqliteManager.QuoteIdent(db)}.sqlite_master WHERE (type='table' OR type='view') AND name NOT LIKE 'sqlite_%' ORDER BY type, name;";
                var rows = _sqliteManager.Query(sql);
                foreach (var r in rows)
                {
                    var name = Convert.ToString(r["name"]) ?? "";
                    var type = Convert.ToString(r["type"]) ?? "";
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    var fullName = string.Equals(db, "main", StringComparison.OrdinalIgnoreCase) ? name : $"{db}.{name}";
                    objects.Add(new { name = fullName, type = type });
                }
            }

            SendMessageToWebView(new { action = "dbObjectsLoaded", objects });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"获取数据库对象失败: {ex.Message}" });
        }
    }

    private void DropDbObjects(JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }

            if (!data.TryGetProperty("objects", out var arr) || arr.ValueKind != JsonValueKind.Array)
            {
                SendMessageToWebView(new { action = "error", message = "删除失败：objects 参数缺失" });
                return;
            }

            // 自动修订点：删除前备份（便于回滚）
            string? rpId = null;
            try
            {
                var names = arr.EnumerateArray()
                    .Select(x => (x.TryGetProperty("name", out var n) ? n.GetString() : null) ?? "")
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Take(10)
                    .ToList();
                var note = $"dropDbObjects: {string.Join(", ", names)}";
                var r = CreateRevisionPointInternal("auto", note, enableBackup: true, sendMessage: false);
                rpId = r.Rp?.Id;
            }
            catch { }

            int ok = 0, fail = 0;
            foreach (var it in arr.EnumerateArray())
            {
                var name = (it.TryGetProperty("name", out var n) ? n.GetString() : null) ?? "";
                var type = (it.TryGetProperty("type", out var t) ? t.GetString() : null) ?? "";
                if (string.IsNullOrWhiteSpace(name)) continue;
                var isView = string.Equals(type, "view", StringComparison.OrdinalIgnoreCase);
                try
                {
                    var drop = isView ? $"DROP VIEW IF EXISTS {SqliteManager.QuoteIdent(name)};" : $"DROP TABLE IF EXISTS {SqliteManager.QuoteIdent(name)};";
                    _sqliteManager.Execute(drop);
                    ok++;
                    // 若删除的是当前主表：清空（避免后续查询误指向不存在表）
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(_currentMainTableName) && string.Equals(_currentMainTableName, name, StringComparison.OrdinalIgnoreCase))
                            _currentMainTableName = "";
                    }
                    catch { }
                }
                catch
                {
                    fail++;
                }
            }

            // 同步派生视图元数据（vw_）
            try { if (!string.IsNullOrWhiteSpace(_activeSchemeId)) SyncDerivedViewsMetaFromDb(_activeSchemeId!); } catch { }

            // 删除后：刷新表列表与对象列表（前端也会自行刷新，但这里更稳）
            try { GetTableList(); } catch { }
            try { GetDbObjects(); } catch { }

            SendMessageToWebView(new { action = "dbObjectsDropped", message = $"删除完成：成功 {ok}，失败 {fail}", revisionPointId = rpId });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"删除失败: {ex.Message}" });
        }
    }

    private void GetDbObjectDependencies(JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }

            var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
            var target = (data.TryGetProperty("name", out var n) ? n.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(target))
            {
                SendMessageToWebView(new { action = "error", message = "依赖分析失败：name 为空" });
                return;
            }

            // 暂不做跨 DB 的复杂解析：用“文本包含”做依赖探测（对 VIEW 足够实用）
            bool ContainsRef(string? sql, string t)
            {
                if (string.IsNullOrWhiteSpace(sql) || string.IsNullOrWhiteSpace(t)) return false;
                var s = sql!;
                var x = t.Trim();
                var bare = x.Contains('.') ? x.Split('.', 2)[1] : x;
                // 使用标识符边界匹配，避免误报（例如 oldName 是 another_oldName 的子串）
                if (ContainsSqlIdentifier(s, x)) return true;
                if (!string.Equals(bare, x, StringComparison.OrdinalIgnoreCase) && ContainsSqlIdentifier(s, bare)) return true;
                return false;
            }

            var dependents = new List<object>();
            var dbs = _sqliteManager.Query("PRAGMA database_list;")
                .Select(r => Convert.ToString(r["name"]) ?? "")
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var db in dbs)
            {
                if (string.Equals(db, "temp", StringComparison.OrdinalIgnoreCase)) continue;
                var sql = $"SELECT name, type, sql FROM {SqliteManager.QuoteIdent(db)}.sqlite_master WHERE type='view' AND sql IS NOT NULL AND name NOT LIKE 'sqlite_%' ORDER BY name;";
                var rows = _sqliteManager.Query(sql);
                foreach (var r in rows)
                {
                    var name = Convert.ToString(r["name"]) ?? "";
                    var type = Convert.ToString(r["type"]) ?? "";
                    var viewSql = Convert.ToString(r["sql"]) ?? "";
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    var fullName = string.Equals(db, "main", StringComparison.OrdinalIgnoreCase) ? name : $"{db}.{name}";
                    if (string.Equals(fullName, target, StringComparison.OrdinalIgnoreCase)) continue;
                    if (ContainsRef(viewSql, target))
                    {
                        dependents.Add(new { name = fullName, type = type });
                    }
                }
            }

            // 项目配置依赖（精确：解析当前项目 settingsJson，返回引用路径）
            var configRefs = new List<object>();
            try
            {
                if (!string.IsNullOrWhiteSpace(_activeSchemeId))
                {
                    var iniPath = GetSchemeIniPath(_activeSchemeId);
                    if (File.Exists(iniPath))
                    {
                        var ini = ReadIni(iniPath);
                        if (ini.TryGetValue("settings", out var st) && st.TryGetValue("settingsJsonB64", out var sj) && !string.IsNullOrWhiteSpace(sj))
                        {
                            var settingsJson = FromB64(sj);
                            if (!string.IsNullOrWhiteSpace(settingsJson))
                            {
                                var bare = target.Contains('.') ? target.Split('.', 2)[1] : target;
                                using var doc = System.Text.Json.JsonDocument.Parse(settingsJson);
                                void Walk(System.Text.Json.JsonElement el, string path, string? propName)
                                {
                                    switch (el.ValueKind)
                                    {
                                        case System.Text.Json.JsonValueKind.Object:
                                            foreach (var p in el.EnumerateObject())
                                                Walk(p.Value, $"{path}.{p.Name}", p.Name);
                                            break;
                                        case System.Text.Json.JsonValueKind.Array:
                                            int i = 0;
                                            foreach (var it in el.EnumerateArray())
                                                Walk(it, $"{path}[{i++}]", propName);
                                            break;
                                        case System.Text.Json.JsonValueKind.String:
                                            var s = el.GetString() ?? "";
                                            if (string.IsNullOrWhiteSpace(s)) break;
                                            if (s.IndexOf(target, StringComparison.OrdinalIgnoreCase) >= 0 || s.IndexOf(bare, StringComparison.OrdinalIgnoreCase) >= 0)
                                            {
                                                var snippet = s.Length > 140 ? s.Substring(0, 140) + "..." : s;
                                                var hint =
                                                    (propName ?? "").IndexOf("sql", StringComparison.OrdinalIgnoreCase) >= 0 ? "SQL片段" :
                                                    (propName ?? "").IndexOf("table", StringComparison.OrdinalIgnoreCase) >= 0 ? "表名字段" :
                                                    "字符串";
                                                configRefs.Add(new { scope = "当前项目配置", path, hint, snippet });
                                            }
                                            break;
                                        default:
                                            break;
                                    }
                                }
                                Walk(doc.RootElement, "$", null);
                            }
                        }
                    }
                }
            }
            catch { }

            SendMessageToWebView(new { action = "dbObjectDependenciesLoaded", requestId, target, dependents, configRefs });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "error", message = $"依赖分析失败: {ex.Message}" });
        }
    }

    private void RenameDbObject(JsonElement data)
    {
        try
        {
            if (_sqliteManager == null)
            {
                SendMessageToWebView(new { action = "error", message = "SQLite管理器未初始化" });
                return;
            }

            var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
            var oldName = (data.TryGetProperty("oldName", out var o) ? o.GetString() : null) ?? "";
            var newName = (data.TryGetProperty("newName", out var nn) ? nn.GetString() : null) ?? "";
            var objType = (data.TryGetProperty("type", out var t) ? t.GetString() : null) ?? "table"; // table|view
            var cascade = data.TryGetProperty("cascadeDependents", out var c) && c.ValueKind == JsonValueKind.True;

            oldName = oldName.Trim();
            newName = newName.Trim();
            objType = objType.Trim().ToLowerInvariant();

            void Fail(string msg)
            {
                try
                {
                    SendMessageToWebView(new
                    {
                        action = "dbObjectRenameFailed",
                        requestId,
                        oldName,
                        newName,
                        type = objType,
                        message = msg
                    });
                }
                catch { }
                SendMessageToWebView(new { action = "error", message = msg });
            }

            if (string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName))
            {
                Fail("重命名失败：oldName/newName 为空");
                return;
            }

            // 自动修订点：重命名前备份（便于回滚）
            string? rpId = null;
            try
            {
                var note = $"renameDbObject: {oldName} -> {newName}";
                var r = CreateRevisionPointInternal("auto", note, enableBackup: true, sendMessage: false);
                rpId = r.Rp?.Id;
            }
            catch { }

            // 先保守：不支持对附加库对象（含 alias. 前缀）做 rename
            if (oldName.Contains('.') || newName.Contains('.'))
            {
                Fail("暂不支持对“附加库对象（含 db. 前缀）”重命名，请先在对应库内处理。");
                return;
            }

            // 检查是否已存在
            if (_sqliteManager.TableExists(newName))
            {
                Fail($"重命名失败：目标名称已存在：{newName}");
                return;
            }

            // 依赖视图：由前端先做依赖分析后传入（可选；仅支持 main 里的视图名）
            var dependents = new List<string>();
            if (data.TryGetProperty("dependents", out var depArr) && depArr.ValueKind == JsonValueKind.Array)
            {
                foreach (var it in depArr.EnumerateArray())
                {
                    var dn = (it.GetString() ?? "").Trim();
                    if (!string.IsNullOrWhiteSpace(dn) && !dn.Contains('.')) dependents.Add(dn);
                }
            }

            if (objType == "table")
            {
                if (!cascade && dependents.Count > 0)
                {
                    Fail($"重命名失败：存在依赖视图（{dependents.Count}个），请先处理依赖或选择级联更新。");
                    return;
                }

                // 级联：先备份依赖视图 SQL，删掉它们，重命名表后再重建
                var depSql = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (cascade && dependents.Count > 0)
                {
                    foreach (var dn in dependents)
                    {
                        var r2 = _sqliteManager.Query("SELECT sql FROM sqlite_master WHERE type='view' AND name=@name LIMIT 1;", new { name = dn });
                        var s2 = r2.FirstOrDefault()?["sql"]?.ToString() ?? "";
                        if (!string.IsNullOrWhiteSpace(s2)) depSql[dn] = s2;
                    }
                    foreach (var dn in dependents)
                    {
                        _sqliteManager.Execute($"DROP VIEW IF EXISTS {SqliteManager.QuoteIdent(dn)};");
                    }
                }

                _sqliteManager.Execute($"ALTER TABLE {SqliteManager.QuoteIdent(oldName)} RENAME TO {SqliteManager.QuoteIdent(newName)};");

                try
                {
                    if (!string.IsNullOrWhiteSpace(_currentMainTableName) && string.Equals(_currentMainTableName, oldName, StringComparison.OrdinalIgnoreCase))
                        _currentMainTableName = newName;
                }
                catch { }

                if (cascade && depSql.Count > 0)
                {
                    // 更新依赖视图 SQL：把 oldName 的引用替换为 newName（含 [old]）
                    string ReplaceRef(string s)
                    {
                        if (string.IsNullOrWhiteSpace(s)) return s;
                        return ReplaceSqlIdentifier(s, oldName, newName);
                    }

                    var pending = depSql.Select(kv => new KeyValuePair<string, string>(kv.Key, ReplaceRef(kv.Value))).ToList();
                    for (int pass = 0; pass < 10 && pending.Count > 0; pass++)
                    {
                        int okCount = 0;
                        for (int i = pending.Count - 1; i >= 0; i--)
                        {
                            try
                            {
                                _sqliteManager.Execute(pending[i].Value);
                                pending.RemoveAt(i);
                                okCount++;
                            }
                            catch
                            {
                                // 等下一轮（依赖未满足）
                            }
                        }
                        if (okCount == 0) break;
                    }
                    if (pending.Count > 0)
                        throw new Exception($"级联重建依赖视图失败：仍有 {pending.Count} 个视图无法创建。");
                }
            }
            else
            {
                // view：获取 SQL 定义
                var rows = _sqliteManager.Query("SELECT sql FROM sqlite_master WHERE type='view' AND name=@name LIMIT 1;", new { name = oldName });
                var sql = rows.FirstOrDefault()?["sql"]?.ToString() ?? "";
                if (string.IsNullOrWhiteSpace(sql))
                {
                    Fail($"重命名失败：未找到视图定义：{oldName}");
                    return;
                }

                // 生成新视图 SQL（替换 CREATE VIEW 头部）
                var re = new System.Text.RegularExpressions.Regex(@"(?is)^\\s*CREATE\\s+VIEW\\s+(?:IF\\s+NOT\\s+EXISTS\\s+)?(.+?)\\s+AS\\s+", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                var newSql = re.Replace(sql, $"CREATE VIEW {SqliteManager.QuoteIdent(newName)} AS ", 1);

                if (!cascade && dependents.Count > 0)
                {
                    Fail($"重命名失败：存在依赖视图（{dependents.Count}个），请先处理依赖或选择级联更新。");
                    return;
                }

                // 级联：先备份依赖视图 SQL，删掉它们，重建
                var depSql = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (cascade && dependents.Count > 0)
                {
                    foreach (var dn in dependents)
                    {
                        var r2 = _sqliteManager.Query("SELECT sql FROM sqlite_master WHERE type='view' AND name=@name LIMIT 1;", new { name = dn });
                        var s2 = r2.FirstOrDefault()?["sql"]?.ToString() ?? "";
                        if (!string.IsNullOrWhiteSpace(s2)) depSql[dn] = s2;
                    }
                    // drop dependents（链式依赖：全删最稳）
                    foreach (var dn in dependents)
                    {
                        _sqliteManager.Execute($"DROP VIEW IF EXISTS {SqliteManager.QuoteIdent(dn)};");
                    }
                }

                _sqliteManager.Execute($"DROP VIEW IF EXISTS {SqliteManager.QuoteIdent(oldName)};");
                _sqliteManager.Execute(newSql);

                if (cascade && depSql.Count > 0)
                {
                    // 更新依赖 SQL：把 oldName 的引用替换为 newName（含 [old]）
                    string ReplaceRef(string s)
                    {
                        if (string.IsNullOrWhiteSpace(s)) return s;
                        return ReplaceSqlIdentifier(s, oldName, newName);
                    }

                    var pending = depSql.Select(kv => new KeyValuePair<string, string>(kv.Key, ReplaceRef(kv.Value))).ToList();
                    for (int pass = 0; pass < 10 && pending.Count > 0; pass++)
                    {
                        int okCount = 0;
                        for (int i = pending.Count - 1; i >= 0; i--)
                        {
                            try
                            {
                                _sqliteManager.Execute(pending[i].Value);
                                pending.RemoveAt(i);
                                okCount++;
                            }
                            catch
                            {
                                // 等下一轮（依赖未满足）
                            }
                        }
                        if (okCount == 0) break;
                    }
                    if (pending.Count > 0)
                    {
                        throw new Exception($"级联重建依赖视图失败：仍有 {pending.Count} 个视图无法创建。");
                    }
                }
            }

            // 同步当前项目 settingsJson（若存在）：用“文本替换/结构化”更新主表/关联/模板等引用
            var settingsUpdateReport = new List<JsonUpdateItem>();
            var excelSettingsUpdateReport = new List<IniUpdateItem>();
            try
            {
                if (!string.IsNullOrWhiteSpace(_activeSchemeId))
                {
                    var iniPath = GetSchemeIniPath(_activeSchemeId);
                    if (File.Exists(iniPath))
                    {
                        var ini = ReadIni(iniPath);
                        if (ini.TryGetValue("settings", out var st) && st.TryGetValue("settingsJsonB64", out var sj) && !string.IsNullOrWhiteSpace(sj))
                        {
                            var settingsJson = FromB64(sj);
                            if (!string.IsNullOrWhiteSpace(settingsJson))
                            {
                                var r = UpdateSchemeSettingsJsonReferences(settingsJson, oldName, newName);
                                var updated = r.UpdatedJson;
                                settingsUpdateReport = r.Report ?? new List<JsonUpdateItem>();
                                if (!string.Equals(updated, settingsJson, StringComparison.Ordinal))
                                {
                                    st["settingsJsonB64"] = ToB64(updated);
                                    WriteIni(iniPath, ini);
                                    // 更新 Project Meta 的主表名（避免 project hub 展示旧值）
                                    try
                                    {
                                        var pm = LoadOrCreateProjectMeta(_activeSchemeId, _activeSchemeId, updated);
                                        pm.MainTableName = ExtractMainTableNameFromSettings(updated) ?? pm.MainTableName;
                                        SaveProjectMeta(pm);
                                    }
                                    catch { }
                                }
                            }
                        }
                    }
                }
            }
            catch { }

            // 同步各源文件的 excelsqlite.ini（模块设置），避免模块仍引用旧表名
            try { excelSettingsUpdateReport = UpdateExcelSettingsIniForActiveScheme(oldName, newName); } catch { }

            // 同步派生视图元数据（vw_）：重命名表/视图会影响实际 DB 中视图
            try { if (!string.IsNullOrWhiteSpace(_activeSchemeId)) SyncDerivedViewsMetaFromDb(_activeSchemeId!); } catch { }

            try { GetTableList(); } catch { }
            try { GetDbObjects(); } catch { }
            SendMessageToWebView(new
            {
                action = "dbObjectRenamed",
                requestId,
                oldName,
                newName,
                type = objType,
                revisionPointId = rpId,
                settingsUpdateReport = settingsUpdateReport,
                excelSettingsUpdateReport = excelSettingsUpdateReport
            });
        }
        catch (Exception ex)
        {
            try
            {
                var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
                var oldName = (data.TryGetProperty("oldName", out var o) ? o.GetString() : null) ?? "";
                var newName = (data.TryGetProperty("newName", out var nn) ? nn.GetString() : null) ?? "";
                var objType = (data.TryGetProperty("type", out var t) ? t.GetString() : null) ?? "table";
                SendMessageToWebView(new { action = "dbObjectRenameFailed", requestId, oldName, newName, type = objType, message = $"重命名失败: {ex.Message}" });
            }
            catch { }
            SendMessageToWebView(new { action = "error", message = $"重命名失败: {ex.Message}" });
        }
    }

    private void RestoreRevisionPoint(JsonElement data)
    {
        try
        {
            var schemeId = (data.TryGetProperty("schemeId", out var sid) ? sid.GetString() : null) ?? (_activeSchemeId ?? "");
            var revisionId = (data.TryGetProperty("revisionId", out var rid) ? rid.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(schemeId) || string.IsNullOrWhiteSpace(revisionId))
            {
                SendMessageToWebView(new { action = "revisionPointRestored", ok = false, message = "参数缺失：schemeId/revisionId" });
                return;
            }

            // 仅允许恢复当前激活项目（避免误操作其它项目）
            if (!string.IsNullOrWhiteSpace(_activeSchemeId) && !string.Equals(_activeSchemeId, schemeId, StringComparison.OrdinalIgnoreCase))
            {
                SendMessageToWebView(new { action = "revisionPointRestored", ok = false, message = "请先在项目中心切换到对应项目后再恢复修订点" });
                return;
            }

            var meta = LoadOrCreateProjectMeta(schemeId, displayName: null, settingsJson: null);
            var rp = (meta.RevisionPoints ?? new List<RevisionPointV1>())
                .FirstOrDefault(x => string.Equals(x.Id, revisionId, StringComparison.OrdinalIgnoreCase));
            if (rp == null || string.IsNullOrWhiteSpace(rp.DbBackupPath) || !File.Exists(rp.DbBackupPath))
            {
                SendMessageToWebView(new { action = "revisionPointRestored", ok = false, message = "修订点备份文件不存在或已被清理" });
                return;
            }

            var targetDb = GetSchemeDbPath(schemeId);
            Directory.CreateDirectory(Path.GetDirectoryName(targetDb) ?? GetDbDir());

            // 关闭现有连接，避免文件被占用
            try { _sqliteManager?.Dispose(); } catch { }

            // 覆盖恢复
            File.Copy(rp.DbBackupPath, targetDb, overwrite: true);

            _activeSchemeId = schemeId;
            _activeSchemeDbPath = targetDb;
            SwitchSqliteToFileDb(targetDb, backupFromExisting: false);

            // 更新项目元数据
            try
            {
                meta.DbPath = targetDb;
                meta.DbTables = GetDbTablesForScheme(schemeId);
                meta.DbSizeBytes = SafeFileSize(targetDb);
                meta.UpdatedAt = DateTime.Now;
                SaveProjectMeta(meta);
            }
            catch { }

            // 同步派生视图元数据（vw_）
            try { SyncDerivedViewsMetaFromDb(schemeId); } catch { }

            // 通知前端刷新
            try { GetTableList(); } catch { }
            try
            {
                SendMessageToWebView(new
                {
                    action = "projectDetail",
                    schemeId = schemeId,
                    meta = meta
                });
            }
            catch { }

            SendMessageToWebView(new { action = "revisionPointRestored", ok = true, message = "修订点已恢复", schemeId = schemeId, revisionId = revisionId });
        }
        catch (Exception ex)
        {
            WriteErrorLog("恢复修订点失败", ex);
            SendMessageToWebView(new { action = "revisionPointRestored", ok = false, message = $"恢复修订点失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void AppendAuditLog(JsonElement data)
    {
        try
        {
            string schemeId = (data.TryGetProperty("schemeId", out var sid) ? sid.GetString() : null) ?? (_activeSchemeId ?? "");
            if (string.IsNullOrWhiteSpace(schemeId))
            {
                SendMessageToWebView(new { action = "error", message = "审计记录失败：schemeId 为空" });
                return;
            }

            var meta = LoadOrCreateProjectMeta(schemeId, displayName: null, settingsJson: null);
            meta.AuditLogs = meta.AuditLogs ?? new List<AuditLogV1>();

            string type = (data.TryGetProperty("type", out var tp) ? tp.GetString() : null) ?? "audit";
            string note = (data.TryGetProperty("note", out var nt) ? nt.GetString() : null) ?? "";
            string? revId = (data.TryGetProperty("revisionPointId", out var rp) ? rp.GetString() : null);

            var revIds = new List<string>();
            if (data.TryGetProperty("revisionPointIds", out var rps) && rps.ValueKind == JsonValueKind.Array)
                revIds = rps.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).Take(50).ToList();

            var reportPaths = new List<string>();
            if (data.TryGetProperty("reportPaths", out var arr) && arr.ValueKind == JsonValueKind.Array)
                reportPaths = arr.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).Take(20).ToList();
            else if (data.TryGetProperty("reportPath", out var one) && one.ValueKind == JsonValueKind.String)
            {
                var p = one.GetString() ?? "";
                if (!string.IsNullOrWhiteSpace(p)) reportPaths.Add(p);
            }

            string payloadJson = "";
            try
            {
                if (data.TryGetProperty("payload", out var pl))
                    payloadJson = pl.GetRawText();
            }
            catch { }

            var log = new AuditLogV1
            {
                Id = Guid.NewGuid().ToString("N"),
                Time = DateTime.Now,
                Type = type ?? "audit",
                Note = note ?? "",
                RevisionPointId = string.IsNullOrWhiteSpace(revId) ? null : revId,
                RevisionPointIds = revIds,
                ReportPaths = reportPaths,
                PayloadJson = payloadJson ?? ""
            };

            meta.AuditLogs.Insert(0, log);
            // 控制大小：最多 500 条
            if (meta.AuditLogs.Count > 500) meta.AuditLogs = meta.AuditLogs.Take(500).ToList();
            SaveProjectMeta(meta);

            SendMessageToWebView(new { action = "auditLogAppended", ok = true, schemeId = schemeId, log = log, meta = meta });
        }
        catch (Exception ex)
        {
            WriteErrorLog("写入审计日志失败", ex);
            SendMessageToWebView(new { action = "error", message = $"写入审计日志失败: {ex.Message}", hasErrorLog = true });
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
            bool skipCount = (data.TryGetProperty("skipCount", out var sc) && sc.ValueKind == System.Text.Json.JsonValueKind.True)
                || (data.TryGetProperty("skipCount", out sc) && sc.ValueKind == System.Text.Json.JsonValueKind.False ? false : false);
            IDictionary<string, object?>? sqlParams = null;
            try { sqlParams = ReadSqlParams(data); } catch { sqlParams = null; }

            if (string.IsNullOrWhiteSpace(sql))
            {
                SendMessageToWebView(new { action = "error", message = "导出失败：SQL为空" });
                return;
            }

            // 大数量提示：尽量计算行数（对大表 count(*) 也可能耗时，但比导出更快）
            long cnt = 0;
            bool forceAllText = string.Equals(format, "xlsx_text", StringComparison.OrdinalIgnoreCase);

            // 很多复杂 SQL（JOIN/UNION/子查询）在 COUNT(*) 预统计阶段会非常慢，用户体验像“点击导出没反应”。
            // 规则：显式 skipCount=true 或 SQL 已带 LIMIT 时，跳过预统计，直接弹保存对话框并导出。
            bool hasLimit = sql.IndexOf(" limit ", StringComparison.OrdinalIgnoreCase) >= 0;
            bool isXlsx = string.Equals(format, "xlsx", StringComparison.OrdinalIgnoreCase) || forceAllText;
            if (!skipCount && !hasLimit)
            {
                try
                {
                    var sqlNoSemi = sql.Trim().TrimEnd(';');
                    var countSql = $"SELECT COUNT(*) AS cnt FROM ({sqlNoSemi}) t";
                    var cntRows = _sqliteManager.Query(countSql, sqlParams);
                    if (cntRows.Count > 0 && cntRows[0].TryGetValue("cnt", out var v) && v != null)
                        cnt = Convert.ToInt64(v);

                    if (isXlsx && cnt > 1048576)
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
            }
            else
            {
                // skipCount 情况下：若导出 Excel 且无 LIMIT，给一次轻提示（避免用户无意导出超大表）
                if (isXlsx && limit <= 0 && !hasLimit)
                {
                    var dr = MessageBox.Show(
                        "当前已开启“跳过预统计”，将直接导出。\n\n若结果行数超过 Excel 上限（1,048,576）会导出失败，建议改用 CSV 或加 LIMIT/过滤条件。\n\n是否继续？",
                        "导出提示",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);
                    if (dr != DialogResult.Yes) return;
                }
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
                StartExportJob(exportSql, outPath, openAfter, format: "csv", forceAllText: false, expectedRows: cnt, parameters: sqlParams);
            }
            else
            {
                // 说明：ClosedXML 为内存型写入，超大结果建议 CSV（避免内存飙升/卡死）
                if (cnt > 200000 && !forceAllText)
                {
                    var dr = MessageBox.Show(
                        $"预计导出 {cnt:n0} 行，Excel 导出可能非常慢且占用大量内存。\n\n建议改用 CSV（更快更稳）。\n\n是否改用 CSV 导出？",
                        "导出建议",
                        MessageBoxButtons.YesNoCancel,
                        MessageBoxIcon.Warning);
                    if (dr == DialogResult.Cancel) return;
                    if (dr == DialogResult.Yes)
                    {
                        // 改为 CSV
                        var csvPath = Path.ChangeExtension(outPath, ".csv");
                        StartExportJob(exportSql, csvPath, openAfter, format: "csv", forceAllText: false, expectedRows: cnt, parameters: sqlParams);
                        return;
                    }
                }
                StartExportJob(exportSql, outPath, openAfter, format: "xlsx", forceAllText: forceAllText, expectedRows: cnt, parameters: sqlParams);
            }
        }
        catch (Exception ex)
        {
            WriteErrorLog("导出查询结果失败", ex);
            SendMessageToWebView(new { action = "error", message = $"导出失败: {ex.Message}", hasErrorLog = true });
        }
    }

    private void StartExportJob(string sql, string outPath, bool openAfter, string format, bool forceAllText, long expectedRows, IDictionary<string, object?>? parameters)
    {
        // 取消上一个导出任务
        try { _exportCts?.Cancel(); } catch { }
        try { _exportCts?.Dispose(); } catch { }
        _exportCts = new CancellationTokenSource();
        _exportJobId = Guid.NewGuid().ToString("N");
        var jobId = _exportJobId;
        var token = _exportCts.Token;

        SendMessageToWebView(new
        {
            action = "exportProgress",
            jobId,
            stage = "starting",
            processed = 0L,
            total = expectedRows,
            percent = expectedRows > 0 ? 0 : (int?)null,
            filePath = outPath
        });

        Task.Run(() =>
        {
            try
            {
                if (_sqliteManager == null) throw new InvalidOperationException("SQLite管理器未初始化");

                // 为后台任务创建独立连接：文件库直接复用连接串；内存库先备份到临时文件
                string connStr = _sqliteManager.ConnectionString;
                if (connStr.Contains(":memory:", StringComparison.OrdinalIgnoreCase))
                {
                    var tmpDb = Path.Combine(Path.GetTempPath(), $"export-{DateTime.Now:yyyyMMdd-HHmmss}-{Guid.NewGuid():N}.db");
                    using var dest = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={tmpDb}");
                    dest.Open();
                    _sqliteManager.Connection?.BackupDatabase(dest);
                    dest.Close();
                    connStr = $"Data Source={tmpDb}";
                }

                using var conn = new Microsoft.Data.Sqlite.SqliteConnection(connStr);
                conn.Open();

                Action<long> report = (processed) =>
                {
                    try
                    {
                        var pct = expectedRows > 0 ? (int)Math.Min(100, Math.Round(processed * 100.0 / expectedRows)) : (int?)null;
                        SendMessageToWebView(new
                        {
                            action = "exportProgress",
                            jobId,
                            stage = "running",
                            processed,
                            total = expectedRows,
                            percent = pct,
                            filePath = outPath
                        });
                    }
                    catch { }
                };

                if (string.Equals(format, "csv", StringComparison.OrdinalIgnoreCase))
                {
                    ExportSqlToCsvWithProgress(conn, sql, parameters, outPath, protectForExcel: true, token, report);
                }
                else
                {
                    ExportSqlToXlsxWithProgress(conn, sql, parameters, outPath, forceAllText, token, report);
                }

                SendMessageToWebView(new { action = "exportComplete", jobId, ok = true, filePath = outPath });

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
            catch (OperationCanceledException)
            {
                try { SendMessageToWebView(new { action = "exportCancelled", jobId, ok = false, message = "已取消导出" }); } catch { }
            }
            catch (Exception ex)
            {
                try { WriteErrorLog("导出查询结果失败（后台任务）", ex); } catch { }
                try { SendMessageToWebView(new { action = "exportComplete", jobId, ok = false, message = ex.Message, hasErrorLog = true }); } catch { }
            }
        }, token);
    }

    private static void AddSqlParams(Microsoft.Data.Sqlite.SqliteCommand cmd, IDictionary<string, object?>? parameters)
    {
        if (parameters == null || parameters.Count == 0) return;
        foreach (var kv in parameters)
        {
            var key0 = (kv.Key ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(key0)) continue;
            var val = kv.Value ?? DBNull.Value;
            var bare = key0.TrimStart(':', '@', '$');
            if (string.IsNullOrWhiteSpace(bare)) continue;

            void AddOne(string name)
            {
                try
                {
                    if (cmd.Parameters.Contains(name)) return;
                    cmd.Parameters.AddWithValue(name, val);
                }
                catch { }
            }

            AddOne("@" + bare);
            AddOne(":" + bare);
            AddOne("$" + bare);
        }
    }

    private void ExportSqlToCsvWithProgress(
        Microsoft.Data.Sqlite.SqliteConnection conn,
        string sql,
        IDictionary<string, object?>? parameters,
        string outputPath,
        bool protectForExcel,
        CancellationToken token,
        Action<long> report)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        try { AddSqlParams(cmd, parameters); } catch { }
        using var reader = cmd.ExecuteReader();
        using var writer = new StreamWriter(outputPath, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));

        var colCount = reader.FieldCount;
        for (int i = 0; i < colCount; i++)
        {
            if (i > 0) writer.Write(",");
            writer.Write(CsvEscape(reader.GetName(i)));
        }
        writer.WriteLine();

        long processed = 0;
        while (reader.Read())
        {
            token.ThrowIfCancellationRequested();
            for (int i = 0; i < colCount; i++)
            {
                if (i > 0) writer.Write(",");
                var v = reader.IsDBNull(i) ? "" : Convert.ToString(reader.GetValue(i));
                var cell = v ?? "";
                if (protectForExcel && LooksLikeSensitiveNumber(cell))
                {
                    cell = "=\"" + cell.Replace("\"", "\"\"") + "\"";
                }
                writer.Write(CsvEscape(cell));
            }
            writer.WriteLine();
            processed++;
            if (processed % 5000 == 0) report(processed);
        }
        report(processed);
    }

    private void ExportSqlToXlsxWithProgress(
        Microsoft.Data.Sqlite.SqliteConnection conn,
        string sql,
        IDictionary<string, object?>? parameters,
        string outputPath,
        bool forceAllText,
        CancellationToken token,
        Action<long> report)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        try { AddSqlParams(cmd, parameters); } catch { }
        using var reader = cmd.ExecuteReader();
        var colCount = reader.FieldCount;

        var wb = new ClosedXML.Excel.XLWorkbook();
        int sheetIdx = 1;
        ClosedXML.Excel.IXLWorksheet ws = wb.Worksheets.Add($"Query{sheetIdx}");
        ApplyExcelReportDefaults(ws);

        void WriteHeader()
        {
            for (int c = 0; c < colCount; c++)
            {
                ws.Cell(1, c + 1).Value = reader.GetName(c);
                ws.Cell(1, c + 1).Style.Font.Bold = true;
            }
            ws.Row(1).Height = 20;
        }
        WriteHeader();

        int row = 2;
        long processed = 0;
        while (reader.Read())
        {
            token.ThrowIfCancellationRequested();
            // Excel 单表最大行数 1,048,576（含表头）—— 超出则自动分 Sheet
            if (row > 1048576)
            {
                try { ApplyTextLeftNumberRightAlignment(ws, headerRow: 1); } catch { }
                sheetIdx++;
                ws = wb.Worksheets.Add($"Query{sheetIdx}");
                ApplyExcelReportDefaults(ws);
                WriteHeader();
                row = 2;
            }
            for (int c = 0; c < colCount; c++)
            {
                var v = reader.IsDBNull(c) ? null : reader.GetValue(c);
                if (v == null)
                {
                    ws.Cell(row, c + 1).Value = "";
                    continue;
                }
                if (forceAllText)
                {
                    var s = Convert.ToString(v) ?? "";
                    ws.Cell(row, c + 1).SetValue(s);
                    ws.Cell(row, c + 1).Style.NumberFormat.Format = "@";
                }
                else if (v is string s)
                {
                    // 关键：字符串类型一律按文本写入，不再“尝试转数字”
                    // 否则像 chargeCode / serviceCode 这类编码会被 Excel 变成数值（丢前导0/格式变化）
                    ws.Cell(row, c + 1).SetValue(s);
                    if (LooksLikeSensitiveNumber(s))
                        ws.Cell(row, c + 1).Style.NumberFormat.Format = "@";
                }
                else
                {
                    // 非字符串：尽量保留数值/日期类型（便于排序/统计）
                    try
                    {
                        if (v is long || v is int || v is short || v is byte || v is sbyte || v is uint || v is ulong || v is ushort)
                        {
                            ws.Cell(row, c + 1).Value = Convert.ToInt64(v);
                        }
                        else if (v is float || v is double || v is decimal)
                        {
                            ws.Cell(row, c + 1).Value = Convert.ToDouble(v);
                        }
                        else if (v is DateTime dt)
                        {
                            ws.Cell(row, c + 1).Value = dt;
                        }
                        else
                        {
                            ws.Cell(row, c + 1).SetValue(Convert.ToString(v) ?? "");
                        }
                    }
                    catch
                    {
                        ws.Cell(row, c + 1).SetValue(Convert.ToString(v) ?? "");
                    }
                }
            }
            row++;
            processed++;
            if (processed % 2000 == 0) report(processed);
        }
        report(processed);

        try { ApplyTextLeftNumberRightAlignment(ws, headerRow: 1); } catch { }
        try
        {
            foreach (var x in wb.Worksheets)
            {
                try { x.Columns().AdjustToContents(); } catch { }
            }
        }
        catch { }
        wb.SaveAs(outputPath);
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
            int sheetIdx = 1;
            var ws = wb.Worksheets.Add($"Sheet{sheetIdx}");
            ApplyExcelReportDefaults(ws);

            void WriteHeader()
            {
                for (int i = 0; i < cols.Count; i++)
                {
                    ws.Cell(1, i + 1).Value = cols[i];
                    ws.Cell(1, i + 1).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#343a40");
                    ws.Cell(1, i + 1).Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
                    ws.Cell(1, i + 1).Style.Font.Bold = true;
                }
                ws.Row(1).Height = 20;
            }
            WriteHeader();

            int rIdx = 2;
            foreach (var rowEl in rowsEl.EnumerateArray())
            {
                // Excel 单表最大行数 1,048,576（含表头）—— 超出则自动分 Sheet
                if (rIdx > 1048576)
                {
                    try { ApplyTextLeftNumberRightAlignment(ws, headerRow: 1); } catch { }
                    sheetIdx++;
                    ws = wb.Worksheets.Add($"Sheet{sheetIdx}");
                    ApplyExcelReportDefaults(ws);
                    WriteHeader();
                    rIdx = 2;
                }
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
            try { ApplyTextLeftNumberRightAlignment(ws, headerRow: 1); } catch { }
            try
            {
                foreach (var x in wb.Worksheets)
                {
                    try { x.Columns().AdjustToContents(); } catch { }
                }
            }
            catch { }
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

            SendMessageToWebView(new { action = "htmlReportSaved", filePath = fullPath, reportName = reportName, fileName = fileName });

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
                ApplyExcelReportDefaults(ws);
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
                try { ApplyTextLeftNumberRightAlignment(ws, headerRow: 1); } catch { }
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

            // raw+clean 策略：默认清洗基于 raw__ 主表，输出到 clean__ 主表
            try
            {
                if (!string.IsNullOrWhiteSpace(_activeSchemeId))
                {
                    var meta = LoadOrCreateProjectMeta(_activeSchemeId!, displayName: null, settingsJson: null);
                    if (string.Equals(meta.RawCleanPolicy?.Mode, "raw+clean", StringComparison.OrdinalIgnoreCase))
                    {
                        var rawTable = $"raw__{sourceTable}";
                        var cleanTable = $"clean__{sourceTable}";

                        // 若 raw 表不存在，则从当前 sourceTable 快照一份 raw
                        if (!_sqliteManager.TableExists(rawTable))
                        {
                            try
                            {
                                _sqliteManager.Execute($"DROP TABLE IF EXISTS {SqliteManager.QuoteIdent(rawTable)};");
                                _sqliteManager.Execute($"CREATE TABLE {SqliteManager.QuoteIdent(rawTable)} AS SELECT * FROM {SqliteManager.QuoteIdent(sourceTable)};");
                            }
                            catch { }
                        }

                        // 默认情况下：清洗源固定使用 raw 表（避免多次清洗叠加）
                        sourceTable = rawTable;

                        // 若前端未显式指定 targetTable（仍是默认 Cleansed），则输出到 clean__*
                        if (string.Equals(targetTable, "Cleansed", StringComparison.OrdinalIgnoreCase))
                            targetTable = cleanTable;

                        // raw+clean 模式下不建议 db_overwrite：强制改为生成 clean 表
                        if (settings.TryGetProperty("outputOptions", out var oo2)
                            && oo2.ValueKind != System.Text.Json.JsonValueKind.Undefined
                            && oo2.TryGetProperty("target", out var ot2)
                            && string.Equals(ot2.GetString(), "db_overwrite", StringComparison.OrdinalIgnoreCase))
                        {
                            // 直接把 targetTable 改为 clean 表，后续会走“else 分支”（CREATE TABLE AS）
                            targetTable = cleanTable;
                        }
                    }
                }
            }
            catch { }

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

            // raw+clean 模式下禁止 db_overwrite：自动改为 db_copy（输出 clean__*）
            if (string.Equals(outputTarget, "db_overwrite", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(_activeSchemeId))
                    {
                        var meta = LoadOrCreateProjectMeta(_activeSchemeId!, displayName: null, settingsJson: null);
                        if (string.Equals(meta.RawCleanPolicy?.Mode, "raw+clean", StringComparison.OrdinalIgnoreCase))
                        {
                            outputTarget = "db_copy";
                        }
                    }
                }
                catch { }
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

            // 更新项目元数据（DbTables/大小/主表建议等）
            try
            {
                if (!string.IsNullOrWhiteSpace(_activeSchemeId))
                {
                    var meta = LoadOrCreateProjectMeta(_activeSchemeId!, displayName: null, settingsJson: null);
                    meta.DbTables = GetDbTablesForScheme(_activeSchemeId!);
                    meta.DbSizeBytes = SafeFileSize(meta.DbPath);
                    meta.UpdatedAt = DateTime.Now;
                    SaveProjectMeta(meta);
                }
            }
            catch { }

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

            bool compareFields = true, compareData = true, compareTypes = false;
            bool ignoreCase = false, ignoreWhitespace = false, showOnlyDiff = true;
            try
            {
                if (settings.TryGetProperty("compareTypes", out var ct))
                {
                    compareFields = ct.TryGetProperty("fields", out var v) ? v.GetBoolean() : true;
                    compareData = ct.TryGetProperty("data", out var v2) ? v2.GetBoolean() : true;
                    compareTypes = ct.TryGetProperty("types", out var v3) ? v3.GetBoolean() : false;
                }
                if (settings.TryGetProperty("advancedOptions", out var ao))
                {
                    ignoreCase = ao.TryGetProperty("ignoreCase", out var v) ? v.GetBoolean() : false;
                    ignoreWhitespace = ao.TryGetProperty("ignoreWhitespace", out var v2) ? v2.GetBoolean() : false;
                    showOnlyDiff = ao.TryGetProperty("showOnlyDiff", out var v3) ? v3.GetBoolean() : true;
                }
            }
            catch { /* ignore */ }

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
            if (compareFields)
            {
                foreach (var c in beforeCols.Except(afterCols, StringComparer.OrdinalIgnoreCase))
                    fieldDiffs.Add(new { type = "缺失字段", field = c, before = "存在", after = "-" });
                foreach (var c in afterCols.Except(beforeCols, StringComparer.OrdinalIgnoreCase))
                    fieldDiffs.Add(new { type = "新增字段", field = c, before = "-", after = "存在" });
            }

            if (compareTypes)
            {
                try
                {
                    var beforeTypeMap = beforeSchema.ToDictionary(x => x.ColumnName, x => x.DataType ?? "", StringComparer.OrdinalIgnoreCase);
                    var afterTypeMap = afterSchema.ToDictionary(x => x.ColumnName, x => x.DataType ?? "", StringComparer.OrdinalIgnoreCase);
                    foreach (var c in commonCols)
                    {
                        beforeTypeMap.TryGetValue(c, out var bt);
                        afterTypeMap.TryGetValue(c, out var atp);
                        bt ??= "";
                        atp ??= "";
                        if (!string.Equals(bt, atp, StringComparison.OrdinalIgnoreCase))
                            fieldDiffs.Add(new { type = "类型变更", field = c, before = bt, after = atp });
                    }
                }
                catch { }
            }

            int beforeRows = _sqliteManager.GetRowCount(beforeTable);
            int afterRows = _sqliteManager.GetRowCount(afterTable);

            var dataDiffs = new List<object>();
            long dataDiffCount = 0;
            if (compareData && commonCols.Count > 0)
            {
                // 使用归一化表达式进行比较（忽略大小写/空白）
                string BuildNormSelect(string table)
                {
                    var exprs = commonCols.Select(c => $"{BuildCompareNormalizeExpr(c, ignoreCase, ignoreWhitespace)} AS {SqlIdent(c)}");
                    return $"SELECT {string.Join(", ", exprs)} FROM {SqlIdent(table)}";
                }

                var beforeSel = BuildNormSelect(beforeTable);
                var afterSel = BuildNormSelect(afterTable);
                var onlyBeforeCnt = SqlScalarLong($"SELECT COUNT(*) AS v FROM ({beforeSel} EXCEPT {afterSel}) t");
                var onlyAfterCnt = SqlScalarLong($"SELECT COUNT(*) AS v FROM ({afterSel} EXCEPT {beforeSel}) t");
                dataDiffCount = onlyBeforeCnt + onlyAfterCnt;

                // 抽样展示（用整行JSON字符串展示，避免复杂“主键对齐”）
                var sampleRows = _sqliteManager.Query($"SELECT * FROM ({beforeSel} EXCEPT {afterSel}) t LIMIT 5");
                int idx = 1;
                foreach (var row in sampleRows)
                {
                    dataDiffs.Add(new { rowNumber = idx++, fieldName = "*", beforeValue = System.Text.Json.JsonSerializer.Serialize(row), afterValue = "", diffType = "仅清洗前存在" });
                }
                var sampleRows2 = _sqliteManager.Query($"SELECT * FROM ({afterSel} EXCEPT {beforeSel}) t LIMIT 5");
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
                    compareTime = Math.Round(sw.Elapsed.TotalSeconds, 2),
                    options = new
                    {
                        compareFields,
                        compareData,
                        compareTypes,
                        ignoreCase,
                        ignoreWhitespace,
                        showOnlyDiff
                    }
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

    private static string BuildCompareNormalizeExpr(string colName, bool ignoreCase, bool ignoreWhitespace)
    {
        // 统一按 TEXT 比较：避免数值/日期类型差异导致 EXCEPT 结果不稳定
        var baseExpr = $"COALESCE(CAST({SqlIdent(colName)} AS TEXT),'')";
        if (ignoreWhitespace)
        {
            // 去除空格 + 常见换行/制表符
            baseExpr = $"REPLACE(REPLACE(REPLACE(REPLACE({baseExpr}, ' ', ''), char(9), ''), char(10), ''), char(13), '')";
        }
        if (ignoreCase)
        {
            baseExpr = $"UPPER({baseExpr})";
        }
        return baseExpr;
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
        if (t.Length == 0) return false;

        // 纯数字：更容易被 Excel 自动当数值/科学计数/丢前导0
        if (t.All(char.IsDigit))
        {
            // 前导0：编码字段（如 000123）必须按文本导出
            if (t.StartsWith("0")) return true;
            // 常见编码长度：>=6 基本可视为“编码/ID”场景（严格保留文本）
            if (t.Length >= 6) return true;
            // Excel 15位精度风险
            if (t.Length >= 15) return true;
        }
        return false;
    }

    // ====================== 导出报表样式规范（微软雅黑9号字，文本左/数字右） ======================

    private const string DefaultExcelFontName = "Microsoft YaHei";
    private const double DefaultExcelFontSize = 9;

    private static void ApplyExcelReportDefaults(ClosedXML.Excel.IXLWorksheet ws)
    {
        ws.Style.Font.FontName = DefaultExcelFontName;
        ws.Style.Font.FontSize = DefaultExcelFontSize;
        ws.Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Center;
        ws.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Left;
    }

    private static bool HeaderPreferRightAlign(string header)
    {
        var h = (header ?? "").Trim();
        if (h.Length == 0) return false;
        var lower = h.ToLowerInvariant();
        // 序号/ID/编码/金额/数量/比例等，统一按“数值对比”场景右对齐
        if (lower.Contains("id") || h.Contains("编号") || h.Contains("序号") || h.Contains("编码")) return true;
        if (h.Contains("金额") || h.Contains("价格") || h.Contains("成本") || h.Contains("收入") || h.Contains("支出")) return true;
        if (h.Contains("数量") || h.Contains("数") || h.Contains("行数") || h.Contains("次数") || h.Contains("比率") || h.Contains("比例") || h.Contains("%") || h.Contains("率")) return true;
        return false;
    }

    private static bool LooksLikeNumberText(string s)
    {
        var t = (s ?? "").Trim();
        if (t.Length == 0) return false;
        // 允许：货币符号、千分位、负号、小数、百分号
        // 示例：¥1,200.50 / 1,024 / 99.5% / 001
        // 快速判定：至少包含一个数字，且字符集只来自 [数字,逗号,点,负号,空格,货币符号,百分号]
        if (!t.Any(char.IsDigit)) return false;
        foreach (var ch in t)
        {
            if (char.IsDigit(ch)) continue;
            if (ch == ',' || ch == '.' || ch == '-' || ch == ' ' || ch == '%' || ch == '¥' || ch == '￥' || ch == '$') continue;
            return false;
        }
        return true;
    }

    private static void ApplyTextLeftNumberRightAlignment(ClosedXML.Excel.IXLWorksheet ws, int headerRow = 1, int sampleLimit = 200)
    {
        var used = ws.RangeUsed();
        if (used == null) return;
        var firstRow = used.RangeAddress.FirstAddress.RowNumber;
        var firstCol = used.RangeAddress.FirstAddress.ColumnNumber;
        var lastRow = used.RangeAddress.LastAddress.RowNumber;
        var lastCol = used.RangeAddress.LastAddress.ColumnNumber;

        if (lastRow < headerRow) return;

        for (int c = firstCol; c <= lastCol; c++)
        {
            var header = ws.Cell(headerRow, c).GetString();
            if (HeaderPreferRightAlign(header))
            {
                ws.Column(c).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Right;
                continue;
            }

            int numericLike = 0, textLike = 0;
            int start = Math.Max(headerRow + 1, firstRow);
            int end = Math.Min(lastRow, headerRow + sampleLimit);
            for (int r = start; r <= end; r++)
            {
                var cell = ws.Cell(r, c);
                if (cell.IsEmpty()) continue;
                if (cell.DataType == ClosedXML.Excel.XLDataType.Number)
                {
                    numericLike++;
                }
                else
                {
                    var s = cell.GetString();
                    if (LooksLikeNumberText(s)) numericLike++;
                    else textLike++;
                }
            }

            // 混合列按文本处理；纯数字列/以数字为主列右对齐
            if (numericLike > 0 && textLike == 0)
                ws.Column(c).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Right;
            else if (numericLike >= Math.Max(4, textLike * 2))
                ws.Column(c).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Right;
            else
                ws.Column(c).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Left;
        }

        // 表头统一左对齐（避免列样式把表头也右对齐）
        try { ws.Row(headerRow).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Left; } catch { }
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
        int sheetIdx = 1;
        var ws = wb.Worksheets.Add($"Query{sheetIdx}");

        // 全局样式：微软雅黑9号字（企业报表默认），文本左/数字右（后续按列自动矫正）
        ApplyExcelReportDefaults(ws);

        void WriteHeader()
        {
            for (int c = 0; c < colCount; c++)
            {
                ws.Cell(1, c + 1).Value = reader.GetName(c);
                ws.Cell(1, c + 1).Style.Font.Bold = true;
            }
            ws.Row(1).Height = 20;
        }
        WriteHeader();

        int row = 2;
        while (reader.Read())
        {
            if (row > 1048576)
            {
                try { ApplyTextLeftNumberRightAlignment(ws, headerRow: 1); } catch { }
                sheetIdx++;
                ws = wb.Worksheets.Add($"Query{sheetIdx}");
                ApplyExcelReportDefaults(ws);
                WriteHeader();
                row = 2;
            }
            for (int c = 0; c < colCount; c++)
            {
                var v = reader.IsDBNull(c) ? null : reader.GetValue(c);
                if (v == null)
                {
                    ws.Cell(row, c + 1).Value = "";
                    continue;
                }
                if (forceAllText)
                {
                    // 全列文本：任何值都按文本写，避免科学计数法/前导0/精度截断
                    var s = Convert.ToString(v) ?? "";
                    ws.Cell(row, c + 1).SetValue(s);
                    ws.Cell(row, c + 1).Style.NumberFormat.Format = "@";
                }
                else if (v is string s)
                {
                    // 字符串类型保持原样：不转数值
                    ws.Cell(row, c + 1).SetValue(s);
                    if (LooksLikeSensitiveNumber(s))
                        ws.Cell(row, c + 1).Style.NumberFormat.Format = "@";
                }
                else
                {
                    // 非字符串：尽量保留数值/日期类型（统计/排序更友好）
                    try
                    {
                        if (v is long || v is int || v is short || v is byte || v is sbyte || v is uint || v is ulong || v is ushort)
                        {
                            ws.Cell(row, c + 1).Value = Convert.ToInt64(v);
                        }
                        else if (v is float || v is double || v is decimal)
                        {
                            ws.Cell(row, c + 1).Value = Convert.ToDouble(v);
                        }
                        else if (v is DateTime dt)
                        {
                            ws.Cell(row, c + 1).Value = dt;
                        }
                        else
                        {
                            ws.Cell(row, c + 1).SetValue(Convert.ToString(v) ?? "");
                        }
                    }
                    catch
                    {
                        ws.Cell(row, c + 1).SetValue(Convert.ToString(v) ?? "");
                    }
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

        // 列对齐：文本居左、数字居右（含 ID/序号/金额/比例等按表头优先右对齐）
        try { ApplyTextLeftNumberRightAlignment(ws, headerRow: 1); } catch { }

        // 列宽：自适应 + 合理上限（避免超长文本拉爆）
        try
        {
            foreach (var x in wb.Worksheets)
            {
                try { x.Columns().AdjustToContents(); } catch { }
            }
        }
        catch { }
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

    // ==================== 脱敏：Vault/Policy 初始化与执行 ====================

    private void InitVaultDb(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            var cfg = EnsureDesensitizationConfigForActiveScheme();
            EnsureVaultReady(cfg);
            // namespace 注册（幂等）
            using (var cmd = _vaultConn!.CreateCommand())
            {
                cmd.CommandText = "INSERT OR IGNORE INTO namespace_registry(namespace, display_name, created_at, created_by) VALUES($ns,$dn,datetime('now'),$by);";
                cmd.Parameters.AddWithValue("$ns", cfg.Namespace);
                cmd.Parameters.AddWithValue("$dn", _activeSchemeId ?? cfg.Namespace);
                cmd.Parameters.AddWithValue("$by", Environment.UserName ?? "user");
                cmd.ExecuteNonQuery();
            }
            SendMessageToWebView(new
            {
                action = "vaultDbReady",
                requestId,
                ok = true,
                data = new { vaultDbPath = (_vaultDbPath ?? "").Replace('\\', '/'), @namespace = cfg.Namespace }
            });
        }
        catch (Exception ex)
        {
            WriteErrorLog("InitVaultDb失败", ex);
            SendMessageToWebView(new { action = "vaultDbReady", requestId, ok = false, message = ex.Message, errorCode = "E_SQLITE_FAIL" });
        }
    }

    private void InitPolicyRepoDb(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            var cfg = EnsureDesensitizationConfigForActiveScheme();
            EnsurePolicyRepoReady(cfg);
            SendMessageToWebView(new
            {
                action = "policyRepoReady",
                requestId,
                ok = true,
                data = new { policyDbPath = (_policyDbPath ?? "").Replace('\\', '/'), @namespace = cfg.Namespace }
            });
        }
        catch (Exception ex)
        {
            WriteErrorLog("InitPolicyRepoDb失败", ex);
            SendMessageToWebView(new { action = "policyRepoReady", requestId, ok = false, message = ex.Message, errorCode = "E_SQLITE_FAIL" });
        }
    }

    private void SeedDefaultTemplates(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            var cfg = EnsureDesensitizationConfigForActiveScheme();
            EnsurePolicyRepoReady(cfg);
            ExecSqlScript(_policyConn!, DefaultTemplateSeedSql);

            long tplCnt = _policyConn!.ExecuteScalar<long>("SELECT COUNT(*) FROM template;");
            long ruleCnt = _policyConn!.ExecuteScalar<long>("SELECT COUNT(*) FROM template_rule;");
            SendMessageToWebView(new { action = "defaultTemplatesSeeded", requestId, ok = true, data = new { templateCount = tplCnt, ruleCount = ruleCnt } });
        }
        catch (Exception ex)
        {
            WriteErrorLog("SeedDefaultTemplates失败", ex);
            SendMessageToWebView(new { action = "defaultTemplatesSeeded", requestId, ok = false, message = ex.Message, errorCode = "E_SQLITE_FAIL" });
        }
    }

    private void DsSeedEnterpriseTemplates(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            var cfg = EnsureDesensitizationConfigForActiveScheme();
            EnsurePolicyRepoReady(cfg);
            var payload = data.TryGetProperty("payload", out var pl) ? pl : data;
            var kind = (payload.TryGetProperty("kind", out var kd) ? kd.GetString() : null) ?? "enterprise_basic";
            var sql = string.Equals(kind, "enterprise_extended", StringComparison.OrdinalIgnoreCase)
                ? EnterpriseExtendedTemplateSeedSql
                : EnterpriseBasicTemplateSeedSql;

            ExecSqlScript(_policyConn!, sql);

            long tplCnt = _policyConn!.ExecuteScalar<long>("SELECT COUNT(*) FROM template;");
            long ruleCnt = _policyConn!.ExecuteScalar<long>("SELECT COUNT(*) FROM template_rule;");
            SendMessageToWebView(new { action = "dsEnterpriseTemplatesSeeded", requestId, ok = true, data = new { templateCount = tplCnt, ruleCount = ruleCnt, kind } });
        }
        catch (Exception ex)
        {
            WriteErrorLog("DsSeedEnterpriseTemplates失败", ex);
            SendMessageToWebView(new { action = "dsEnterpriseTemplatesSeeded", requestId, ok = false, message = ex.Message, errorCode = "E_SQLITE_FAIL" });
        }
    }

    private void DsTemplateMatchPreview(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            if (_sqliteManager?.Connection == null) throw new InvalidOperationException("SQLite未就绪");
            var cfg = EnsureDesensitizationConfigForActiveScheme();
            EnsurePolicyRepoReady(cfg);

            var payload = data.TryGetProperty("payload", out var pl) ? pl : data;
            var templateId = (payload.TryGetProperty("templateId", out var tid) ? tid.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(templateId)) throw new InvalidOperationException("templateId 不能为空");
            var table = (payload.TryGetProperty("table", out var tb) ? tb.GetString() : null) ?? MainTableNameOrDefault();

            var rules = _policyConn!.Query(@"
SELECT id, column_name, match_mode, match_pattern, data_type, action
FROM template_rule WHERE template_id=@id AND enabled=1 ORDER BY sort_order;", new { id = templateId })
                .Select(r => new
                {
                    id = (string)r.id,
                    columnName = (string)r.column_name,
                    matchMode = (string?)r.match_mode ?? "exact",
                    matchPattern = (string?)r.match_pattern,
                    dataType = (string)r.data_type,
                    action = (string)r.action
                })
                .ToList();

            bool Match(string col, dynamic rr)
            {
                string mode = ((string)rr.matchMode ?? "exact").Trim().ToLowerInvariant();
                string pat = string.IsNullOrWhiteSpace((string?)rr.matchPattern) ? (string)rr.columnName : (string)rr.matchPattern;
                if (mode == "contains") return col.Contains(pat, StringComparison.OrdinalIgnoreCase);
                if (mode == "regex")
                {
                    try { return System.Text.RegularExpressions.Regex.IsMatch(col, pat, System.Text.RegularExpressions.RegexOptions.IgnoreCase); }
                    catch { return false; }
                }
                return string.Equals(col, (string)rr.columnName, StringComparison.OrdinalIgnoreCase);
            }

            var schema = _sqliteManager.GetTableSchema(table);
            if (schema == null || schema.Count == 0) throw new InvalidOperationException("表不存在或无字段");
            var cols = new List<object>();
            foreach (var c in schema)
            {
                var col = c.ColumnName;
                var hit = rules.FirstOrDefault(r => Match(col, r));
                cols.Add(new
                {
                    column = col,
                    matched = hit != null,
                    ruleId = hit?.id,
                    dataType = hit?.dataType,
                    action = hit?.action,
                    matchMode = hit?.matchMode,
                    matchPattern = hit?.matchPattern
                });
            }

            SendMessageToWebView(new { action = "dsTemplateMatchPreviewLoaded", requestId, ok = true, data = new { table, templateId, matchedColumns = cols, ruleCount = rules.Count } });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "dsTemplateMatchPreviewLoaded", requestId, ok = false, message = ex.Message, errorCode = "E_SQLITE_FAIL" });
        }
    }

    // ==================== 枚举值存档（字段画像 / 可缓存） ====================

    private string GetEnumArchiveDir(string schemeId)
    {
        var db = GetSchemeDbPath(schemeId);
        var dir = Path.GetDirectoryName(db) ?? GetDbDir();
        var outDir = Path.Combine(dir, "profiles");
        Directory.CreateDirectory(outDir);
        return outDir;
    }

    private void DsEnumArchiveCancel(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            _enumArchiveCts?.Cancel();
            SendMessageToWebView(new { action = "dsEnumArchiveCanceled", requestId, ok = true });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "dsEnumArchiveCanceled", requestId, ok = false, message = ex.Message });
        }
    }

    private void DsEnumArchiveList(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            if (string.IsNullOrWhiteSpace(_activeSchemeId)) throw new InvalidOperationException("未打开项目");
            var dir = GetEnumArchiveDir(_activeSchemeId!);
            var files = Directory.GetFiles(dir, "enum_*.json")
                .Select(p => new FileInfo(p))
                .OrderByDescending(f => f.LastWriteTimeUtc)
                .Take(500)
                .Select(f => new { name = f.Name, path = f.FullName.Replace('\\', '/'), size = f.Length, time = f.LastWriteTime.ToString("s") })
                .ToList();
            SendMessageToWebView(new { action = "dsEnumArchiveListLoaded", requestId, ok = true, data = new { files } });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "dsEnumArchiveListLoaded", requestId, ok = false, message = ex.Message, errorCode = "E_IO_FAIL" });
        }
    }

    private void DsEnumArchiveLoad(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            if (string.IsNullOrWhiteSpace(_activeSchemeId)) throw new InvalidOperationException("未打开项目");
            var payload = data.TryGetProperty("payload", out var pl) ? pl : data;
            var fileName = (payload.TryGetProperty("fileName", out var fn) ? fn.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(fileName)) throw new InvalidOperationException("fileName 不能为空");
            var dir = GetEnumArchiveDir(_activeSchemeId!);
            var p = Path.Combine(dir, fileName);
            if (!File.Exists(p)) throw new FileNotFoundException("存档不存在", p);
            var json = File.ReadAllText(p, Encoding.UTF8);
            SendMessageToWebView(new { action = "dsEnumArchiveLoaded", requestId, ok = true, data = new { fileName, json } });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "dsEnumArchiveLoaded", requestId, ok = false, message = ex.Message, errorCode = "E_IO_FAIL" });
        }
    }

    private void DsEnumArchiveDelete(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            if (string.IsNullOrWhiteSpace(_activeSchemeId)) throw new InvalidOperationException("未打开项目");
            var payload = data.TryGetProperty("payload", out var pl) ? pl : data;
            var fileName = (payload.TryGetProperty("fileName", out var fn) ? fn.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(fileName)) throw new InvalidOperationException("fileName 不能为空");
            var dir = GetEnumArchiveDir(_activeSchemeId!);
            var p = Path.Combine(dir, fileName);
            if (!File.Exists(p)) throw new FileNotFoundException("存档不存在", p);
            File.Delete(p);
            SendMessageToWebView(new { action = "dsEnumArchiveDeleted", requestId, ok = true, data = new { fileName } });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "dsEnumArchiveDeleted", requestId, ok = false, message = ex.Message, errorCode = "E_IO_FAIL" });
        }
    }

    private void DsEnumArchiveOpenFolder(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            if (string.IsNullOrWhiteSpace(_activeSchemeId)) throw new InvalidOperationException("未打开项目");
            var dir = GetEnumArchiveDir(_activeSchemeId!);
            try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(dir) { UseShellExecute = true }); } catch { }
            SendMessageToWebView(new { action = "dsEnumArchiveFolderOpened", requestId, ok = true, data = new { dir = dir.Replace('\\', '/') } });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "dsEnumArchiveFolderOpened", requestId, ok = false, message = ex.Message, errorCode = "E_IO_FAIL" });
        }
    }

    private void DsEnumArchiveStart(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            if (_sqliteManager?.Connection == null) throw new InvalidOperationException("SQLite未就绪");
            if (string.IsNullOrWhiteSpace(_activeSchemeId)) throw new InvalidOperationException("未打开项目");

            // 取消旧任务
            try { _enumArchiveCts?.Cancel(); } catch { }
            _enumArchiveCts = new CancellationTokenSource();
            var token = _enumArchiveCts.Token;

            var payload = data.TryGetProperty("payload", out var pl) ? pl : data;
            var table = (payload.TryGetProperty("table", out var tb) ? tb.GetString() : null) ?? MainTableNameOrDefault();
            if (string.IsNullOrWhiteSpace(table)) throw new InvalidOperationException("table 不能为空");
            var cols = new List<string>();
            if (payload.TryGetProperty("columns", out var cs) && cs.ValueKind == System.Text.Json.JsonValueKind.Array)
                cols = cs.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
            if (cols.Count == 0) throw new InvalidOperationException("columns 不能为空");

            int sampleRows = 50000;
            int maxValues = 200;
            bool withFreq = payload.TryGetProperty("withFreq", out var wf) && wf.ValueKind == System.Text.Json.JsonValueKind.True;
            try { if (payload.TryGetProperty("sampleRows", out var sr) && sr.TryGetInt32(out var sv)) sampleRows = Math.Clamp(sv, 1000, 500000); } catch { }
            try { if (payload.TryGetProperty("maxValues", out var mv) && mv.TryGetInt32(out var vv)) maxValues = Math.Clamp(vv, 20, 2000); } catch { }

            Task.Run(() =>
            {
                try
                {
                    var started = DateTime.Now;
                    var outDir = GetEnumArchiveDir(_activeSchemeId!);
                    var safeTable = new string(table.Where(ch => char.IsLetterOrDigit(ch) || ch == '_' || ch == '-').ToArray());
                    if (string.IsNullOrWhiteSpace(safeTable)) safeTable = "table";
                    var outPath = Path.Combine(outDir, $"enum_{safeTable}_{DateTime.Now:yyyyMMdd-HHmmss}.json");

                    var result = new Dictionary<string, object?>();
                    result["schemaVersion"] = 1;
                    result["schemeId"] = _activeSchemeId;
                    result["table"] = table;
                    result["sampleRows"] = sampleRows;
                    result["maxValues"] = maxValues;
                    result["withFreq"] = withFreq;
                    result["createdAt"] = DateTime.Now.ToString("s");

                    var colResults = new List<object>();
                    for (int idx = 0; idx < cols.Count; idx++)
                    {
                        token.ThrowIfCancellationRequested();
                        var col = cols[idx];

                        SendMessageToWebView(new { action = "dsEnumArchiveProgress", requestId, ok = true, data = new { phase = "extract", table, column = col, index = idx + 1, total = cols.Count } });

                        // 用 LIMIT 做顺序采样（避免 ORDER BY RANDOM()）
                        // 频次模式：在采样子集上 group by topK
                        var values = new List<object>();
                        bool truncated = false;
                        using (var cmd = _sqliteManager.Connection.CreateCommand())
                        {
                            if (withFreq)
                            {
                                cmd.CommandText = $@"
SELECT v, cnt FROM (
  SELECT {SqliteManager.QuoteIdent(col)} AS v, COUNT(*) AS cnt
  FROM (SELECT {SqliteManager.QuoteIdent(col)} FROM {SqliteManager.QuoteIdent(table)} LIMIT {sampleRows})
  WHERE v IS NOT NULL AND TRIM(CAST(v AS TEXT)) <> ''
  GROUP BY v
  ORDER BY cnt DESC
  LIMIT {maxValues + 1}
);";
                                using var rd = cmd.ExecuteReader();
                                while (rd.Read())
                                {
                                    token.ThrowIfCancellationRequested();
                                    var v = rd.IsDBNull(0) ? "" : Convert.ToString(rd.GetValue(0)) ?? "";
                                    var cnt = rd.IsDBNull(1) ? 0 : Convert.ToInt64(rd.GetValue(1));
                                    values.Add(new { value = v, count = cnt });
                                    if (values.Count > maxValues) { truncated = true; break; }
                                }
                            }
                            else
                            {
                                cmd.CommandText = $@"
SELECT v FROM (
  SELECT DISTINCT CAST({SqliteManager.QuoteIdent(col)} AS TEXT) AS v
  FROM (SELECT {SqliteManager.QuoteIdent(col)} FROM {SqliteManager.QuoteIdent(table)} LIMIT {sampleRows})
  WHERE v IS NOT NULL AND TRIM(v) <> ''
  LIMIT {maxValues + 1}
);";
                                using var rd = cmd.ExecuteReader();
                                while (rd.Read())
                                {
                                    token.ThrowIfCancellationRequested();
                                    var v = rd.IsDBNull(0) ? "" : (rd.GetString(0) ?? "");
                                    values.Add(v);
                                    if (values.Count > maxValues) { truncated = true; break; }
                                }
                            }
                        }
                        if (truncated && values.Count > maxValues) values = values.Take(maxValues).ToList();

                        colResults.Add(new
                        {
                            column = col,
                            valueCount = values.Count,
                            truncated,
                            values
                        });
                    }

                    result["columns"] = colResults;
                    result["elapsedMs"] = (long)(DateTime.Now - started).TotalMilliseconds;

                    var json = System.Text.Json.JsonSerializer.Serialize(result, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText(outPath, json, Encoding.UTF8);

                    SendMessageToWebView(new { action = "dsEnumArchiveCompleted", requestId, ok = true, data = new { fileName = Path.GetFileName(outPath), path = outPath.Replace('\\', '/'), columnCount = cols.Count } });
                }
                catch (OperationCanceledException)
                {
                    SendMessageToWebView(new { action = "dsEnumArchiveCompleted", requestId, ok = false, message = "已取消", errorCode = "E_CANCELED" });
                }
                catch (Exception ex)
                {
                    SendMessageToWebView(new { action = "dsEnumArchiveCompleted", requestId, ok = false, message = ex.Message, errorCode = "E_SQLITE_FAIL" });
                }
            }, token);
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "dsEnumArchiveCompleted", requestId, ok = false, message = ex.Message, errorCode = "E_SQLITE_FAIL" });
        }
    }

    private void DsListTemplates(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            var cfg = EnsureDesensitizationConfigForActiveScheme();
            EnsurePolicyRepoReady(cfg);
            var ns = cfg.Namespace ?? "DEFAULT";
            // 注意：namespace 是 C# 关键字。这里 SQL 别名成 ns，避免 dynamic 访问冲突（r.namespace 会报语法错误）
            var rows = _policyConn!.Query("SELECT id, namespace AS ns, name, description FROM template WHERE namespace = @ns OR namespace = 'DEFAULT' ORDER BY name;", new { ns })
                .Select(r => new { id = (string)r.id, @namespace = (string)r.ns, name = (string)r.name, description = (string?)r.description })
                .ToList();
            SendMessageToWebView(new { action = "dsTemplateListLoaded", requestId, ok = true, data = new { templates = rows } });
        }
        catch (Exception ex)
        {
            WriteErrorLog("DsListTemplates失败", ex);
            SendMessageToWebView(new { action = "dsTemplateListLoaded", requestId, ok = false, message = ex.Message, errorCode = "E_SQLITE_FAIL" });
        }
    }

    private void DsGetTemplate(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            var payload = data.TryGetProperty("payload", out var pl) ? pl : data;
            var templateId = (payload.TryGetProperty("templateId", out var tid) ? tid.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(templateId)) throw new InvalidOperationException("templateId 不能为空");

            var cfg = EnsureDesensitizationConfigForActiveScheme();
            EnsurePolicyRepoReady(cfg);

            var tpl = _policyConn!.QueryFirstOrDefault("SELECT id, namespace AS ns, name, description FROM template WHERE id=@id LIMIT 1;", new { id = templateId });
            if (tpl == null) throw new InvalidOperationException("模板不存在");
            var rules = _policyConn.Query(@"
SELECT id, template_id, table_name, column_name, match_mode, match_pattern, data_type, action,
       output_token_col, output_mask_col, keep_raw_col,
       normalize_profile, on_error, enabled, sort_order
FROM template_rule WHERE template_id=@id ORDER BY sort_order, column_name;", new { id = templateId })
                .Select(r => new
                {
                    id = (string)r.id,
                    templateId = (string)r.template_id,
                    tableName = (string?)r.table_name,
                    columnName = (string)r.column_name,
                    matchMode = (string?)r.match_mode ?? "exact",
                    matchPattern = (string?)r.match_pattern,
                    dataType = (string)r.data_type,
                    action = (string)r.action,
                    outputTokenCol = (string?)r.output_token_col,
                    outputMaskCol = (string?)r.output_mask_col,
                    keepRawCol = Convert.ToInt32(r.keep_raw_col) != 0,
                    normalizeProfile = (string)r.normalize_profile,
                    onError = (string)r.on_error,
                    enabled = Convert.ToInt32(r.enabled) != 0,
                    sortOrder = Convert.ToInt32(r.sort_order)
                })
                .ToList();

            SendMessageToWebView(new
            {
                action = "dsTemplateLoaded",
                requestId,
                ok = true,
                data = new
                {
                    template = new { id = (string)tpl.id, @namespace = (string)tpl.ns, name = (string)tpl.name, description = (string?)tpl.description },
                    rules
                }
            });
        }
        catch (Exception ex)
        {
            WriteErrorLog("DsGetTemplate失败", ex);
            SendMessageToWebView(new { action = "dsTemplateLoaded", requestId, ok = false, message = ex.Message, errorCode = "E_SQLITE_FAIL" });
        }
    }

    private void DsUpsertTemplate(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            var cfg = EnsureDesensitizationConfigForActiveScheme();
            EnsurePolicyRepoReady(cfg);

            var payload = data.TryGetProperty("payload", out var pl) ? pl : data;
            if (!payload.TryGetProperty("template", out var tplEl) || tplEl.ValueKind != System.Text.Json.JsonValueKind.Object)
                throw new InvalidOperationException("payload.template 缺失");

            string tplId = (tplEl.TryGetProperty("id", out var idEl) ? idEl.GetString() : null) ?? "";
            var name = (tplEl.TryGetProperty("name", out var nEl) ? nEl.GetString() : null) ?? "";
            var desc = (tplEl.TryGetProperty("description", out var dEl) ? dEl.GetString() : null) ?? "";
            var ns = (tplEl.TryGetProperty("namespace", out var nsEl) ? nsEl.GetString() : null) ?? cfg.Namespace ?? "DEFAULT";
            if (string.IsNullOrWhiteSpace(name)) throw new InvalidOperationException("模板名称不能为空");
            if (string.IsNullOrWhiteSpace(tplId)) tplId = "tpl_" + Guid.NewGuid().ToString("N");

            var rules = new List<System.Text.Json.JsonElement>();
            if (payload.TryGetProperty("rules", out var rulesEl) && rulesEl.ValueKind == System.Text.Json.JsonValueKind.Array)
                rules = rulesEl.EnumerateArray().Where(x => x.ValueKind == System.Text.Json.JsonValueKind.Object).ToList();

            using var tx = _policyConn!.BeginTransaction();
            _policyConn.Execute("INSERT OR REPLACE INTO template(id, namespace, name, description, created_at, created_by, updated_at, updated_by) VALUES(@id,@ns,@name,@desc,COALESCE((SELECT created_at FROM template WHERE id=@id),datetime('now')),'user',datetime('now'),'user');",
                new { id = tplId, ns, name, desc }, tx);
            _policyConn.Execute("DELETE FROM template_rule WHERE template_id=@id;", new { id = tplId }, tx);

            int sort = 10;
            foreach (var r in rules)
            {
                var col = (r.TryGetProperty("columnName", out var cEl) ? cEl.GetString() : null) ?? "";
                if (string.IsNullOrWhiteSpace(col)) continue;
                var matchMode = (r.TryGetProperty("matchMode", out var mmEl) ? mmEl.GetString() : null) ?? "exact";
                var matchPattern = (r.TryGetProperty("matchPattern", out var mpEl) ? mpEl.GetString() : null) ?? "";
                var dataType = (r.TryGetProperty("dataType", out var tEl) ? tEl.GetString() : null) ?? "PHONE";
                var action = (r.TryGetProperty("action", out var aEl) ? aEl.GetString() : null) ?? "TOKENIZE";
                var normalizeProfile = (r.TryGetProperty("normalizeProfile", out var npEl) ? npEl.GetString() : null) ?? "default";
                var onError = (r.TryGetProperty("onError", out var oeEl) ? oeEl.GetString() : null) ?? "fail";
                var keepRaw = (r.TryGetProperty("keepRawCol", out var krEl) && krEl.ValueKind == System.Text.Json.JsonValueKind.True) ? 1 : 0;
                var enabled = !(r.TryGetProperty("enabled", out var enEl) && enEl.ValueKind == System.Text.Json.JsonValueKind.False) ? 1 : 0;
                var sortOrder = (r.TryGetProperty("sortOrder", out var soEl) && soEl.TryGetInt32(out var soV)) ? soV : sort;
                sort += 10;

                _policyConn.Execute(@"
INSERT INTO template_rule(
  id, template_id, table_name, column_name, match_mode, match_pattern, data_type, action,
  output_token_col, output_mask_col, keep_raw_col,
  normalize_profile, normalize_params, on_error, enabled, sort_order
) VALUES(
  @id,@tpl,NULL,@col,@mm,@mp,@dt,@act,NULL,NULL,@keepRaw,@np,NULL,@oe,@en,@so
);",
                    new
                    {
                        id = "tplr_" + Guid.NewGuid().ToString("N"),
                        tpl = tplId,
                        col,
                        mm = matchMode,
                        mp = string.IsNullOrWhiteSpace(matchPattern) ? null : matchPattern,
                        dt = dataType,
                        act = action,
                        keepRaw,
                        np = normalizeProfile,
                        oe = onError,
                        en = enabled,
                        so = sortOrder
                    }, tx);
            }

            tx.Commit();
            SendMessageToWebView(new { action = "dsTemplateSaved", requestId, ok = true, data = new { templateId = tplId } });
        }
        catch (Exception ex)
        {
            WriteErrorLog("DsUpsertTemplate失败", ex);
            SendMessageToWebView(new { action = "dsTemplateSaved", requestId, ok = false, message = ex.Message, errorCode = "E_SQLITE_FAIL" });
        }
    }

    private void DsDeleteTemplate(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            var cfg = EnsureDesensitizationConfigForActiveScheme();
            EnsurePolicyRepoReady(cfg);
            var payload = data.TryGetProperty("payload", out var pl) ? pl : data;
            var templateId = (payload.TryGetProperty("templateId", out var tid) ? tid.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(templateId)) throw new InvalidOperationException("templateId 不能为空");
            using var tx = _policyConn!.BeginTransaction();
            _policyConn.Execute("DELETE FROM template_rule WHERE template_id=@id;", new { id = templateId }, tx);
            _policyConn.Execute("DELETE FROM template WHERE id=@id;", new { id = templateId }, tx);
            tx.Commit();
            SendMessageToWebView(new { action = "dsTemplateDeleted", requestId, ok = true });
        }
        catch (Exception ex)
        {
            WriteErrorLog("DsDeleteTemplate失败", ex);
            SendMessageToWebView(new { action = "dsTemplateDeleted", requestId, ok = false, message = ex.Message, errorCode = "E_SQLITE_FAIL" });
        }
    }

    private void DsGetTableColumns(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            if (_sqliteManager?.Connection == null) throw new InvalidOperationException("SQLite未就绪");
            var payload = data.TryGetProperty("payload", out var pl) ? pl : data;
            var table = (payload.TryGetProperty("table", out var tn) ? tn.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(table)) throw new InvalidOperationException("table 不能为空");

            var cols = new List<string>();
            using var cmd = _sqliteManager.Connection.CreateCommand();
            cmd.CommandText = $"PRAGMA table_info({SqliteManager.QuoteIdent(table)});";
            using var rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                var name = rd.IsDBNull(1) ? "" : rd.GetString(1);
                if (!string.IsNullOrWhiteSpace(name)) cols.Add(name);
            }
            SendMessageToWebView(new { action = "dsTableColumnsLoaded", requestId, ok = true, data = new { table, columns = cols } });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "dsTableColumnsLoaded", requestId, ok = false, message = ex.Message, errorCode = "E_SQLITE_FAIL" });
        }
    }

    private void DsListAuditLogs(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            var cfg = EnsureDesensitizationConfigForActiveScheme();
            EnsureVaultReady(cfg);
            var payload = data.TryGetProperty("payload", out var pl) ? pl : data;
            int limit = 200;
            try { if (payload.TryGetProperty("limit", out var li) && li.TryGetInt32(out var lv)) limit = Math.Clamp(lv, 20, 1000); } catch { }

            var rows = _vaultConn!.Query(@"
SELECT id, action, namespace, operator, role, reason_ticket, input_ref, output_ref,
       policy_id, policy_version, row_count, col_count, started_at, finished_at,
       status, error_message, detail_json, created_at
FROM audit_log
WHERE namespace = @ns OR @ns = '' OR @ns IS NULL
ORDER BY created_at DESC
LIMIT @limit;", new { ns = cfg.Namespace ?? "", limit })
                .Select(r => new
                {
                    id = (string)r.id,
                    action = (string)r.action,
                    @namespace = (string?)r.@namespace,
                    @operator = (string?)r.@operator,
                    role = (string?)r.role,
                    reasonTicket = (string?)r.reason_ticket,
                    inputRef = (string?)r.input_ref,
                    outputRef = (string?)r.output_ref,
                    policyId = (string?)r.policy_id,
                    policyVersion = r.policy_version == null ? (int?)null : Convert.ToInt32(r.policy_version),
                    rowCount = r.row_count == null ? (int?)null : Convert.ToInt32(r.row_count),
                    colCount = r.col_count == null ? (int?)null : Convert.ToInt32(r.col_count),
                    startedAt = (string?)r.started_at,
                    finishedAt = (string?)r.finished_at,
                    status = (string?)r.status,
                    errorMessage = (string?)r.error_message,
                    detailJson = (string?)r.detail_json,
                    createdAt = (string?)r.created_at
                })
                .ToList();

            SendMessageToWebView(new { action = "dsAuditLogsLoaded", requestId, ok = true, data = new { logs = rows } });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "dsAuditLogsLoaded", requestId, ok = false, message = ex.Message, errorCode = "E_SQLITE_FAIL" });
        }
    }

    // ==================== OutputMode 路由（Raw/Masked） ====================

    private void SetOutputMode(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            var payload = data.TryGetProperty("payload", out var pl) ? pl : data;
            var mode = (payload.TryGetProperty("mode", out var md) ? md.GetString() : null) ?? "raw";
            mode = string.Equals(mode, "masked", StringComparison.OrdinalIgnoreCase) ? "masked" : "raw";
            _outputMode = mode;
            // 持久化到项目 meta（下次打开项目自动恢复）
            try
            {
                if (!string.IsNullOrWhiteSpace(_activeSchemeId))
                {
                    var meta = LoadOrCreateProjectMeta(_activeSchemeId!, displayName: null, settingsJson: null);
                    if (meta.Desensitization == null) meta.Desensitization = new DesensitizationConfigV1();
                    meta.Desensitization.OutputMode = _outputMode;
                    SaveProjectMeta(meta);
                }
            }
            catch { }
            ApplyOutputModeRouting();
            SendMessageToWebView(new { action = "outputModeSet", requestId, ok = true, data = new { mode = _outputMode } });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "outputModeSet", requestId, ok = false, message = ex.Message, errorCode = "E_SQLITE_FAIL" });
        }
    }

    private void ApplyOutputModeRouting()
    {
        if (_sqliteManager?.Connection == null) return;
        if (_outputMode == "masked")
        {
            CreateMaskedShadowViews();
        }
        else
        {
            DropMaskedShadowViews();
        }
    }

    private void CreateMaskedShadowViews()
    {
        if (_sqliteManager?.Connection == null) return;

        // 规则：若存在同名 *_masked 表，则创建 TEMP VIEW <rawName> 指向 <rawName>_masked
        // TEMP schema 优先级高于 MAIN，可实现“无感路由”。
        var tables = _sqliteManager.GetTables() ?? new List<string>();
        var set = new HashSet<string>(tables, StringComparer.OrdinalIgnoreCase);
        foreach (var t in tables)
        {
            if (t.EndsWith("_masked", StringComparison.OrdinalIgnoreCase)) continue;
            var masked = t + "_masked";
            if (!set.Contains(masked)) continue;

            try
            {
                using var cmd = _sqliteManager.Connection.CreateCommand();
                cmd.CommandText = $"CREATE TEMP VIEW IF NOT EXISTS {SqliteManager.QuoteIdent(t)} AS SELECT * FROM main.{SqliteManager.QuoteIdent(masked)};";
                cmd.ExecuteNonQuery();
                _maskedShadowViews.Add(t);
            }
            catch { }
        }
    }

    private void DropMaskedShadowViews()
    {
        if (_sqliteManager?.Connection == null) return;
        foreach (var t in _maskedShadowViews.ToList())
        {
            try
            {
                using var cmd = _sqliteManager.Connection.CreateCommand();
                cmd.CommandText = $"DROP VIEW IF EXISTS temp.{SqliteManager.QuoteIdent(t)};";
                cmd.ExecuteNonQuery();
            }
            catch { }
        }
        _maskedShadowViews.Clear();
    }

    private void DsGetRoutingStatus(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            if (_sqliteManager?.Connection == null) throw new InvalidOperationException("SQLite未就绪");
            var tables = _sqliteManager.GetTables() ?? new List<string>();
            var set = new HashSet<string>(tables, StringComparer.OrdinalIgnoreCase);
            var pairs = new List<object>();
            foreach (var t in tables)
            {
                if (t.EndsWith("_masked", StringComparison.OrdinalIgnoreCase)) continue;
                var masked = t + "_masked";
                if (!set.Contains(masked)) continue;
                pairs.Add(new { raw = t, masked });
            }
            var shadow = _maskedShadowViews.OrderBy(x => x).ToList();
            SendMessageToWebView(new
            {
                action = "dsRoutingStatusLoaded",
                requestId,
                ok = true,
                data = new
                {
                    mode = _outputMode,
                    shadowViews = shadow,
                    maskedPairs = pairs
                }
            });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "dsRoutingStatusLoaded", requestId, ok = false, message = ex.Message, errorCode = "E_SQLITE_FAIL" });
        }
    }

    // ==================== 脱敏审查（非AI版） ====================

    private static readonly string[] ReviewHighRiskKeywords = new[]
    {
        "税号","统一社会信用","纳税人识别","组织机构","证件","身份证","证照","执照","许可证","资质","法人","法定代表人",
        "银行","账号","开户","户名","发票","流水","商户","合同","协议","签章","公章","签署",
        "手机号","电话","联系人","地址","收款","付款"
    };

    private static string GuessRiskByColumnName(string col)
    {
        var s = col ?? "";
        foreach (var k in ReviewHighRiskKeywords)
            if (s.Contains(k, StringComparison.OrdinalIgnoreCase)) return "high";
        return "normal";
    }

    private void DsReviewCoverage(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            if (_sqliteManager?.Connection == null) throw new InvalidOperationException("SQLite未就绪");
            var cfg = EnsureDesensitizationConfigForActiveScheme();
            EnsurePolicyRepoReady(cfg);

            var payload = data.TryGetProperty("payload", out var pl) ? pl : data;
            var templateId = (payload.TryGetProperty("templateId", out var tid) ? tid.GetString() : null) ?? "tpl_pii_basic_tokenize_v1";
            var table = (payload.TryGetProperty("table", out var tb) ? tb.GetString() : null) ?? MainTableNameOrDefault();

            // 取规则（enabled=1）
            var rules = _policyConn!.Query(@"
SELECT column_name, match_mode, match_pattern, data_type, action
FROM template_rule WHERE template_id=@id AND enabled=1 ORDER BY sort_order;", new { id = templateId })
                .Select(r => new
                {
                    ColumnName = (string)r.column_name,
                    MatchMode = (string?)r.match_mode ?? "exact",
                    MatchPattern = (string?)r.match_pattern,
                    DataType = (string)r.data_type,
                    Action = (string)r.action
                })
                .ToList();

            bool Match(string col, dynamic rr)
            {
                string mode = ((string)rr.MatchMode ?? "exact").Trim().ToLowerInvariant();
                string pat = string.IsNullOrWhiteSpace((string?)rr.MatchPattern) ? (string)rr.ColumnName : (string)rr.MatchPattern;
                if (mode == "contains") return col.Contains(pat, StringComparison.OrdinalIgnoreCase);
                if (mode == "regex")
                {
                    try { return System.Text.RegularExpressions.Regex.IsMatch(col, pat, System.Text.RegularExpressions.RegexOptions.IgnoreCase); }
                    catch { return false; }
                }
                return string.Equals(col, (string)rr.ColumnName, StringComparison.OrdinalIgnoreCase);
            }

            var schema = _sqliteManager.GetTableSchema(table);
            if (schema == null || schema.Count == 0) throw new InvalidOperationException("表不存在或无字段");
            var result = new List<object>();
            foreach (var c in schema)
            {
                var col = c.ColumnName;
                var hit = rules.FirstOrDefault(r => Match(col, r));
                result.Add(new
                {
                    column = col,
                    matched = hit != null,
                    risk = GuessRiskByColumnName(col),
                    dataType = hit?.DataType,
                    action = hit?.Action
                });
            }
            SendMessageToWebView(new { action = "dsReviewCoverageResult", requestId, ok = true, data = new { table, templateId, columns = result } });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "dsReviewCoverageResult", requestId, ok = false, message = ex.Message, errorCode = "E_SQLITE_FAIL" });
        }
    }

    private void DsReviewStrength(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            var cfg = EnsureDesensitizationConfigForActiveScheme();
            EnsurePolicyRepoReady(cfg);
            var payload = data.TryGetProperty("payload", out var pl) ? pl : data;
            var templateId = (payload.TryGetProperty("templateId", out var tid) ? tid.GetString() : null) ?? "tpl_pii_basic_tokenize_v1";

            var highRiskTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "TAX_ID","IDNO","BANK_ACCOUNT","LICENSE_NO","CERT_NO","INVOICE_NO","INVOICE_TAX_ID",
                "TRANSACTION_ID","PAYMENT_CHANNEL","CONTRACT_NO","SEAL_INFO"
            };

            var rows = _policyConn!.Query(@"
SELECT column_name, match_mode, match_pattern, data_type, action, enabled
FROM template_rule WHERE template_id=@id ORDER BY sort_order;", new { id = templateId })
                .Select(r => new
                {
                    columnName = (string)r.column_name,
                    matchMode = (string?)r.match_mode ?? "exact",
                    matchPattern = (string?)r.match_pattern,
                    dataType = (string)r.data_type,
                    action = (string)r.action,
                    enabled = Convert.ToInt32(r.enabled) != 0
                })
                .ToList();

            var issues = new List<object>();
            foreach (var r in rows.Where(x => x.enabled))
            {
                if (highRiskTypes.Contains(r.dataType) && string.Equals(r.action, "PASS", StringComparison.OrdinalIgnoreCase))
                {
                    issues.Add(new { level = "high", message = "高风险类型不建议 PASS", rule = r });
                }
                if (string.Equals(r.dataType, "BANK_ACCOUNT", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(r.action, "TOKENIZE", StringComparison.OrdinalIgnoreCase))
                {
                    issues.Add(new { level = "high", message = "银行账号建议 TOKENIZE", rule = r });
                }
            }

            SendMessageToWebView(new { action = "dsReviewStrengthResult", requestId, ok = true, data = new { templateId, issues, ruleCount = rows.Count } });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "dsReviewStrengthResult", requestId, ok = false, message = ex.Message, errorCode = "E_SQLITE_FAIL" });
        }
    }

    private void DsReviewSampleScan(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            if (_sqliteManager?.Connection == null) throw new InvalidOperationException("SQLite未就绪");
            var payload = data.TryGetProperty("payload", out var pl) ? pl : data;
            var table = (payload.TryGetProperty("table", out var tb) ? tb.GetString() : null) ?? MainTableNameOrDefault();
            int limit = 100;
            try { if (payload.TryGetProperty("limit", out var li) && li.TryGetInt32(out var lv)) limit = Math.Clamp(lv, 20, 1000); } catch { }

            // 默认扫描 table_masked（如不存在则扫描 table）
            var scanTable = table;
            try
            {
                var ts = _sqliteManager.GetTables();
                if (ts.Contains(table + "_masked", StringComparer.OrdinalIgnoreCase)) scanTable = table + "_masked";
            }
            catch { }

            var patterns = new Dictionary<string, System.Text.RegularExpressions.Regex>
            {
                ["PHONE_11"] = new System.Text.RegularExpressions.Regex(@"\\b\\d{11}\\b"),
                ["TAX_ID_18"] = new System.Text.RegularExpressions.Regex(@"\\b[0-9A-Z]{18}\\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase),
                ["BANK_16_19"] = new System.Text.RegularExpressions.Regex(@"\\b\\d{16,19}\\b"),
            };

            var hits = patterns.Keys.ToDictionary(k => k, _ => 0);
            using var cmd = _sqliteManager.Connection.CreateCommand();
            cmd.CommandText = $"SELECT * FROM {SqliteManager.QuoteIdent(scanTable)} LIMIT {limit};";
            using var rd = cmd.ExecuteReader();
            int rows = 0;
            while (rd.Read())
            {
                rows++;
                for (int i = 0; i < rd.FieldCount; i++)
                {
                    if (rd.IsDBNull(i)) continue;
                    var s = Convert.ToString(rd.GetValue(i)) ?? "";
                    if (s.Length == 0) continue;
                    foreach (var kv in patterns)
                    {
                        if (kv.Value.IsMatch(s)) hits[kv.Key] += 1;
                    }
                }
            }

            SendMessageToWebView(new { action = "dsReviewSampleResult", requestId, ok = true, data = new { table = scanTable, sampledRows = rows, hits } });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "dsReviewSampleResult", requestId, ok = false, message = ex.Message, errorCode = "E_SQLITE_FAIL" });
        }
    }

    // ==================== 备份管理（最小可用版） ====================

    private string GetBackupDir(string schemeId)
    {
        var db = GetSchemeDbPath(schemeId);
        var dir = Path.GetDirectoryName(db) ?? GetDbDir();
        var bk = Path.Combine(dir, "backups", schemeId);
        Directory.CreateDirectory(bk);
        return bk;
    }

    private void DsBackupCreate(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            if (string.IsNullOrWhiteSpace(_activeSchemeId)) throw new InvalidOperationException("未打开项目");
            var payload = data.TryGetProperty("payload", out var pl) ? pl : data;
            bool includeRawDb = !(payload.TryGetProperty("includeRawDb", out var ir) && ir.ValueKind == System.Text.Json.JsonValueKind.False);
            var cfg = EnsureDesensitizationConfigForActiveScheme();

            var meta = LoadOrCreateProjectMeta(_activeSchemeId!, displayName: null, settingsJson: null);
            var paths = new List<string>();
            if (includeRawDb && File.Exists(meta.DbPath)) paths.Add(meta.DbPath);
            if (File.Exists(cfg.VaultDbPath)) paths.Add(cfg.VaultDbPath);
            if (File.Exists(cfg.PolicyDbPath)) paths.Add(cfg.PolicyDbPath);
            if (File.Exists(cfg.MaskedDbPath)) paths.Add(cfg.MaskedDbPath);
            var pj = GetProjectMetaPath(_activeSchemeId!);
            if (File.Exists(pj)) paths.Add(pj);

            var outDir = GetBackupDir(_activeSchemeId!);
            var outPath = Path.Combine(outDir, $"backup-{DateTime.Now:yyyyMMdd-HHmmss}.zip");
            using (var zip = ZipFile.Open(outPath, ZipArchiveMode.Create))
            {
                foreach (var p in paths.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    try { zip.CreateEntryFromFile(p, Path.GetFileName(p), System.IO.Compression.CompressionLevel.Optimal); } catch { }
                }
            }
            SendMessageToWebView(new { action = "dsBackupCreated", requestId, ok = true, data = new { backupPath = outPath.Replace('\\', '/'), fileCount = paths.Count } });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "dsBackupCreated", requestId, ok = false, message = ex.Message, errorCode = "E_IO_FAIL" });
        }
    }

    private void DsBackupList(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            if (string.IsNullOrWhiteSpace(_activeSchemeId)) throw new InvalidOperationException("未打开项目");
            var dir = GetBackupDir(_activeSchemeId!);
            var list = Directory.GetFiles(dir, "*.zip")
                .Select(p => new FileInfo(p))
                .OrderByDescending(f => f.LastWriteTimeUtc)
                .Take(200)
                .Select(f => new { name = f.Name, path = f.FullName.Replace('\\', '/'), size = f.Length, time = f.LastWriteTime.ToString("s") })
                .ToList();
            SendMessageToWebView(new { action = "dsBackupListLoaded", requestId, ok = true, data = new { backups = list } });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "dsBackupListLoaded", requestId, ok = false, message = ex.Message, errorCode = "E_IO_FAIL" });
        }
    }

    private void DsBackupRestore(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            if (string.IsNullOrWhiteSpace(_activeSchemeId)) throw new InvalidOperationException("未打开项目");
            var payload = data.TryGetProperty("payload", out var pl) ? pl : data;
            var fileName = (payload.TryGetProperty("fileName", out var fn) ? fn.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(fileName)) throw new InvalidOperationException("fileName 不能为空");
            var dir = GetBackupDir(_activeSchemeId!);
            var zipPath = Path.Combine(dir, fileName);
            if (!File.Exists(zipPath)) throw new FileNotFoundException("备份文件不存在", zipPath);

            var cfg = EnsureDesensitizationConfigForActiveScheme();
            var meta = LoadOrCreateProjectMeta(_activeSchemeId!, displayName: null, settingsJson: null);

            // 先关闭连接，避免文件占用
            try { _vaultConn?.Close(); } catch { }
            try { _policyConn?.Close(); } catch { }
            try { _vaultConn?.Dispose(); } catch { }
            try { _policyConn?.Dispose(); } catch { }
            _vaultConn = null; _policyConn = null;

            using (var zip = ZipFile.OpenRead(zipPath))
            {
                foreach (var e in zip.Entries)
                {
                    var name = e.Name;
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    string? outFile = null;
                    if (string.Equals(name, Path.GetFileName(meta.DbPath), StringComparison.OrdinalIgnoreCase)) outFile = meta.DbPath;
                    else if (string.Equals(name, Path.GetFileName(cfg.VaultDbPath), StringComparison.OrdinalIgnoreCase)) outFile = cfg.VaultDbPath;
                    else if (string.Equals(name, Path.GetFileName(cfg.PolicyDbPath), StringComparison.OrdinalIgnoreCase)) outFile = cfg.PolicyDbPath;
                    else if (string.Equals(name, Path.GetFileName(cfg.MaskedDbPath), StringComparison.OrdinalIgnoreCase)) outFile = cfg.MaskedDbPath;
                    else if (string.Equals(name, Path.GetFileName(GetProjectMetaPath(_activeSchemeId!)), StringComparison.OrdinalIgnoreCase)) outFile = GetProjectMetaPath(_activeSchemeId!);
                    if (outFile == null) continue;
                    Directory.CreateDirectory(Path.GetDirectoryName(outFile) ?? dir);
                    e.ExtractToFile(outFile, overwrite: true);
                }
            }

            // 重新打开 Vault/Policy（masked/业务库由用户后续刷新/重新导入决定）
            try { EnsureVaultReady(cfg); } catch { }
            try { EnsurePolicyRepoReady(cfg); } catch { }

            SendMessageToWebView(new { action = "dsBackupRestored", requestId, ok = true, data = new { fileName } });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "dsBackupRestored", requestId, ok = false, message = ex.Message, errorCode = "E_IO_FAIL" });
        }
    }

    private void DsBackupDelete(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            if (string.IsNullOrWhiteSpace(_activeSchemeId)) throw new InvalidOperationException("未打开项目");
            var payload = data.TryGetProperty("payload", out var pl) ? pl : data;
            var fileName = (payload.TryGetProperty("fileName", out var fn) ? fn.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(fileName)) throw new InvalidOperationException("fileName 不能为空");
            var dir = GetBackupDir(_activeSchemeId!);
            var zipPath = Path.Combine(dir, fileName);
            if (!File.Exists(zipPath)) throw new FileNotFoundException("备份文件不存在", zipPath);
            File.Delete(zipPath);
            SendMessageToWebView(new { action = "dsBackupDeleted", requestId, ok = true, data = new { fileName } });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "dsBackupDeleted", requestId, ok = false, message = ex.Message, errorCode = "E_IO_FAIL" });
        }
    }

    private void DsBackupOpenFolder(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            if (string.IsNullOrWhiteSpace(_activeSchemeId)) throw new InvalidOperationException("未打开项目");
            var dir = GetBackupDir(_activeSchemeId!);
            try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(dir) { UseShellExecute = true }); } catch { }
            SendMessageToWebView(new { action = "dsBackupFolderOpened", requestId, ok = true, data = new { dir = dir.Replace('\\', '/') } });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "dsBackupFolderOpened", requestId, ok = false, message = ex.Message, errorCode = "E_IO_FAIL" });
        }
    }

    private void GetDesensitizationStatus(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            var cfg = EnsureDesensitizationConfigForActiveScheme();
            var maskedTables = new List<string>();
            try
            {
                if (_sqliteManager != null)
                {
                    var ts = _sqliteManager.GetTables();
                    maskedTables = ts.Where(t => t.EndsWith("_masked", StringComparison.OrdinalIgnoreCase)).OrderBy(x => x).ToList();
                }
            }
            catch { }

            SendMessageToWebView(new
            {
                action = "desensitizationStatus",
                requestId,
                ok = true,
                data = new
                {
                    schemeId = _activeSchemeId,
                    @namespace = cfg.Namespace,
                    vaultDbPath = (cfg.VaultDbPath ?? "").Replace('\\', '/'),
                    policyDbPath = (cfg.PolicyDbPath ?? "").Replace('\\', '/'),
                    maskedDbPath = (cfg.MaskedDbPath ?? "").Replace('\\', '/'),
                    vaultReady = _vaultConn != null && _vaultConn.State == ConnectionState.Open,
                    policyReady = _policyConn != null && _policyConn.State == ConnectionState.Open,
                    maskedTables
                }
            });
        }
        catch (Exception ex)
        {
            SendMessageToWebView(new { action = "desensitizationStatus", requestId, ok = false, message = ex.Message, errorCode = "E_NOT_READY" });
        }
    }

    private sealed class TemplateRuleRow
    {
        public string ColumnName { get; set; } = "";
        public string MatchMode { get; set; } = "exact";   // exact/contains/regex
        public string? MatchPattern { get; set; }          // contains/regex 时可用；为空则用 ColumnName
        public string DataType { get; set; } = "";
        public string Action { get; set; } = "";
        public string NormalizeProfile { get; set; } = "default";
        public string OnError { get; set; } = "fail";
        public int KeepRawCol { get; set; } = 0;
        public string? OutputTokenCol { get; set; }
        public string? OutputMaskCol { get; set; }
    }

    private string GetOrCreateToken(DesensitizationConfigV1 cfg, string dataType, string raw, string policyId, int policyVersion, string createdBy)
    {
        EnsureVaultReady(cfg);
        if (_vaultHmacSecret == null) throw new InvalidOperationException("Vault secret 未就绪");

        var normalized = NormalizeValue(dataType, raw);
        if (string.IsNullOrWhiteSpace(normalized)) return "";

        var fp = HmacSha256Hex(_vaultHmacSecret, $"{cfg.Namespace}|{dataType}|{normalized}");
        // 查重
        using (var cmd = _vaultConn!.CreateCommand())
        {
            cmd.CommandText = "SELECT token FROM token_map WHERE namespace=$ns AND type=$t AND fingerprint=$fp LIMIT 1;";
            cmd.Parameters.AddWithValue("$ns", cfg.Namespace);
            cmd.Parameters.AddWithValue("$t", dataType);
            cmd.Parameters.AddWithValue("$fp", fp);
            var existed = cmd.ExecuteScalar() as string;
            if (!string.IsNullOrWhiteSpace(existed))
            {
                try
                {
                    using var upd = _vaultConn.CreateCommand();
                    upd.CommandText = "UPDATE token_map SET use_count=use_count+1, last_used_at=datetime('now') WHERE token=$tk;";
                    upd.Parameters.AddWithValue("$tk", existed);
                    upd.ExecuteNonQuery();
                }
                catch { }
                return existed!;
            }
        }

        // 新建（少量重试：防 token UNIQUE 冲突）
        for (int retry = 0; retry < 5; retry++)
        {
            var token = BuildReadableToken(dataType, cfg.Namespace);
            var enc = EncryptByDpapi(normalized); // 目前存 normalized；如需存 raw，可改这里
            try
            {
                using var ins = _vaultConn!.CreateCommand();
                ins.CommandText = @"
INSERT INTO token_map(id, namespace, type, fingerprint, token, enc_value, policy_id, policy_version, created_at, created_by, use_count, last_used_at)
VALUES($id,$ns,$t,$fp,$tk,$enc,$pid,$pv,datetime('now'),$by,1,datetime('now'));";
                ins.Parameters.AddWithValue("$id", Guid.NewGuid().ToString("N"));
                ins.Parameters.AddWithValue("$ns", cfg.Namespace);
                ins.Parameters.AddWithValue("$t", dataType);
                ins.Parameters.AddWithValue("$fp", fp);
                ins.Parameters.AddWithValue("$tk", token);
                ins.Parameters.Add("$enc", SqliteType.Blob).Value = enc;
                ins.Parameters.AddWithValue("$pid", policyId);
                ins.Parameters.AddWithValue("$pv", policyVersion);
                ins.Parameters.AddWithValue("$by", createdBy);
                ins.ExecuteNonQuery();
                return token;
            }
            catch (SqliteException)
            {
                // 可能是 UNIQUE(token) 或 UNIQUE(fingerprint)，重查 fingerprint
                using var cmd = _vaultConn!.CreateCommand();
                cmd.CommandText = "SELECT token FROM token_map WHERE namespace=$ns AND type=$t AND fingerprint=$fp LIMIT 1;";
                cmd.Parameters.AddWithValue("$ns", cfg.Namespace);
                cmd.Parameters.AddWithValue("$t", dataType);
                cmd.Parameters.AddWithValue("$fp", fp);
                var existed = cmd.ExecuteScalar() as string;
                if (!string.IsNullOrWhiteSpace(existed)) return existed!;
            }
        }
        throw new InvalidOperationException("生成 token 失败（重试次数耗尽）");
    }

    private void ExecuteMaskJob(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            if (_sqliteManager?.Connection == null) throw new InvalidOperationException("SQLite未就绪，请先在项目中心打开项目并导入数据");
            bool expertMode = !(data.TryGetProperty("expertMode", out var em) && em.ValueKind == System.Text.Json.JsonValueKind.False);
            if (!expertMode) throw new InvalidOperationException("当前未开启【专家模式】，不允许执行脱敏生成（涉及可逆还原与审计）");

            var cfg = EnsureDesensitizationConfigForActiveScheme();
            EnsureVaultReady(cfg);
            EnsurePolicyRepoReady(cfg);

            var payload = data.TryGetProperty("payload", out var pl) ? pl : data;
            var sourceTable = (payload.TryGetProperty("sourceTable", out var st) ? st.GetString() : null) ?? MainTableNameOrDefault();
            var templateId = (payload.TryGetProperty("templateId", out var tid) ? tid.GetString() : null) ?? "tpl_pii_basic_tokenize_v1";
            bool overwrite = payload.TryGetProperty("overwrite", out var ow) && ow.ValueKind == System.Text.Json.JsonValueKind.True;
            bool includeMasked = !(payload.TryGetProperty("includeMasked", out var im) && im.ValueKind == System.Text.Json.JsonValueKind.False);

            // 加载模板规则（支持：精确/包含/正则 匹配字段名）
            var rules = new List<TemplateRuleRow>();
            using (var cmd = _policyConn!.CreateCommand())
            {
                cmd.CommandText = @"
SELECT column_name, match_mode, match_pattern, data_type, action,
       normalize_profile, on_error,
       keep_raw_col, output_token_col, output_mask_col
FROM template_rule WHERE template_id=$id AND enabled=1 ORDER BY sort_order;";
                cmd.Parameters.AddWithValue("$id", templateId);
                using var rd = cmd.ExecuteReader();
                while (rd.Read())
                {
                    rules.Add(new TemplateRuleRow
                    {
                        ColumnName = rd.IsDBNull(0) ? "" : rd.GetString(0),
                        MatchMode = rd.IsDBNull(1) ? "exact" : rd.GetString(1),
                        MatchPattern = rd.IsDBNull(2) ? null : rd.GetString(2),
                        DataType = rd.IsDBNull(3) ? "" : rd.GetString(3),
                        Action = rd.IsDBNull(4) ? "" : rd.GetString(4),
                        NormalizeProfile = rd.IsDBNull(5) ? "default" : rd.GetString(5),
                        OnError = rd.IsDBNull(6) ? "fail" : rd.GetString(6),
                        KeepRawCol = rd.IsDBNull(7) ? 0 : rd.GetInt32(7),
                        OutputTokenCol = rd.IsDBNull(8) ? null : rd.GetString(8),
                        OutputMaskCol = rd.IsDBNull(9) ? null : rd.GetString(9),
                    });
                }
            }
            if (rules.Count == 0) throw new InvalidOperationException("模板规则为空：请先初始化 Policy Repo 并导入默认模板");

            var schema = _sqliteManager.GetTableSchema(sourceTable);
            if (schema == null || schema.Count == 0) throw new InvalidOperationException($"源表无字段：{sourceTable}");

            TemplateRuleRow? FindRuleForColumn(string colName)
            {
                foreach (var r in rules)
                {
                    if (string.IsNullOrWhiteSpace(r.ColumnName)) continue;
                    var mode = (r.MatchMode ?? "exact").Trim().ToLowerInvariant();
                    var pat = string.IsNullOrWhiteSpace(r.MatchPattern) ? r.ColumnName : r.MatchPattern!;
                    if (mode == "contains")
                    {
                        if (colName.Contains(pat, StringComparison.OrdinalIgnoreCase)) return r;
                        continue;
                    }
                    if (mode == "regex")
                    {
                        try
                        {
                            if (System.Text.RegularExpressions.Regex.IsMatch(colName, pat, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                                return r;
                        }
                        catch { }
                        continue;
                    }
                    // exact（默认）
                    if (string.Equals(colName, r.ColumnName, StringComparison.OrdinalIgnoreCase)) return r;
                }
                return null;
            }

            var matchedRuleMap = new Dictionary<string, TemplateRuleRow>(StringComparer.OrdinalIgnoreCase);
            foreach (var c in schema)
            {
                var rr = FindRuleForColumn(c.ColumnName);
                if (rr != null) matchedRuleMap[c.ColumnName] = rr;
            }

            var sensitiveCols = schema
                .Select(c => c.ColumnName)
                .Where(cn => matchedRuleMap.ContainsKey(cn))
                .ToList();
            if (sensitiveCols.Count == 0) throw new InvalidOperationException("未匹配到任何敏感字段：请确认字段名（例如 手机号/身份证号/姓名/地址）或扩展模板规则");

            string targetTable = string.IsNullOrWhiteSpace(cfg.MaskedTablePrefix)
                ? $"{sourceTable}_masked"
                : $"{cfg.MaskedTablePrefix}{sourceTable}";

            var startedAt = DateTime.Now;

            // 写审计（开始）
            var auditId = Guid.NewGuid().ToString("N");
            try
            {
                using var ac = _vaultConn!.CreateCommand();
                ac.CommandText = @"INSERT INTO audit_log(id,action,namespace,operator,role,reason_ticket,input_ref,output_ref,policy_id,policy_version,row_count,col_count,started_at,status,created_at)
VALUES($id,'MASK',$ns,$op,'expert',NULL,$in,$out,$pid,$pv,0,$cc,$st,'RUNNING',datetime('now'));";
                ac.Parameters.AddWithValue("$id", auditId);
                ac.Parameters.AddWithValue("$ns", cfg.Namespace);
                ac.Parameters.AddWithValue("$op", Environment.UserName ?? "user");
                ac.Parameters.AddWithValue("$in", $"{_activeSchemeId}:{sourceTable}");
                ac.Parameters.AddWithValue("$out", $"{_activeSchemeId}:{targetTable}");
                ac.Parameters.AddWithValue("$pid", templateId);
                ac.Parameters.AddWithValue("$pv", 1);
                ac.Parameters.AddWithValue("$cc", schema.Count);
                ac.Parameters.AddWithValue("$st", startedAt.ToString("s"));
                ac.ExecuteNonQuery();
            }
            catch { }

            // 创建目标表
            using (var txn = _sqliteManager.Connection.BeginTransaction())
            {
                if (overwrite)
                {
                    using var drop = _sqliteManager.Connection.CreateCommand();
                    drop.Transaction = txn;
                    drop.CommandText = $"DROP TABLE IF EXISTS {SqliteManager.QuoteIdent(targetTable)};";
                    drop.ExecuteNonQuery();
                }
                // 构造列定义
                var colDefs = new List<string>();
                foreach (var c in schema)
                {
                    var cn = c.ColumnName;
                    bool isSensitive = matchedRuleMap.ContainsKey(cn);
                    if (isSensitive && !cfg.KeepRawInMasked && matchedRuleMap[cn].KeepRawCol == 0) continue;
                    var dt = string.IsNullOrWhiteSpace(c.DataType) ? "TEXT" : c.DataType;
                    colDefs.Add($"{SqliteManager.QuoteIdent(cn)} {dt}");
                }
                foreach (var cn in sensitiveCols)
                {
                    var r0 = matchedRuleMap[cn];
                    var tokenCol = string.IsNullOrWhiteSpace(r0.OutputTokenCol) ? $"{cn}_token" : r0.OutputTokenCol!;
                    colDefs.Add($"{SqliteManager.QuoteIdent(tokenCol)} TEXT");
                    if (includeMasked)
                    {
                        var maskCol = string.IsNullOrWhiteSpace(r0.OutputMaskCol) ? $"{cn}_masked" : r0.OutputMaskCol!;
                        colDefs.Add($"{SqliteManager.QuoteIdent(maskCol)} TEXT");
                    }
                }
                colDefs.Add("[mask_job_id] TEXT");
                colDefs.Add("[policy_id] TEXT");
                colDefs.Add("[policy_version] INTEGER");
                colDefs.Add("[masked_at] TEXT");

                using var create = _sqliteManager.Connection.CreateCommand();
                create.Transaction = txn;
                create.CommandText = $"CREATE TABLE IF NOT EXISTS {SqliteManager.QuoteIdent(targetTable)} ({string.Join(", ", colDefs)});";
                create.ExecuteNonQuery();
                txn.Commit();
            }

            // 写入数据（行级处理）
            int total = _sqliteManager.GetRowCount(sourceTable);
            int processed = 0;
            var insertCols = new List<string>();
            var insertParams = new List<string>();
            var paramNames = new List<string>();

            foreach (var c in schema)
            {
                var cn = c.ColumnName;
                bool isSensitive = matchedRuleMap.ContainsKey(cn);
                if (isSensitive && !cfg.KeepRawInMasked && matchedRuleMap[cn].KeepRawCol == 0) continue;
                insertCols.Add(cn);
            }
            foreach (var cn in sensitiveCols)
            {
                var r0 = matchedRuleMap[cn];
                insertCols.Add(string.IsNullOrWhiteSpace(r0.OutputTokenCol) ? $"{cn}_token" : r0.OutputTokenCol!);
                if (includeMasked) insertCols.Add(string.IsNullOrWhiteSpace(r0.OutputMaskCol) ? $"{cn}_masked" : r0.OutputMaskCol!);
            }
            insertCols.AddRange(new[] { "mask_job_id", "policy_id", "policy_version", "masked_at" });

            for (int i = 0; i < insertCols.Count; i++)
            {
                var pn = $"$p{i}";
                insertParams.Add(pn);
                paramNames.Add(pn);
            }

            using var readCmd = _sqliteManager.Connection!.CreateCommand();
            readCmd.CommandText = $"SELECT * FROM {SqliteManager.QuoteIdent(sourceTable)};";
            using var reader = readCmd.ExecuteReader();

            using var writeTxn = _sqliteManager.Connection!.BeginTransaction();
            using var insCmd = _sqliteManager.Connection.CreateCommand();
            insCmd.Transaction = writeTxn;
            insCmd.CommandText = $"INSERT INTO {SqliteManager.QuoteIdent(targetTable)} ({string.Join(",", insertCols.Select(SqliteManager.QuoteIdent))}) VALUES ({string.Join(",", insertParams)});";
            foreach (var pn in paramNames) insCmd.Parameters.Add(new SqliteParameter(pn, DBNull.Value));

            var by = Environment.UserName ?? "user";
            var tokenCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var jobId = Guid.NewGuid().ToString("N");
            var maskedAt = DateTime.Now.ToString("s");

            while (reader.Read())
            {
                var values = new List<object?>();
                foreach (var c in schema)
                {
                    var cn = c.ColumnName;
                    bool isSensitive = matchedRuleMap.ContainsKey(cn);
                    if (isSensitive && !cfg.KeepRawInMasked && matchedRuleMap[cn].KeepRawCol == 0) continue;
                    var v = reader[cn];
                    values.Add(v == DBNull.Value ? "" : v);
                }

                foreach (var cn in sensitiveCols)
                {
                    var rawObj = reader[cn];
                    var raw = rawObj == DBNull.Value ? "" : Convert.ToString(rawObj) ?? "";
                    var rr = matchedRuleMap[cn];
                    var dt = rr.DataType ?? "";

                    string token = "";
                    string masked = "";
                    if (!string.IsNullOrWhiteSpace(raw))
                    {
                        var normalized = NormalizeValue(dt, raw);
                        var fp = (_vaultHmacSecret == null) ? "" : HmacSha256Hex(_vaultHmacSecret, $"{cfg.Namespace}|{dt}|{normalized}");
                        var cacheKey = $"{dt}|{fp}";
                        if (!string.IsNullOrWhiteSpace(fp) && tokenCache.TryGetValue(cacheKey, out var cached))
                        {
                            token = cached;
                        }
                        else
                        {
                            token = GetOrCreateToken(cfg, dt, raw, templateId, 1, by);
                            if (!string.IsNullOrWhiteSpace(fp)) tokenCache[cacheKey] = token;
                        }
                        masked = MaskValue(dt, raw);
                    }

                    values.Add(token);
                    if (includeMasked) values.Add(masked);
                }

                values.Add(jobId);
                values.Add(templateId);
                values.Add(1);
                values.Add(maskedAt);

                for (int i = 0; i < values.Count; i++)
                    insCmd.Parameters[i].Value = values[i] ?? "";
                insCmd.ExecuteNonQuery();

                processed++;
                if (processed % 2000 == 0)
                {
                    var pct = total <= 0 ? 0 : (int)Math.Round(processed * 100.0 / total);
                    SendMessageToWebView(new { action = "maskJobProgress", requestId, ok = true, data = new { percent = pct, processed, total } });
                }
            }
            writeTxn.Commit();

            // 更新审计（完成）
            try
            {
                using var uc = _vaultConn!.CreateCommand();
                uc.CommandText = "UPDATE audit_log SET finished_at=$ft, status='OK', row_count=$rc WHERE id=$id;";
                uc.Parameters.AddWithValue("$ft", DateTime.Now.ToString("s"));
                uc.Parameters.AddWithValue("$rc", processed);
                uc.Parameters.AddWithValue("$id", auditId);
                uc.ExecuteNonQuery();
            }
            catch { }

            // 同步脱敏表到 masked db（独立库）
            try { SyncMaskedTableToMaskedDb(cfg, targetTable); } catch { }

            try { GetTableList(); } catch { }

            SendMessageToWebView(new { action = "maskJobCompleted", requestId, ok = true, data = new { sourceTable, targetTable, rowCount = processed, auditLogId = auditId } });
        }
        catch (Exception ex)
        {
            WriteErrorLog("ExecuteMaskJob失败", ex);
            SendMessageToWebView(new { action = "maskJobCompleted", requestId, ok = false, message = ex.Message, errorCode = "E_SQLITE_FAIL" });
        }
    }

    private void SyncMaskedTableToMaskedDb(DesensitizationConfigV1 cfg, string targetTable)
    {
        if (_sqliteManager?.Connection == null) return;
        if (string.IsNullOrWhiteSpace(cfg.MaskedDbPath)) return;
        if (string.IsNullOrWhiteSpace(targetTable)) return;

        var maskedPath = cfg.MaskedDbPath;
        Directory.CreateDirectory(Path.GetDirectoryName(maskedPath) ?? AppContext.BaseDirectory);

        using var cmd = _sqliteManager.Connection.CreateCommand();
        // 注意：ATTACH/DETACH 需要单连接执行
        cmd.CommandText = "ATTACH DATABASE $p AS masked;";
        cmd.Parameters.AddWithValue("$p", maskedPath);
        cmd.ExecuteNonQuery();

        try
        {
            using var drop = _sqliteManager.Connection.CreateCommand();
            drop.CommandText = $"DROP TABLE IF EXISTS masked.{SqliteManager.QuoteIdent(targetTable)};";
            drop.ExecuteNonQuery();

            using var create = _sqliteManager.Connection.CreateCommand();
            create.CommandText = $"CREATE TABLE masked.{SqliteManager.QuoteIdent(targetTable)} AS SELECT * FROM main.{SqliteManager.QuoteIdent(targetTable)};";
            create.ExecuteNonQuery();
        }
        finally
        {
            try
            {
                using var detach = _sqliteManager.Connection.CreateCommand();
                detach.CommandText = "DETACH DATABASE masked;";
                detach.ExecuteNonQuery();
            }
            catch { }
        }
    }

    private void DetokenizeTableExport(System.Text.Json.JsonElement data)
    {
        var requestId = (data.TryGetProperty("requestId", out var rid) ? rid.GetString() : null) ?? "";
        try
        {
            if (_sqliteManager?.Connection == null) throw new InvalidOperationException("SQLite未就绪");
            bool expertMode = !(data.TryGetProperty("expertMode", out var em) && em.ValueKind == System.Text.Json.JsonValueKind.False);
            if (!expertMode) throw new InvalidOperationException("未开启专家模式：不允许还原明文");

            var payload = data.TryGetProperty("payload", out var pl) ? pl : data;
            var reason = (payload.TryGetProperty("reasonTicket", out var rt) ? rt.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(reason)) throw new InvalidOperationException("reasonTicket 必填（工单号/用途说明）");

            var cfg = EnsureDesensitizationConfigForActiveScheme();
            EnsureVaultReady(cfg);

            var table = (payload.TryGetProperty("table", out var tn) ? tn.GetString() : null) ?? "";
            if (string.IsNullOrWhiteSpace(table)) throw new InvalidOperationException("table 不能为空");
            var tokenCols = new List<string>();
            if (payload.TryGetProperty("tokenColumns", out var tc) && tc.ValueKind == System.Text.Json.JsonValueKind.Array)
                tokenCols = tc.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
            if (tokenCols.Count == 0) throw new InvalidOperationException("tokenColumns 不能为空");

            var mode = (payload.TryGetProperty("mode", out var md) ? md.GetString() : null) ?? "saveas"; // saveas/open
            bool openAfter = string.Equals(mode, "open", StringComparison.OrdinalIgnoreCase);
            string outPath;
            if (openAfter)
            {
                outPath = Path.Combine(Path.GetTempPath(), $"detokenized-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx");
            }
            else
            {
                using var sfd = new SaveFileDialog();
                sfd.Title = "导出明文报表（Detokenize）";
                sfd.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
                sfd.FileName = $"detokenized-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx";
                if (sfd.ShowDialog() != DialogResult.OK) return;
                outPath = sfd.FileName;
            }

            // 审计：开始
            var auditId = Guid.NewGuid().ToString("N");
            try
            {
                using var ac = _vaultConn!.CreateCommand();
                ac.CommandText = @"INSERT INTO audit_log(id,action,namespace,operator,role,reason_ticket,input_ref,output_ref,started_at,status,created_at)
VALUES($id,'DETOKENIZE',$ns,$op,'expert',$rs,$in,$out,$st,'RUNNING',datetime('now'));";
                ac.Parameters.AddWithValue("$id", auditId);
                ac.Parameters.AddWithValue("$ns", cfg.Namespace);
                ac.Parameters.AddWithValue("$op", Environment.UserName ?? "user");
                ac.Parameters.AddWithValue("$rs", reason);
                ac.Parameters.AddWithValue("$in", $"{_activeSchemeId}:{table}");
                ac.Parameters.AddWithValue("$out", outPath);
                ac.Parameters.AddWithValue("$st", DateTime.Now.ToString("s"));
                ac.ExecuteNonQuery();
            }
            catch { }

            using var cmd = _sqliteManager.Connection.CreateCommand();
            cmd.CommandText = $"SELECT * FROM {SqliteManager.QuoteIdent(table)};";
            using var rd = cmd.ExecuteReader();
            int colCount = rd.FieldCount;

            using var wb = new ClosedXML.Excel.XLWorkbook();
            var ws = wb.Worksheets.Add("Detokenized");
            ApplyExcelReportDefaults(ws);

            var headers = new List<string>();
            for (int i = 0; i < colCount; i++) headers.Add(rd.GetName(i));
            var realCols = tokenCols.Select(c => $"{c}_real").ToList();
            headers.AddRange(realCols);
            for (int c = 0; c < headers.Count; c++)
            {
                ws.Cell(1, c + 1).Value = headers[c];
                ws.Cell(1, c + 1).Style.Font.Bold = true;
            }

            int row = 2;
            var cache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            while (rd.Read())
            {
                for (int c = 0; c < colCount; c++)
                {
                    var v = rd.IsDBNull(c) ? "" : Convert.ToString(rd.GetValue(c)) ?? "";
                    ws.Cell(row, c + 1).SetValue(v);
                }

                for (int i = 0; i < tokenCols.Count; i++)
                {
                    var tcName = tokenCols[i];
                    string token = "";
                    try
                    {
                        var idx = rd.GetOrdinal(tcName);
                        token = rd.IsDBNull(idx) ? "" : Convert.ToString(rd.GetValue(idx)) ?? "";
                    }
                    catch { token = ""; }

                    string real = "";
                    if (!string.IsNullOrWhiteSpace(token))
                    {
                        if (!cache.TryGetValue(token, out real))
                        {
                            using var vc = _vaultConn!.CreateCommand();
                            vc.CommandText = "SELECT enc_value FROM token_map WHERE token=$t LIMIT 1;";
                            vc.Parameters.AddWithValue("$t", token);
                            var obj = vc.ExecuteScalar();
                            if (obj is byte[] b) real = DecryptByDpapi(b);
                            cache[token] = real;
                        }
                    }
                    ws.Cell(row, colCount + i + 1).SetValue(real);
                }

                row++;
                if (row % 2000 == 0)
                    SendMessageToWebView(new { action = "detokenizeProgress", requestId, ok = true, data = new { processed = row - 2 } });
            }

            try { ApplyTextLeftNumberRightAlignment(ws, headerRow: 1); } catch { }
            ws.Columns().AdjustToContents();
            wb.SaveAs(outPath);

            if (openAfter)
            {
                try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outPath) { UseShellExecute = true }); } catch { }
            }

            // 审计：完成
            try
            {
                using var uc = _vaultConn!.CreateCommand();
                uc.CommandText = "UPDATE audit_log SET finished_at=$ft, status='OK', row_count=$rc, output_ref=$out WHERE id=$id;";
                uc.Parameters.AddWithValue("$ft", DateTime.Now.ToString("s"));
                uc.Parameters.AddWithValue("$rc", Math.Max(0, row - 2));
                uc.Parameters.AddWithValue("$out", outPath);
                uc.Parameters.AddWithValue("$id", auditId);
                uc.ExecuteNonQuery();
            }
            catch { }

            SendMessageToWebView(new { action = "detokenizeExported", requestId, ok = true, data = new { outputPath = outPath.Replace('\\', '/'), auditLogId = auditId } });
        }
        catch (Exception ex)
        {
            WriteErrorLog("DetokenizeTableExport失败", ex);
            SendMessageToWebView(new { action = "detokenizeExported", requestId, ok = false, message = ex.Message, errorCode = "E_SQLITE_FAIL" });
        }
    }

    private void SendMessageToWebView(object message)
    {
        try
        {
            if (this.InvokeRequired)
            {
                try { this.BeginInvoke(new Action(() => SendMessageToWebView(message))); } catch { }
                return;
            }
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

    // 兜底：部分环境下 WebView 内按钮点击可能被吞，这里提供宿主热键强制切换
    // Ctrl+Shift+U：普通 ↔ 专家（并触发 reload）
    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        try
        {
            if (keyData == (Keys.Control | Keys.Shift | Keys.U))
            {
                ToggleUserModeFromHost();
                return true;
            }
        }
        catch { }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    private async void ToggleUserModeFromHost()
    {
        try
        {
            _webBootUserMode = string.Equals(_webBootUserMode, "expert", StringComparison.OrdinalIgnoreCase) ? "normal" : "expert";
            try { await EnsureBootUserModeScriptAsync(); } catch { }

            // 先给前端回执，便于观察宿主是否真的切换
            try { SendMessageToWebView(new { action = "userModeReloading", mode = _webBootUserMode, source = "hostHotkey" }); } catch { }

            // 触发 reload（优先虚拟域名映射）
            if (!string.IsNullOrWhiteSpace(_webHtmlPathCache) && File.Exists(_webHtmlPathCache))
            {
                try { _webBaseDirCache = Path.GetDirectoryName(_webHtmlPathCache); } catch { }
                try { EnsureVirtualHostMapping(); } catch { }
                try
                {
                    var ts = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
                    var entryFileName = GetEntryHtmlFileNameForMode(_webBootUserMode);
                    webView21.Source = new Uri($"https://{WebVirtualHost}/{entryFileName}?ts={ts}&mode={_webBootUserMode}");
                }
                catch { }
            }
            else
            {
                var html = string.IsNullOrWhiteSpace(_webHtmlContentCache) ? "" : _webHtmlContentCache;
                try { webView21.NavigateToString(html); } catch { }
            }
        }
        catch { }
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        _sqliteManager?.Dispose();
        base.OnFormClosing(e);
    }
}
