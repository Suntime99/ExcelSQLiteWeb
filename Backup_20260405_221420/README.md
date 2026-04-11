# Excel→SQLite 快速导入（ExcelDataReader 版）

你确认“可以装 NuGet”，那就能把 **导入慢** 的根因（EPPlus 逐单元格读取）替换掉。

本补丁用 **ExcelDataReader** 进行 **流式读取**（Streaming），通常会比 EPPlus 快很多，尤其是：
- 52 万行级别的大表
- 或列数较多、单元格访问频繁的表

---

## 1）需要安装的 NuGet 包

在你的 WinForms 项目（宿主）里安装：

1. `ExcelDataReader`
2. `ExcelDataReader.DataSet`（可选，本补丁不依赖 DataSet，但团队常会用到）
3. `System.Text.Encoding.CodePages`（**重要**：用于支持 `.xls` 编码）

> 说明：ExcelDataReader 同时支持 `.xlsx` 与 `.xls`。  
> `.xls` 必须注册编码 Provider：`Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);`

---

## 2）替换文件

将本目录下文件替换到你的项目：

- `Services/DataImporter.cs`

保持与现有 Form1.cs 调用一致：
```csharp
ImportWorksheetAsync(filePath, worksheetName, tableName, importMode, progress, cancellationToken)
```

---

## 3）导入模式（与前端全局参数一致）

本导入器支持：

### A. importMode = text（全列文本格式，最保真）
- 所有列建表为 TEXT
- 读到的值统一转字符串写入 SQLite
- 适合查询/检索/导出

### B. importMode = smart（智能转换列格式 + 记录失败）
- 先抽样最多 1000 行推断列类型
- 再按类型转换导入
- 转换失败写 NULL，并返回 `conversionStats`（失败次数与样本）

---

## 4）进度说明

ExcelDataReader 无法在不二次扫描的情况下快速得到总行数，因此进度条采用“活动式”反馈：
- 持续更新“已处理 N 行”
- 百分比为估算（保证 UI 有反馈，不假死）

如果你必须要“精确百分比”，只能：
1）先扫描一遍计数（会牺牲速度），或  
2）用 OpenXML 解析 sheet 的 dimension（实现复杂），或  
3）后台线程估算（体验还可以）

---

## 5）下一步建议（可选）

如果你还想再提速 1 个数量级：
- SQLite 插入改为 `SqliteCommand` + `Prepare()`（本补丁已经是预编译参数）
- 更激进：使用 `BEGIN;` 更大 batch（例如 20000）+ `PRAGMA cache_size` 调大
- 写入文件库（非内存）时，journal_mode/synchronous 需要更谨慎

