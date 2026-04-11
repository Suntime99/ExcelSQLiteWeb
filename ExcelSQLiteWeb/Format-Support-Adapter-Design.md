# ExcelSQLite 多格式导入适配器方案（设计稿）

> 目标：在不破坏现有“Excel → SQLite → 查询/清洗/统计/导出”主流程的前提下，扩展支持 **CSV / Access / DBF** 等常用数据源，并形成可持续扩展的“导入适配器（Importer Adapter）”架构。

## 1. 现状与问题
当前工程的“打开文件/方案管理”主要围绕 Excel 工作表（sheet）展开：
- UI 选择文件 → 选择工作表 → 设为主表/副表 → 导入 SQLite
- 后续所有功能基于 SQLite 表运算（查询、清洗、统计、校验、导出）

要支持 CSV/Access/DBF，需要解决：
1) 数据源不再天然“sheet”，需要统一抽象为“数据集（Dataset）”。  
2) 编码/分隔符/类型推断/大文件性能，需要统一策略。  
3) 宿主侧实现应可渐进扩展：先 CSV，再 Access/DBF。

## 2. 核心设计：数据集抽象
定义统一概念：
- **DataSource**：一个文件路径（或连接串）
- **Dataset**：数据源内的一个可导入实体（Excel 的 sheet；Access 的 table/query；DBF 的表；CSV 的单表）

统一 UI 展示为：
- 文件清单（DataSource）
- 数据集清单（Dataset）——每个文件展开后显示 datasets
- 勾选 datasets 导入 SQLite（每个 dataset → 一个 SQLite table）

## 3. 导入适配器接口（宿主端）
建议在宿主端定义接口（伪代码）：
```csharp
public interface IDataImporter
{
    bool CanHandle(string filePath);
    string Kind { get; } // excel/csv/access/dbf...
    IEnumerable<DatasetInfo> ListDatasets(string filePath);
    DatasetSchema GetSchema(string filePath, string datasetId);
    IAsyncEnumerable<RowData> ReadRows(string filePath, string datasetId, ImportOptions options);
}
```

### 3.1 DatasetInfo
```csharp
public record DatasetInfo(
  string Id,          // sheetName / tableName / "CSV"
  string DisplayName, // UI显示名
  long? EstimatedRows // 可选
);
```

### 3.2 ImportOptions（CSV/大文件关键）
- Encoding：UTF-8/GBK/自动探测
- Delimiter：, \t ; |
- HasHeader：是否首行为表头
- TypeInference：自动推断/全列文本/指定列类型
- SampleRowsForInference：推断采样行数（默认 200~1000）
- MaxRows：限制导入行数（安全阈值）
- NullTokens：空值标记（"", "NULL", "N/A" 等）

## 4. 前端（WebView）消息协议建议
新增 action（名称可按现有风格调整）：
1) `listDatasets`：传 filePath → 返回 datasets  
2) `getDatasetSchema`：filePath + datasetId → 返回 schema  
3) `importDatasetToSqlite`：filePath + datasetId + options + tableName → 执行导入，返回进度与结果  

返回事件：
- `datasetsListed`  
- `datasetSchemaLoaded`  
- `importProgressUpdate`  
- `importComplete`

## 5. 分阶段落地计划（你已选择“先做方案设计”）
### Phase 0（本次）：完成接口与 UI 结构准备
- UI：将“工作表清单”升级为“数据集清单”（Excel 仍按 sheet 展示，CSV 暂时单 dataset）
- 代码：实现 importer 框架 + ExcelImporter 适配到新接口（不改变旧功能）

### Phase 1：CSVImporter（优先落地）
原因：实现成本最低、需求最普遍、能验证架构。
- 支持：编码探测、分隔符选择、首行表头、全列文本/类型推断
- 性能：流式读取 + 批量插入 SQLite（事务 + prepared statement）
- 安全：默认采样推断 + 默认最大行数阈值（可在高级选项调整）

### Phase 2：AccessImporter（.mdb/.accdb）
- 依赖：OLEDB/ACE 驱动（环境差异大，需要“驱动缺失提示”与降级方案）
- dataset：table / query
- schema：字段类型映射到 SQLite（TEXT/REAL/INTEGER）

### Phase 3：DBFImporter
- 依赖：DBF 解析库（注意中文编码：GBK/Big5）
- dataset：单 DBF 文件=单表（或目录批量）

## 6. 与现有模块的兼容策略
1) **所有后续功能只依赖 SQLite 表**：Importer 只负责把外部数据转成 SQLite 表。  
2) **导出统一样式**：导出 Excel 使用“微软雅黑 9pt + 居中 + 合理行高列宽”（CSV 无样式）。  
3) **方案/模板复用**：DataSource + Dataset + ImportOptions 均应进入“方案”配置，确保可复现。

## 7. 需要你确认的 3 个实现选择（下一步再开工前确认）
1) CSV：默认是否“全列文本”还是“类型推断”？（建议默认全列文本更安全）  
2) CSV：默认编码探测策略（仅 UTF-8/GBK 两档？还是引入更强探测库）  
3) Access：是否接受“必须安装 ACE 驱动”的前置条件？（否则需要内置解析库，成本高）

