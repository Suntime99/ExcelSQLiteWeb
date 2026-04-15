# 项目架构层次与核心对象设计规范 (Phase 1)

为支持 ExcelSQLiteWeb 向多数据源、多模型体系（DataSources & Datasets）以及 AI 增强平滑演进，本规范将整个数据流转生命周期划分为 6 个层级，并对各层级的**职责边界**与**核心对象**进行严格定义。

---

## 1. 项目中心 (Project Hub)
**职责边界：**
全局的“工作台”入口。负责项目（Scheme/Project）生命周期的宏观管理，包括新建、切换、删除项目，以及全局设置（如脱敏 Vault 库配置、临时文件清理策略等）。该层**不触碰具体表数据**。

**核心对象：**
- `Project / Scheme`: 项目元数据容器（包含关联的 SQLite 数据库路径、工作目录等）。
- `Workspace`: 运行时的环境上下文。
- `SystemConfig`: 全局系统配置（临时库保留天数、UI 偏好等）。

---

## 2. 项目信息 (Project Info)
**职责边界：**
特定项目（Active Project）的概览面板。负责展示当前项目的宏观健康度、数据源概况、文件最后修改时间（用于校验是否需要重新导入）、以及整体容量（Row Count、Disk Size）。提供“导出分析报告”的统一入口。

**核心对象：**
- `ProjectMeta`: 项目的详情配置，包括：
  - `ProjectName`: 项目名称
  - `CreatedAt` / `UpdatedAt`: 时间戳
  - `SummaryMetrics`: 统计摘要（总表数、总行数等缓存快照）

---

## 3. 项目编辑 / 源文件关系 (Project Edit / Source File Relations)
**职责边界（重构核心重点）：**
负责**“外部数据源（External Data）”到“内部物理表（Raw SQLite Tables）”的映射与转换**。
不再局限于 Excel 工作表，而是统一抽象为 `DataSource`（文件/连接）与 `Dataset`（表/Sheet）。该层负责：
1. 外部数据源挂载与解析（如 CSV 分隔符配置、Excel 选 Sheet）。
2. 指定每个 Dataset 映射到 SQLite 中的哪张物理表（`TableName`），以及业务逻辑显示名（`DisplayName`）。
3. 指定事实表/维表角色（Fact/Dim），并控制导入（Import）或增量更新。

**核心对象：**
- `DataSource`: 外部数据源（如 `sales_2026.csv`, `finance.xlsx`）。
- `Dataset`: 具体的表实体（如 Excel 中的 `Sheet1`，CSV 的默认集）。
- `DatasetMapping`: 映射配置关系。
  - `SourceRef`: 指向具体的 Dataset。
  - `TargetTable`: SQLite 物理表名（如 `raw_sales`）。
  - `DisplayName`: UI 展示名称（如 “2026年销售明细”）。
  - `Role`: `Primary` (主表/事实表) 或 `Secondary` (副表/维表)。
- `ImportOptions`: 针对特定数据集的导入配置（如编码、分隔符、首行表头等）。

---

## 4. 建模管理 (Modeling Management)
**职责边界：**
负责在“物理表”的基础上，建立“逻辑模型”。将松散的多张表通过主外键或关联键（Join Keys）绑定在一起，形成**多维分析模型（Star Schema / Snowflake Schema）**。
产出物为**视图（View）**或**逻辑关系草稿（MTJ Draft）**，它不修改物理表数据，只生成 SQL 查询逻辑。

**核心对象：**
- `Relation / MTJ (Multi-Table Join)`: 表与表之间的关联关系（`LeftTable.Key = RightTable.Key`）。
- `DerivedView`: 派生视图（以 `vw_` 命名）。
- `Dimension / Metric`: 业务维度的定义（如将某时间戳字段定义为 `YearMonth` 维度）。

---

## 5. 数据库管理 (Database Management)
**职责边界：**
底层执行引擎。负责 SQLite 数据库生命周期、连接池、事务控制（Transaction）、并发锁（WAL 模式）、以及原生 SQL 脚本的执行与日志记录。系统内所有的数据流（导入、建模、分析）最终都转化为 SQL 语句交由该层执行。

**核心对象：**
- `SqliteManager / DbContext`: 数据库连接与事务上下文。
- `QueryPlan / EXPLAIN`: SQL 执行计划。
- `ExecutionResult`: 包含受影响行数、耗时、报错信息的标准返回体。

---

## 6. 对象管理 / 索引管理 (Object Management / Index Management)
**职责边界：**
负责物理层面的性能优化与存储管理。对 SQLite 库内的实体对象（Table, View, Index, Trigger）进行直接干预。包括基于慢查询日志或 AI 建议，一键创建或删除索引（Index）；清理无用的临时表或历史视图（Vacuum/Drop）。

**核心对象：**
- `DbObject`: 数据库物理对象（Type: table/view/index）。
- `IndexStrategy`: 索引策略（如针对 `raw_sales(user_id)` 创建 B-Tree 索引）。
- `SchemaInfo`: 底层 SQLite 的 `sqlite_master` 表映射数据。