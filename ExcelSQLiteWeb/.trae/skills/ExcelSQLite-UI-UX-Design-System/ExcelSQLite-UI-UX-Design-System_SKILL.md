---
name: "ExcelSQLite-UI-UX-Design-System"
description: "ExcelSQLite项目UI/UX设计规范与开发指南。Invoke when designing UI components, implementing new features, or optimizing user experience for ExcelSQLite."
---

# ExcelSQLite UI/UX 设计系统

## 一、设计规范总览

### 1.1 字体规范
- **主字体**: Microsoft YaHei（微软雅黑）
- **字号体系**:
  - 标题: 10-11pt，加粗
  - 正文: 9pt，不加粗
  - 辅助文字: 8pt，灰色系
  - 代码/数据: Consolas/Monaco，9-10pt

### 1.2 配色体系
```css
/* 主色调 */
--primary-blue: #0078d4;          /* Office经典蓝 */
--primary-hover: #106ebe;         /* 悬停状态 */
--primary-active: #005a9e;        /* 激活状态 */

/* 中性色 */
--bg-primary: #f0f0f0;            /* 主背景 */
--bg-secondary: #f8f9fa;          /* 次级背景 */
--bg-card: #ffffff;               /* 卡片背景 */
--border-color: #e1e4e8;          /* 边框色 */
--text-primary: #333333;          /* 主文字 */
--text-secondary: #666666;        /* 次级文字 */
--text-muted: #999999;            /* 辅助文字 */

/* 功能色 */
--success: #52c41a;               /* 成功 */
--warning: #faad14;               /* 警告 */
--error: #f5222d;                 /* 错误 */
--info: #1890ff;                  /* 信息 */
```

### 1.3 间距规范
```css
/* 基础间距 */
--spacing-xs: 4px;
--spacing-sm: 8px;
--spacing-md: 12px;
--spacing-lg: 16px;
--spacing-xl: 24px;

/* 组件间距 */
--card-padding: 16px;
--button-padding: 6px 12px;
--input-padding: 8px 12px;
--table-cell-padding: 8px 12px;
```

### 1.4 圆角与阴影
```css
/* 圆角 */
--radius-sm: 3px;                 /* 按钮、输入框 */
--radius-md: 4px;                 /* 卡片、面板 */
--radius-lg: 8px;                 /* 弹窗、模态框 */

/* 阴影 */
--shadow-sm: 0 1px 2px rgba(0,0,0,0.05);
--shadow-md: 0 2px 8px rgba(0,0,0,0.1);
--shadow-lg: 0 4px 16px rgba(0,0,0,0.15);
```

## 二、布局架构

### 2.1 整体布局
```
┌─────────────────────────────────────────────────────────┐
│  顶部标题栏 (Title Bar)                                    │
├─────────────────────────────────────────────────────────┤
│  Ribbon功能区 (File, Home, Data, View, Help)            │
├─────────────────────────────────────────────────────────┤
│  文件信息栏 (File Info: 主表/副表/工作表)                  │
├──────────────┬──────────────────────────────────────────┤
│              │                                          │
│  左侧导航栏   │         主内容区 (Main Content)           │
│  (Sidebar)   │                                          │
│              │  ┌────────────────────────────────────┐  │
│  - 最近文件   │  │  向导容器 (Wizard Container)        │  │
│  - 常用操作   │  │  ┌────────────────────────────────┐│  │
│  - 工具      │  │  │ 步骤导航 (Step Navigation)     ││  │
│              │  │  ├────────────────────────────────┤│  │
│              │  │  │ 步骤内容 (Step Content)        ││  │
│              │  │  │ ┌────────────────────────────┐ ││  │
│              │  │  │ │ 卡片 (Card)                │ ││  │
│              │  │  │ │ ┌────────────────────────┐ │ ││  │
│              │  │  │ │ │ 表单组 (Form Group)    │ │ ││  │
│              │  │  │ │ │ ┌────────────────────┐ │ │ ││  │
│              │  │  │ │ │ │ 控件 (Controls)   │ │ │ ││  │
│              │  │  │ │ │ └────────────────────┘ │ │ ││  │
│              │  │  │ │ └────────────────────────┘ │ ││  │
│              │  │  │ └────────────────────────────┘ ││  │
│              │  │  └────────────────────────────────┘│  │
│              │  └────────────────────────────────────┘  │
│              │                                          │
└──────────────┴──────────────────────────────────────────┘
```

### 2.2 Ribbon功能区分组
1. **文件管理**: 打开文件、文件分析、工作表分析、元数据扫描
2. **数据清洗**: 数据清洗、差异比对
3. **SQL实验室**: SQL编辑器、执行SQL、SQL历史
4. **数据查询**: 单表查询、多表查询、全局搜索
5. **数据统计**: 单表统计、多表统计
6. **文件分拆**: 单表分拆、多表分拆
7. **信息安全**: 数据脱敏
8. **数据可视化**: 图表生成、透视表

### 2.3 响应式断点
```css
/* 桌面端 */
@media (min-width: 1024px) {
  --sidebar-width: 220px;
  --ribbon-height: 100px;
}

/* 平板端 */
@media (max-width: 1023px) {
  --sidebar-width: 180px;
  --ribbon-height: 90px;
}

/* 移动端 */
@media (max-width: 768px) {
  --sidebar-width: 0px;       /* 隐藏侧边栏 */
  --ribbon-height: auto;      /* 自适应高度 */
}
```

## 三、组件规范

### 3.1 按钮 (Button)
```html
<!-- 主按钮 -->
<button class="btn btn-primary">主要操作</button>

<!-- 次按钮 -->
<button class="btn btn-secondary">次要操作</button>

<!-- 成功按钮 -->
<button class="btn btn-success">确认/执行</button>

<!-- 危险按钮 -->
<button class="btn btn-danger">删除/退出</button>

<!-- Ribbon按钮 -->
<button class="ribbon-btn">功能<br>名称</button>

<!-- 向导按钮 -->
<button class="wizard-btn wizard-btn-help">❓ 帮助</button>
<button class="wizard-btn wizard-btn-save">💾 保存</button>
<button class="wizard-btn wizard-btn-prev">← 上一步</button>
<button class="wizard-btn wizard-btn-next">下一步 →</button>
<button class="wizard-btn wizard-btn-execute">▶ 执行</button>
<button class="wizard-btn wizard-btn-exit">✕ 退出</button>
```

**样式规范**:
```css
.btn {
  padding: 6px 12px;
  border-radius: 3px;
  font-size: 9pt;
  font-family: 'Microsoft YaHei', sans-serif;
  cursor: pointer;
  transition: all 0.2s ease;
  border: 1px solid transparent;
}

.btn-primary {
  background: linear-gradient(180deg, #0078d4 0%, #106ebe 100%);
  color: white;
  border-color: #005a9e;
}

.btn-primary:hover {
  background: linear-gradient(180deg, #106ebe 0%, #005a9e 100%);
  box-shadow: 0 2px 4px rgba(0,0,0,0.2);
}

.ribbon-btn {
  width: 60px;
  font-size: 8pt;
  padding: 4px;
```

### 3.2 步骤导航（Step Navigation）

**默认方案（节省页面空间）**：文字在圆圈右侧（胶囊标签），9pt。  
**回落方案（稳妥保底）**：文字在圆圈下方，8pt。

#### 3.2.1 等距与对齐规则（必须）
1) 每个步骤的圆圈中心必须等距（建议 `flex:1` 均分，避免手工 `gap` 导致不同窗口宽度下失真）。  
2) 连接线两端与“首/尾圆圈中心”对齐：  
   - 若步骤数为 `N`，连接线左右内缩为 `50/N %`。  
   - 示例：`N=4 → 12.5%`；`N=3 → 16.666%`。
3) “步骤组合”（圆圈+文字）之间不得互相覆盖；长文案必须截断/折行/回落到下方样式。

#### 3.2.2 文字在右侧的样式建议（避免压线）
- 将文字做成“胶囊标签”（白底/细边框），遮住连线的一小段，避免文字压在线条上显得凌乱。  
- 字号推荐 9pt；若空间不足，则回落到“下方 8pt”。

> 回落原则：当任一步骤文案超过 8~10 个汉字（或容器宽度不足）时，不强行右置，直接使用“下方 8pt”版本。

### 3.5 临时模板（跨启动自动恢复）
**目标**：用户未主动“保存模板”时，系统仍能在“下次启动/下次进入模块”自动恢复上一次配置，降低重复配置成本。  
规则：
1) 每次执行前，如果用户未勾选“保存模板”，自动覆盖写入名为 **【临时模板】** 的模板（每模块/每模式独立一份）。  
2) 启动/进入模块时默认自动加载【临时模板】，但用户可手工切换为其他模板。  
3) 临时模板应持久化（建议 localStorage 或方案文件），保证跨启动生效。

### 3.3 向导底部工具栏（Wizard Bottom Bar）

**目标**：减少鼠标移动，主操作恒定；参照“单表查询”的 `wizard-bottom-bar` 模式。

关键约束：
1) `wizard-bottom-bar` 必须放在 **向导内容末尾（所有 step-content 之后）**，否则在“仅显示当前 step-content”时可能被隐藏/滚动体验异常。  
2) 底部按钮组使用 `.wizard-buttons.wizard-bottom`（复用视觉规范），按钮使用 `.wizard-btn` 体系。  
3) 所有向导底部栏按钮顺序建议：上一步 / 下一步 / 主执行 / 生成SQL /（可选菜单）/ 退出。

### 3.4 数据清洗输出规范（字段格式清洗 / 数据质量清洗 / 业务规则验证）

#### 3.4.1 输出选项（统一四选一）
1) 覆盖当前数据库表（SQLite）：`db_overwrite`  
2) 生成数据库表副本（SQLite）：`db_copy`（需要 `dbCopyTableName`）  
3) 覆盖源文件（Excel）：`file_overwrite`  
4) 生成新文件（Excel/CSV）：`file_new`（需要 `outputPath`）

#### 3.4.2 前后端字段约定（message/settings）
前端 `settings.outputOptions` 建议字段：
```json
{
  "target": "db_overwrite | db_copy | file_overwrite | file_new",
  "dbCopyTableName": "Cleansed_xxx",
  "format": "xlsx | csv",
  "outputPath": "C:/.../xxx.xlsx"
}
```

#### 3.4.3 后端执行语义（建议）
- `db_overwrite`：使用“临时表→替换源表”的方式覆盖，避免直接 DROP 导致中间态不可用。  
- `db_copy`：生成指定命名的新表，不影响源表。  
- `file_overwrite`：写回当前打开的源文件路径（风险二次确认）。  
- `file_new`：导出到指定路径/格式（建议强制全列文本防止科学计数法）。

---

# 需求/决策记录（会话纪要）

> 目的：把本轮会话中与 ExcelSQLiteWeb 相关的需求、决策、约束、已完成改动与后续实施路径记录在案，供后续迭代与验收对照。  
> 时间：2026-04-05 ~ 2026-04-06

## 0. 当前工程与协作方式
- 工程：ExcelSQLiteWeb（WinForms + WebView2 + SQLite + 前端 index.html）
- 本次交付方式：用户通过 Web 版上传 7z 工程包；助手在本地解压修改后回传更新包。
- 用户环境约束：Win10 专业版、公司电脑、可能受域组策略控制；VS 可能后台更新并提示重启。

## 1. 已确认的全局设计原则
1) 小白安全优先：默认受控/只读/预览模式；高级能力通过“专家模式/高级设置”开启。  
2) 统一入口与闭环：所有页面产生的 SQL，最终应能回流到“SQL实验室/SQL编辑器”执行、保存、复用。  
3) 性能保护：默认 LIMIT/采样/阈值；避免全表全字段全量扫描造成卡死。  
4) 可解释：结果尽量提供“来源、统计口径、SQL 可查看”，支持导出报告（本轮优先 HTML 单文件）。

## 2. 搜索与联合搜索（阶段改动记录）
### 2.1 RIBBON 搜索分组与入口
- 调整“全局搜索”分组入口为：单表搜索 / 多表联合搜索 / 全局联合搜索。
- 删除【主表搜索】（含 RIBBON 按钮与页面）。
- 原【指定表搜索】改名为【单表搜索】。

### 2.2 搜索结果口径（按“字段维度汇总”）
结果列表统一为：**来源表、来源字段名称、匹配内容、记录数、操作（查看明细、查看SQL）**。  
由“按 rowid 逐条展示”调整为“按字段维度汇总计数”，并提供明细与 SQL 查看入口。

### 2.3 SQLite UNION/LIMIT 相关异常（已定位为生成SQL结构问题）
曾出现：
- `LIMIT clause should come after UNION ALL not before`
- `near "UNION": syntax error`
结论：SQLite compound select 对分支 LIMIT 有语法限制，最终采用“字段维度 COUNT 统计”的生成策略避免该类错误。

### 2.4 记录明细（多窗口弹层）
- 支持多开弹层；手工改 SQL；每次加载条数；深色表头；仅允许 SELECT。

## 3. 多表联合查询/统计：受控模式/自由模式（已实现）
- 在【多表联合查询】【多表联合统计】手工 SQL 区增加“受控模式”开关（默认受控）。  
- 受控模式：执行前校验 SQL 仅可引用 MTJ 锁定的表/字段池（含 rowid 特判）。  
- 自由模式：放开白名单校验（但仍保留“只允许 SELECT”的硬约束）。

## 4. 单表切换（已实现并加硬）
### 4.1 单表查询/单表统计支持“单表切换”
- 【单表查询】【单表统计】数据源改为可切换的“单表切换”（首次默认主表）。

### 4.2 安全策略（已确认）
- 只能在 Step1 切换才安全：Step2/Step3 禁止切换（禁用下拉）。  
- 增加“↩ 回到 Step1 并重置”按钮（单表查询/单表统计各自一套），用于安全清空配置。  
- Step1 切换时自动静默重置到 Step1，避免配置残留误用。

## 5. 本轮待实施的大项需求清单（用户正式指令）

### 5.1 文件分析优化
1) 去除分析选项界面显示（或全选置灰，按代码实际是否使用决定）。  
2) 修复：文件分析只能识别主表的 BUG。  
3) 修复：操作栏【查看字段】无响应的 BUG。  
4) 增加：操作栏【查询SQL】按钮及功能。  
5) 增加：页面【导出分析报告】按钮及功能（HTML 单文件）。  
6) 工作表清单上方信息框合并为【文件总体概览】分组。  
7) 【工作表清单】更名为【数据地图（工作表清单）】。

### 5.2 工作表分析优化
1) 页面【查询SQL】下拉：纯查询SQL / 带字段别名的查询SQL。  
2) 页面增加【导出分析报告】按钮及功能（HTML 单文件）。  
3) 去除分析选项界面显示（或全选置灰）。

### 5.3 元数据扫描（按约定 P0 重做）
- 扫描选项参数真实生效（includeBasic/includeSheets/includeStats/includeRelations）。  
- 分层展示：文件级 → 工作表级 → 字段级；支持 drill-down。  
- 增加进度/状态/可取消（至少粗粒度）。  
- 增加“一键联动”：加入源文件清单/设为主表/勾选导入/触发导入。

### 5.4 关联关系识别 / 系统检查（按约定）
- 默认分析范围：当前已导入 SQLite 的全部表。  
- 两种模式：主表优先（默认）+ 两两互扫（可选，用户参数控制）。  
- 输出为“候选 MTJ 草稿”，需确认再应用（先按草稿实现）。

### 5.5 左侧栏【最近文件】→【最近方案】（宿主端 + 前端联动）
1) 方案名称：打开文件页可录入，也可自动智能生成。  
2) 一个方案可包含多个文件组合；切换方案自动回填到打开文件页配置。  
3) 方案配置：保存在 `ExcelSQLiteWeb.exe` 所在目录，采用 INI。  
4) 最近方案显示：默认显示 10 个（INI 可配）；超过 10 条保留不显示；超过 50 条覆盖最旧。  
5) 程序启动自动加载：上次方案 + SQLite 临时库；增加文件修改时间/指纹校验，不一致提示需重新导入。  
6) 临时库：允许在 exe 目录下建子目录；数据库按方案命名；INI 记录映射。  
7) RIBBON 顶部增加【清理临时数据库】按钮（用于超期文件删除/手工清理）。

### 5.6 SQL实验室（SQL编辑器/执行SQL/SQL历史）
- Phase0：统一执行入口 + 历史可用化（localStorage）+ EXPLAIN 接口预留。  
- 后续：编辑器增强（优先 CodeMirror）、执行计划、受控/自由模式与模板库。

## 6. 导出报告（统一口径）
- 导出格式：HTML（单文件，CSS 内联）。  
- 内容：统计与字段信息为主；不输出样本数据（避免敏感信息泄露）。

## 7. 宿主端消息协议（实施时同步维护）
预计新增/扩展 WebMessage action（名称以最终代码为准）：
- scheme：saveScheme / loadScheme / listSchemes / deleteScheme / loadLastSchemeOnStartup  
- tempdb：cleanupTempDb / openTempDbFolder  
- report：exportFileAnalysisReport / exportWorksheetAnalysisReport  
- sql-lab：saveSqlHistory / listSqlHistory / pinSqlHistory / deleteSqlHistory / explainQueryPlan  
- relation：detectRelations(primary-first/pairwise) / applyRelationDraftToMtj

---

# 工程对照清单（与当前工程实现比对）

> 用途：用于“规范 ⇄ 代码”双向对齐。下面按 **已符合 / 偏差 / 待确认** 列出，便于你决定是补充规范，还是改代码。

## A. 已符合（当前工程与规范一致或基本一致）
1) **字体体系**：全局以微软雅黑 9pt 为主，代码区域使用 Consolas/Monaco（9~10pt）。  
2) **主色调**：大量使用 `#0078d4 / #106ebe / #005a9e`（Office 蓝系）。  
3) **组件形态**：按钮、卡片、表格、向导布局整体已趋近“Office/Fluent”风格。  
4) **向导底部工具栏**：单表查询已采用 `wizard-bottom-bar` + `wizard-buttons wizard-bottom` 固化主操作区，其他向导逐步对齐中。  

## B. 偏差点（当前工程与规范不一致/不完整）
1) **CSS 变量体系未落地**：规范中给出了 `--primary-blue / --spacing-* / --radius-*` 等 token，但当前工程主要使用“硬编码颜色 + 大量 inline style”。  
   - 影响：后续 UI 调整成本高、全局一致性难保证。  
2) **响应式断点与侧边栏宽度**：规范示例中 sidebar 220/180/0，当前工程 sidebar 固定 200（且部分控件使用固定宽度）。  
3) **步骤条（Step Navigation）文字布局**：规范未明确“文字在右/下”的二选一策略；当前工程存在两种诉求：  
   - 右侧标签：信息密度高，但更容易出现“压线/遮挡/不等距”问题；  
   - 下方文字：更稳妥，但占用垂直空间。  
4) **数据清洗输出选项**：规范原先缺失统一口径（已在本文件 3.4 补齐），但仍需要你确认“业务规则验证”是否也要走同一套 4 输出（见下）。  

## C. 待你补充/确认（确认后我可再补规范/或改代码）
1) **步骤条默认样式**：最终是否统一为“文字下方 8pt”（推荐稳妥），还是允许“右侧胶囊标签 9pt”作为默认？  
2) **UI Token 化改造范围**：是否要把颜色/间距/圆角逐步抽成 CSS variables（例如先覆盖 btn/card/step/navigation 四大类）？  
3) **数据清洗三模块的输出一致性**：  
   - 字段格式清洗、数据质量清洗：已支持 4 输出；  
   - **业务规则验证**：你希望输出为  
     a) 只生成 SQL/报告（不改表），还是  
     b) 也支持 4 输出（把“异常记录/校验结果”落库或导出文件）？
