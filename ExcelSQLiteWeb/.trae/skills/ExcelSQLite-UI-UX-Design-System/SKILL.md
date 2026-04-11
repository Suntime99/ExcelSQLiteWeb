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
