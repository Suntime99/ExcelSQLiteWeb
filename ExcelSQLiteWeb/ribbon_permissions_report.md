# Ribbon 分组与权限清单

> 说明：`data-expert-only="1"` 表示仅专家可见（普通用户会被隐藏）。

## 关键实现代码
- `setUserMode(mode)`：写入 `window.__userMode` 与 localStorage
- `applyUserModeToRibbon()`：按 `data-expert-only` 显隐 Ribbon/Sidebar
- `applyUserModeToPages()`：按 `data-expert-only` 显隐页面

## Ribbon Tab: `start`
### 组：项目管理
| 按钮 | expertOnly | 图标 | onclick |
|---|---:|---|---|
| 项目管理 | 否 | ico-folder-open | switchTab('file-wizard') |
| 文件分析 | 否 | ico-document-search | switchTab('file-analysis') |
| 工作表分析 | 是 | ico-document-page | switchTab('worksheet-analysis') |

### 组：数据质量
| 按钮 | expertOnly | 图标 | onclick |
|---|---:|---|---|
| 业务规则验证 | 否 | ico-broom | switchTab('dc-verify') |
| 数据脱敏 | 否 | ico-shield | switchTab('data-masking') |

## Ribbon Tab: `project`
### 组：项目管理
| 按钮 | expertOnly | 图标 | onclick |
|---|---:|---|---|
| 项目中心 | 否 | ico-chart | switchTab('project-hub') |
| 项目管理 | 否 | ico-folder-open | switchTab('file-wizard') |
| 建模管理 | 是 | ico-table-multiple | switchTab('model-management') |

### 组：文件表管理
| 按钮 | expertOnly | 图标 | onclick |
|---|---:|---|---|
| 文件分析 | 否 | ico-document-search | switchTab('file-analysis') |
| 工作表分析 | 是 | ico-document-page | switchTab('worksheet-analysis') |
| 元数据分析 | 是 | ico-scan | switchTab('metadata-scan') |
| 关联关系识别 | 是 | ico-search | switchTab('relation-check') |
| 多表关联 | 是 | ico-table-multiple | openMultiTableJoin() |

### 组：全局搜索
| 按钮 | expertOnly | 图标 | onclick |
|---|---:|---|---|
| 单表搜索 | 否 | ico-search | switchTab('specific-table-search') |
| 多表联合搜索 | 是 | ico-table-multiple | switchTab('mtj-search') |
| 全局联合搜索 | 是 | ico-scan | switchTab('global-union-search') |

## Ribbon Tab: `quality`
### 组：数据清洗
| 按钮 | expertOnly | 图标 | onclick |
|---|---:|---|---|
| 字段格式清洗 | 否 | ico-broom | switchTab('dc-field') |
| 数据质量清洗 | 否 | ico-data-bar | switchTab('dc-quality') |
| 业务规则验证 | 否 | ico-check | switchTab('dc-verify') |

### 组：差异比对
| 按钮 | expertOnly | 图标 | onclick |
|---|---:|---|---|
| 差异比对 | 否 | ico-arrow-swap | switchTab('data-compare') |

### 组：数据脱敏
| 按钮 | expertOnly | 图标 | onclick |
|---|---:|---|---|
| 数据脱敏 | 否 | ico-shield | switchTab('data-masking') |
| 安全工作台 | 是 | ico-settings | openSecurityWorkbench() |
| 枚举值存档 | 是 | ico-list | openEnumArchive() |

## Ribbon Tab: `query`
### 组：单表查阅
| 按钮 | expertOnly | 图标 | onclick |
|---|---:|---|---|
| 单表查询 | 否 | ico-table | switchTab('single-table-query') |
| 单表统计 | 否 | ico-data-bar | switchTab('single-table-stats') |

### 组：多表查阅
| 按钮 | expertOnly | 图标 | onclick |
|---|---:|---|---|
| 关联关系识别 | 否 | ico-search | switchTab('relation-check') |
| 多表关联 | 否 | ico-table-multiple | openMultiTableJoin() |
| 多表联合查询 | 是 | ico-table-multiple | switchTab('multi-table-query-lite') |
| 多表联合统计 | 是 | ico-data-bar | switchTab('multi-table-stats-lite') |

### 组：可视化分析
| 按钮 | expertOnly | 图标 | onclick |
|---|---:|---|---|
| 图表分析 | 否 | ico-chart | switchTab('chart-generator') |
| 透视表分析 | 否 | ico-pivot | switchTab('pivot-table') |

### 组：数据分拆
| 按钮 | expertOnly | 图标 | onclick |
|---|---:|---|---|
| 数据分拆 | 否 | ico-cut | switchTab('single-table-split') |

## Ribbon Tab: `expert`
### 组：SQL试验室
| 按钮 | expertOnly | 图标 | onclick |
|---|---:|---|---|
| SQL编辑器 | 否 | ico-code | switchTab('sql-editor') |
| 执行SQL | 否 | ico-play | switchTab('sql-generator') |
| SQL历史记录 | 否 | ico-history | switchTab('sql-history') |

### 组：数据库管理
| 按钮 | expertOnly | 图标 | onclick |
|---|---:|---|---|
| 对象管理 | 否 | ico-table-multiple | switchTab('db-object-manager') |

### 组：批量作业
| 按钮 | expertOnly | 图标 | onclick |
|---|---:|---|---|
| 批量处理 | 否 | ico-batch | switchTab('batch-process') |

## Ribbon Tab: `help`
### 组：个性化
| 按钮 | expertOnly | 图标 | onclick |
|---|---:|---|---|
| 多语言 | 否 | ico-translate | showLanguage() |
| 个性化 | 否 | ico-settings | showPersonalization() |

### 组：帮助指引
| 按钮 | expertOnly | 图标 | onclick |
|---|---:|---|---|
| 学习案例 | 否 | ico-book | showLearningCases() |
| 帮助指引 | 否 | ico-help | showHelp() |

### 组：关于
| 按钮 | expertOnly | 图标 | onclick |
|---|---:|---|---|
| 关于 | 否 | ico-info | showAbout() |
