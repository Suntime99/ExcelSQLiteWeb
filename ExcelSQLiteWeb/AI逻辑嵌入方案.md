---
title: AI 逻辑嵌入方案（能力矩阵 + 协议草案）
version: v0.1
---

本文档只做“**逻辑嵌入**”：定义 AI 在系统内的职责边界、触发点、输入输出协议与落盘审计方式；不涉及具体模型厂商、网络调用与计费实现。

---

## 1. 总原则（强约束）

1) **AI 默认只给建议/草稿/解释，不自动执行**  
任何会改变数据库/项目配置的动作（DDL/DML、写入项目文件）必须经过用户确认，并走现有审计链路。

2) **AI 输出必须结构化 + 可追溯**  
每条建议带：理由、置信度、风险等级、关联到的“上下文证据”（表/字段/统计/错误日志）。

3) **安全策略优先**  
支持能力级别的策略开关：例如 SQL Copilot 默认禁止生成 DDL/DML；或只允许生成“候选 SQL”，不直接执行。

---

## 2. AI 能力清单（Capabilities）

| Capability | 作用 | 典型输入 | 典型输出 | 可一键动作（需确认） |
|---|---|---|---|---|
| `modeling_copilot` 建模助手 | 辅助事实/维表划分、JOIN 路径、视图命名 | 表结构、字段名、样例统计、主表、用户目标 | 模型建议、JOIN 路径建议、命名建议 | 生成 MTJ 草稿 / 生成 View 草稿 SQL |
| `relationship_advisor` 关系识别建议 | 对候选关联键排序、解释与组合键建议 | 关系识别候选、命中率/基数、样例值 | 置信度排序、组合键建议、清洗建议 | 生成 JOIN 条件草稿 / 锁定到建模草稿 |
| `sql_copilot` SQL 指引与校正 | 错误解释、修复建议、性能建议（人话版） | SQL、错误信息、EXPLAIN、索引列表、参数 | 修复版 SQL、原因解释、性能建议 | 应用修复到编辑器 / 生成索引建议说明 |
| `data_analyst` 数据分析解读 | 统计/分布/异常点业务化解读 | 汇总统计、缺失率、分布、异常点 | 摘要结论、风险点、下一步分析建议 | 生成报告附录（文本） |
| `masking_interpreter` 脱敏解读与策略 | 给出脱敏方案并解释影响 | 字段分类、样例、合规目标 | 推荐脱敏方法、可逆性说明、影响说明 | 生成脱敏规则草稿 |

---

## 3. 触发点矩阵（页面/模块 × AI）

> 约定：**默认不自动弹出**；采用“AI 面板/按钮”显式触发。仅对明显错误（SQL 执行报错）可以提供“建议弹窗入口（不自动修复）”。

### 3.1 项目编辑（源文件关系）
- AI：`modeling_copilot`、`relationship_advisor`
- 触发按钮：
  - “AI 建议主表/维表角色”
  - “AI 推荐表名/别名”
- 输入：
  - 文件-工作表映射、字段名抽样、用户指定的业务目标（可选）
- 输出展示：
  - 右侧 AI 面板（建议列表 + 置信度/风险）

### 3.2 建模管理（关系识别/多表关联）
- AI：`relationship_advisor`（主）、`modeling_copilot`
- 触发按钮：
  - “AI 解释候选关系”
  - “AI 推荐组合键/清洗”
  - “AI 生成 MTJ 草稿”

### 3.3 SQL 实验室/查询分析器
- AI：`sql_copilot`（主）
- 触发按钮：
  - “AI 解释报错”
  - “AI 修复 SQL（写入编辑器，不执行）”
  - “AI 解释 EXPLAIN（人话版）”
  - “AI 性能建议（结合现有索引建议）”
- 输入：
  - SQL、错误堆栈/消息、EXPLAIN 文本、现有索引信息、参数

### 3.4 分析类页面（文件分析/工作表分析/统计结果）
- AI：`data_analyst`
- 触发按钮：
  - “AI 解读本页统计”
  - “AI 生成结论摘要（用于报告）”

### 3.5 脱敏/清洗
- AI：`masking_interpreter`、（可选）`data_analyst`
- 触发按钮：
  - “AI 推荐脱敏策略”
  - “AI 解释脱敏影响（交付说明）”

---

## 4. 统一协议草案（AIRequest / AIResponse）

### 4.1 AIRequest（前端 → 宿主/AI层）

```json
{
  "requestId": "ai-20260411-0001",
  "capability": "sql_copilot",
  "context": {
    "page": "sql-lab",
    "project": {
      "projectId": "p-xxx",
      "projectName": "项目_20260411"
    },
    "db": {
      "dbPath": "project.sqlite",
      "mainTable": "fact_orders"
    },
    "sql": "select * from t where a = ;",
    "params": {
      "p": { "type": "int", "value": "1" }
    },
    "error": {
      "message": "near \";\": syntax error",
      "stack": ""
    },
    "explain": {
      "planText": "SCAN TABLE t"
    },
    "indexes": [
      { "name": "idx_t_a", "table": "t", "columns": ["a"], "unique": false }
    ],
    "stats": null
  },
  "constraints": {
    "allowGenerateDDL": false,
    "allowGenerateDML": false,
    "maxTokensHint": 800,
    "language": "zh-CN"
  },
  "userGoal": "修复 SQL 并解释原因"
}
```

### 4.2 AIResponse（AI层 → 前端）

```json
{
  "requestId": "ai-20260411-0001",
  "ok": true,
  "summary": "SQL 在 WHERE 条件中缺少右值，导致语法错误。",
  "recommendations": [
    {
      "id": "rec-1",
      "title": "修复 WHERE 条件",
      "detail": "将 `a = ;` 改为 `a = :p;` 或补全常量。",
      "confidence": 0.86,
      "risk": "low",
      "evidence": [
        { "type": "error_message", "value": "near \";\": syntax error" }
      ],
      "actions": [
        {
          "type": "replace_sql_in_editor",
          "label": "应用修复到编辑器",
          "payload": {
            "sql": "select * from t where a = :p;"
          }
        }
      ]
    }
  ],
  "notes": [
    "未执行 SQL；仅提供修复建议。"
  ],
  "telemetry": {
    "model": "TBD",
    "latencyMs": 0
  }
}
```

### 4.3 Action 类型建议（统一枚举）
- `replace_sql_in_editor`
- `append_sql_to_editor`
- `create_mtj_draft`
- `create_view_sql_draft`
- `create_masking_rule_draft`
- `create_report_appendix_text`

> 关键：**Action 只是“草稿/写入 UI”**，真正“执行/写库/落盘”仍由现有按钮与确认链路完成。

---

## 5. 落盘与审计（Project Config 中新增区）

建议在项目配置中新增（名称可调整）：

```json
{
  "ai_suggestions": [
    {
      "id": "ai-log-0001",
      "time": "2026-04-11T10:00:00Z",
      "capability": "sql_copilot",
      "contextHash": "sha256:...",
      "summary": "…",
      "recommendations": [ "…可裁剪存储…" ],
      "appliedActions": [
        {
          "actionType": "replace_sql_in_editor",
          "time": "2026-04-11T10:01:00Z",
          "userConfirmed": true
        }
      ]
    }
  ]
}
```

并要求：
- 每次“应用 Action”都写入审计日志（你们现有的审计/报告链路可以复用）
- 对敏感场景（脱敏/DDL）记录更严格（操作者、确认弹窗摘要、回滚方式）

---

## 6. 最小 UI 嵌入建议（不破坏现有页面）
1) 各页面右侧增加“AI”折叠面板（默认收起，不打扰）
2) 在关键结果区放置按钮（例如 SQL 报错旁“AI 解释/修复”）
3) 所有 AI 输出统一呈现为：建议卡片（可复制、可一键应用、可查看证据）

