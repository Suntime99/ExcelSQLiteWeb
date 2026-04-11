-- ExcelSQLite 脱敏模块：规则/模板仓库（SQLite）建表脚本
-- 目标：规则集版本化、可复用模板、一键套用、可追溯执行

PRAGMA foreign_keys = ON;
PRAGMA journal_mode = WAL;
PRAGMA synchronous = NORMAL;

-- 1) 规则集（Policy）
-- policy_scope:
--   DB_TABLE  : 面向数据库表（table + column）
--   FILE_SHEET: 面向导入文件（sheet + column）
CREATE TABLE IF NOT EXISTS policy (
  policy_id      TEXT PRIMARY KEY, -- GUID
  name           TEXT NOT NULL,
  namespace      TEXT NOT NULL,
  policy_scope   TEXT NOT NULL DEFAULT 'DB_TABLE',
  version        INTEGER NOT NULL DEFAULT 1,
  status         TEXT NOT NULL DEFAULT 'DRAFT', -- DRAFT/ACTIVE/ARCHIVED
  description    TEXT,
  created_at     TEXT NOT NULL DEFAULT (datetime('now')),
  created_by     TEXT,
  updated_at     TEXT,
  updated_by     TEXT
);

CREATE INDEX IF NOT EXISTS ix_policy_ns_status
  ON policy(namespace, status);

CREATE UNIQUE INDEX IF NOT EXISTS ux_policy_name_version
  ON policy(namespace, name, version);

-- 2) 规则条目（Policy Rule）
-- action:
--   TOKENIZE : 生成token列（推荐企业分析/RAG）
--   MASK    : 生成掩码列（打星/截断）
--   GENERALIZE : 泛化（如地址保留省市）
--   NULLIFY : 置空
-- fail_strategy:
--   KEEP_ORIGINAL / SET_NULL / RECORD_ERROR
CREATE TABLE IF NOT EXISTS policy_rule (
  rule_id        TEXT PRIMARY KEY, -- GUID
  policy_id      TEXT NOT NULL,

  -- 定位目标字段（两类scope任选一类填写）
  source_db      TEXT,             -- 可选：源库标识（多连接时用）
  table_name     TEXT,
  sheet_name     TEXT,
  column_name    TEXT NOT NULL,

  data_type      TEXT,             -- STRING/NUMBER/DATE/BOOL（可选）
  sensitive_type TEXT NOT NULL,    -- PHONE/IDNO/EMAIL/NAME/ADDR...
  action         TEXT NOT NULL,    -- TOKENIZE/MASK/GENERALIZE/NULLIFY

  -- 预处理（JSON，按type解释）
  -- 示例：{"trim":true,"upper":true,"removeSymbols":true,"keepDigitsOnly":true,"countryCode":"86"}
  preprocess_json TEXT,

  -- 动作参数（JSON）
  -- TOKENIZE：{"tokenColumn":"col_token","ns":"P01"}
  -- MASK：{"maskColumn":"col_masked","keepPrefix":3,"keepSuffix":4,"maskChar":"*"}
  -- GENERALIZE：{"outColumn":"addr_gen","level":"province_city"}
  action_json     TEXT,

  fail_strategy   TEXT NOT NULL DEFAULT 'RECORD_ERROR',
  enabled         INTEGER NOT NULL DEFAULT 1,
  priority        INTEGER NOT NULL DEFAULT 100, -- 执行顺序（小优先）
  created_at      TEXT NOT NULL DEFAULT (datetime('now')),
  created_by      TEXT,

  FOREIGN KEY(policy_id) REFERENCES policy(policy_id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS ix_rule_policy_enabled
  ON policy_rule(policy_id, enabled, priority);

CREATE INDEX IF NOT EXISTS ix_rule_target
  ON policy_rule(table_name, column_name);

-- 3) 常用模板库（Template）
-- template_type 与 sensitive_type 一致（PHONE/IDNO/NAME/ADDR...）
CREATE TABLE IF NOT EXISTS template (
  template_id    TEXT PRIMARY KEY, -- GUID
  name           TEXT NOT NULL,
  template_type  TEXT NOT NULL,    -- PHONE/IDNO/NAME/ADDR...
  description    TEXT,

  -- 默认预处理与默认动作参数（JSON）
  default_preprocess_json TEXT,
  default_action          TEXT NOT NULL, -- TOKENIZE/MASK/GENERALIZE/NULLIFY
  default_action_json     TEXT,

  created_at     TEXT NOT NULL DEFAULT (datetime('now')),
  created_by     TEXT
);

CREATE INDEX IF NOT EXISTS ix_template_type
  ON template(template_type);

-- 4) 模板应用记录（可选）：用于“一键套用模板→生成规则”
CREATE TABLE IF NOT EXISTS template_apply_log (
  id           TEXT PRIMARY KEY,
  policy_id    TEXT NOT NULL,
  template_id  TEXT NOT NULL,
  applied_at   TEXT NOT NULL DEFAULT (datetime('now')),
  applied_by   TEXT,
  detail_json  TEXT,
  FOREIGN KEY(policy_id) REFERENCES policy(policy_id) ON DELETE CASCADE,
  FOREIGN KEY(template_id) REFERENCES template(template_id) ON DELETE RESTRICT
);

