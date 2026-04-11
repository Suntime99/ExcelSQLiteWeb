-- 规则/模板仓库（Policy & Template Repo）建表脚本
-- 可以单独一个 SQLite（policy.db），也可以合并进 Vault DB（不推荐合并：权限边界更难做）

PRAGMA foreign_keys = ON;
PRAGMA journal_mode = WAL;
PRAGMA synchronous = NORMAL;

-- 1) 规则集（Policy）元数据
CREATE TABLE IF NOT EXISTS policy (
  id                TEXT PRIMARY KEY,           -- GUID
  namespace         TEXT NOT NULL,              -- P01/HR/...
  name              TEXT NOT NULL,
  description       TEXT,
  current_version   INTEGER NOT NULL DEFAULT 1,
  status            TEXT NOT NULL DEFAULT 'draft', -- draft/published/archived
  created_at        TEXT NOT NULL,
  created_by        TEXT,
  updated_at        TEXT,
  updated_by        TEXT
);

CREATE INDEX IF NOT EXISTS idx_policy_ns ON policy(namespace);
CREATE INDEX IF NOT EXISTS idx_policy_status ON policy(status);

-- 2) 规则版本（可选但建议：保证可追溯）
CREATE TABLE IF NOT EXISTS policy_version (
  id                TEXT PRIMARY KEY,           -- GUID
  policy_id         TEXT NOT NULL,
  version           INTEGER NOT NULL,
  note              TEXT,                       -- 发布说明
  created_at        TEXT NOT NULL,
  created_by        TEXT,
  UNIQUE(policy_id, version),
  FOREIGN KEY(policy_id) REFERENCES policy(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_policy_version_policy ON policy_version(policy_id, version);

-- 3) 规则条目（Policy Rule）
-- 一条规则通常对应：某表某字段，用哪种 type，做 token/mask，规范化策略，失败策略，输出列策略等
CREATE TABLE IF NOT EXISTS policy_rule (
  id                TEXT PRIMARY KEY,           -- GUID
  policy_id         TEXT NOT NULL,
  policy_version    INTEGER NOT NULL,

  table_name        TEXT,                       -- 可空：表示“任意表/按字段名匹配”
  column_name       TEXT NOT NULL,
  column_alias      TEXT,                       -- 可选：显示名

  data_type         TEXT NOT NULL,              -- PHONE/IDNO/EMAIL/NAME/ADDR/...
  action            TEXT NOT NULL,              -- TOKENIZE / MASK / DROP / PASS

  -- 输出列策略（建议统一后缀）
  output_token_col  TEXT,                       -- 默认：{col}_token
  output_mask_col   TEXT,                       -- 默认：{col}_masked
  keep_raw_col      INTEGER NOT NULL DEFAULT 0, -- 0=不保留原列（推荐）；1=保留（仅管理员模式可用）

  -- 规范化/预处理
  normalize_profile TEXT NOT NULL DEFAULT 'default', -- phone/idno/email/name/addr/default
  normalize_params  TEXT,                       -- JSON（自定义参数）

  -- 容错策略
  on_error          TEXT NOT NULL DEFAULT 'fail', -- fail/skip/pass/empty

  enabled           INTEGER NOT NULL DEFAULT 1,
  sort_order        INTEGER NOT NULL DEFAULT 0,

  created_at        TEXT NOT NULL,
  created_by        TEXT,

  FOREIGN KEY(policy_id) REFERENCES policy(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_policy_rule_policy ON policy_rule(policy_id, policy_version);
CREATE INDEX IF NOT EXISTS idx_policy_rule_table  ON policy_rule(table_name);
CREATE INDEX IF NOT EXISTS idx_policy_rule_column ON policy_rule(column_name);
CREATE INDEX IF NOT EXISTS idx_policy_rule_type   ON policy_rule(data_type);

-- 4) 模板（Template）
-- 模板 = 常用规则集合，可一键套用到某个 policy/version 或某个表
CREATE TABLE IF NOT EXISTS template (
  id                TEXT PRIMARY KEY,           -- GUID
  namespace         TEXT NOT NULL,
  name              TEXT NOT NULL,
  description       TEXT,
  created_at        TEXT NOT NULL,
  created_by        TEXT,
  updated_at        TEXT,
  updated_by        TEXT,
  UNIQUE(namespace, name)
);

CREATE INDEX IF NOT EXISTS idx_template_ns ON template(namespace);

-- 5) 模板条目（Template Rule）
CREATE TABLE IF NOT EXISTS template_rule (
  id                TEXT PRIMARY KEY,           -- GUID
  template_id       TEXT NOT NULL,

  table_name        TEXT,                       -- 可空：表示“任意表/按字段名匹配”
  column_name       TEXT NOT NULL,
  data_type         TEXT NOT NULL,              -- PHONE/IDNO/NAME/ADDR/...
  action            TEXT NOT NULL,              -- TOKENIZE / MASK / DROP / PASS

  output_token_col  TEXT,
  output_mask_col   TEXT,
  keep_raw_col      INTEGER NOT NULL DEFAULT 0,

  normalize_profile TEXT NOT NULL DEFAULT 'default',
  normalize_params  TEXT,
  on_error          TEXT NOT NULL DEFAULT 'fail',
  enabled           INTEGER NOT NULL DEFAULT 1,
  sort_order        INTEGER NOT NULL DEFAULT 0,

  FOREIGN KEY(template_id) REFERENCES template(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_template_rule_tpl ON template_rule(template_id);
CREATE INDEX IF NOT EXISTS idx_template_rule_col ON template_rule(column_name);

-- 6) 模板应用日志（Template Apply Log）
CREATE TABLE IF NOT EXISTS template_apply_log (
  id                TEXT PRIMARY KEY,           -- GUID
  namespace         TEXT NOT NULL,
  template_id       TEXT NOT NULL,
  policy_id         TEXT,                       -- 应用到哪个 policy
  policy_version    INTEGER,
  table_name        TEXT,
  applied_rule_cnt  INTEGER,
  operator          TEXT,
  created_at        TEXT NOT NULL,
  detail_json       TEXT,
  FOREIGN KEY(template_id) REFERENCES template(id)
);

CREATE INDEX IF NOT EXISTS idx_tpl_apply_ns_time ON template_apply_log(namespace, created_at);

