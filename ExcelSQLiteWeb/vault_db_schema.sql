-- Vault 独立库（Vault DB）建表脚本
-- 目标：本地 SQLite 工具版可落地，未来可迁移到企业 Vault/KMS
-- 规范：同值同Token；Token 可逆还原（enc_value）；强审计（audit_log）

PRAGMA foreign_keys = ON;
PRAGMA journal_mode = WAL;
PRAGMA synchronous = NORMAL;

-- 1) Namespace 注册表（可选，但建议保留：便于多项目隔离与溯源）
CREATE TABLE IF NOT EXISTS namespace_registry (
  namespace          TEXT PRIMARY KEY,          -- 例如：P01 / HR / FIN
  display_name       TEXT,                      -- 例如：项目名称/工程名
  created_at         TEXT NOT NULL,
  created_by         TEXT,
  note               TEXT
);

-- 2) Token 映射表（核心）
-- 同值同Token 的关键：UNIQUE(namespace, type, fingerprint)
-- 还原关键：token -> enc_value（需本地加密；密钥不建议落库）
CREATE TABLE IF NOT EXISTS token_map (
  id                TEXT PRIMARY KEY,           -- GUID
  namespace         TEXT NOT NULL,
  type              TEXT NOT NULL,              -- PHONE/IDNO/EMAIL/NAME/ADDR/BANK/...
  fingerprint        TEXT NOT NULL,             -- HMAC/Hash 后的指纹（不存明文）
  token              TEXT NOT NULL,             -- 例如 PHONE_P01_7K3M9D2QH6_5
  enc_value          BLOB NOT NULL,             -- 原值（或规范化值）加密后存储（可逆还原）
  norm_hint          TEXT,                      -- 可选：仅存“规范化摘要/规则版本”等，不存敏感明文
  policy_id          TEXT,
  policy_version     INTEGER,
  created_at         TEXT NOT NULL,
  created_by         TEXT,
  last_used_at       TEXT,
  use_count          INTEGER NOT NULL DEFAULT 0,
  UNIQUE(namespace, type, fingerprint),
  UNIQUE(token),
  FOREIGN KEY(namespace) REFERENCES namespace_registry(namespace)
);

CREATE INDEX IF NOT EXISTS idx_token_map_ns_type ON token_map(namespace, type);
CREATE INDEX IF NOT EXISTS idx_token_map_token   ON token_map(token);
CREATE INDEX IF NOT EXISTS idx_token_map_fpr     ON token_map(fingerprint);

-- 3) 审计日志（强制要求：DETOKENIZE 必须记录 reason_ticket）
CREATE TABLE IF NOT EXISTS audit_log (
  id                TEXT PRIMARY KEY,           -- GUID
  action            TEXT NOT NULL,              -- MASK / DETOKENIZE / EXPORT / IMPORT / POLICY_PUBLISH / ...
  namespace         TEXT,
  operator          TEXT,                       -- 用户/操作者
  role              TEXT,                       -- admin/user/...
  reason_ticket     TEXT,                       -- 工单号/用途说明（DETOKENIZE 必填）
  input_ref         TEXT,                       -- 例如：file hash / 表名 / 任务ID
  output_ref        TEXT,                       -- 例如：导出文件名 / 生成库路径
  policy_id         TEXT,
  policy_version    INTEGER,
  row_count         INTEGER,
  col_count         INTEGER,
  started_at        TEXT,
  finished_at       TEXT,
  status            TEXT NOT NULL DEFAULT 'OK', -- OK/FAIL/CANCEL
  error_message     TEXT,
  detail_json       TEXT,                       -- 详细上下文（JSON）
  created_at        TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_audit_action_time ON audit_log(action, created_at);
CREATE INDEX IF NOT EXISTS idx_audit_ns_time     ON audit_log(namespace, created_at);
CREATE INDEX IF NOT EXISTS idx_audit_policy      ON audit_log(policy_id, policy_version);

