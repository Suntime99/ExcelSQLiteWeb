-- ExcelSQLite 脱敏模块：Vault 独立库（SQLite）建表脚本
-- 目标：同值同Token、可回读还原（密文存储）、强审计
-- 说明：
-- 1) 本脚本假设由应用层负责加密/解密 enc_value
-- 2) fingerprint 建议为 HMAC-SHA256(secret, namespace|type|normalized_value) 的HEX/Base64
-- 3) token 建议形如：{TYPE}_{NS}_{BODY}_{CHK}，例如 PHONE_P01_7K3M9D2QH6_5

PRAGMA foreign_keys = ON;
PRAGMA journal_mode = WAL;
PRAGMA synchronous = NORMAL;

-- 1) Token 映射表：同值同Token（按 namespace+type+fingerprint 唯一）
CREATE TABLE IF NOT EXISTS token_map (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  namespace     TEXT    NOT NULL,  -- 租户/项目/工作区（建议短码）
  type          TEXT    NOT NULL,  -- PHONE/IDNO/EMAIL/NAME/ADDR...
  fingerprint   TEXT    NOT NULL,  -- 指纹（不要直接存明文）
  token         TEXT    NOT NULL,  -- 可读型Token
  enc_value     BLOB    NOT NULL,  -- 可逆还原：加密后的原值（或规范化值）
  key_version   INTEGER NOT NULL DEFAULT 1, -- 密钥版本（便于轮换）
  created_at    TEXT    NOT NULL DEFAULT (datetime('now')),
  created_by    TEXT,
  policy_id     TEXT,
  policy_version INTEGER,
  extra_json    TEXT
);

CREATE UNIQUE INDEX IF NOT EXISTS ux_token_map_fingerprint
  ON token_map(namespace, type, fingerprint);

CREATE UNIQUE INDEX IF NOT EXISTS ux_token_map_token
  ON token_map(token);

CREATE INDEX IF NOT EXISTS ix_token_map_ns_type
  ON token_map(namespace, type);

-- 2) 审计日志：脱敏/还原/导出/导入等都落审计
-- action 建议枚举：
--   MASK_JOB_CREATE / MASK_JOB_RUN / MASK_JOB_EXPORT
--   DETOKENIZE_IMPORT / DETOKENIZE_RUN / DETOKENIZE_EXPORT
CREATE TABLE IF NOT EXISTS audit_log (
  id             TEXT PRIMARY KEY, -- 建议GUID
  action         TEXT NOT NULL,
  namespace      TEXT NOT NULL,
  operator       TEXT NOT NULL,
  created_at     TEXT NOT NULL DEFAULT (datetime('now')),

  -- 还原类动作强烈建议必填（MASK类可选）
  reason_ticket  TEXT, -- 工单号/用途说明

  -- 输入输出引用（文件hash/表名/报表名等）
  input_ref      TEXT,
  output_ref     TEXT,

  row_count      INTEGER DEFAULT 0,
  success_count  INTEGER DEFAULT 0,
  fail_count     INTEGER DEFAULT 0,

  policy_id      TEXT,
  policy_version INTEGER,

  detail_json    TEXT -- 错误明细、参数快照等
);

CREATE INDEX IF NOT EXISTS ix_audit_log_ns_time
  ON audit_log(namespace, created_at);

CREATE INDEX IF NOT EXISTS ix_audit_log_action_time
  ON audit_log(action, created_at);

-- 3) （可选）Namespace 注册表：便于管理与UI下拉
CREATE TABLE IF NOT EXISTS namespace_registry (
  namespace   TEXT PRIMARY KEY,
  display_name TEXT,
  created_at  TEXT NOT NULL DEFAULT (datetime('now')),
  created_by  TEXT
);

