-- 默认模板初始化数据（可直接导入到 policy_repo_schema.sql 对应的库）
-- 目标：提供“开箱即用”的基础 PII 模板（PHONE/IDNO/NAME/ADDR）
-- 说明：
-- 1) template_rule.column_name 采用“常见字段名枚举”的方式匹配（工具版先这么做最稳）。
-- 2) 后续如需支持正则/模糊匹配，可在 schema 增加 match_mode/match_pattern。

BEGIN;

-- Namespace：DEFAULT（也可按你的工程名/任务名替换）
-- Template：PII-基础脱敏（Tokenize）
INSERT OR IGNORE INTO template(id, namespace, name, description, created_at, created_by, updated_at, updated_by)
VALUES (
  'tpl_pii_basic_tokenize_v1',
  'DEFAULT',
  'PII-基础脱敏(Tokenize)',
  '基础 PII 脱敏模板：手机号/证件号/姓名/地址 -> *_token（可选 *_masked）',
  datetime('now'),
  'system',
  datetime('now'),
  'system'
);

-- PHONE
INSERT OR IGNORE INTO template_rule(
  id, template_id, table_name, column_name, data_type, action,
  output_token_col, output_mask_col, keep_raw_col,
  normalize_profile, normalize_params, on_error, enabled, sort_order
)
VALUES
  ('tplr_phone_zh_手机号', 'tpl_pii_basic_tokenize_v1', NULL, '手机号', 'PHONE', 'TOKENIZE', NULL, NULL, 0, 'phone', NULL, 'fail', 1, 10),
  ('tplr_phone_en_phone',  'tpl_pii_basic_tokenize_v1', NULL, 'phone',  'PHONE', 'TOKENIZE', NULL, NULL, 0, 'phone', NULL, 'fail', 1, 11),
  ('tplr_phone_en_mobile', 'tpl_pii_basic_tokenize_v1', NULL, 'mobile', 'PHONE', 'TOKENIZE', NULL, NULL, 0, 'phone', NULL, 'fail', 1, 12),
  ('tplr_phone_zh_联系电话', 'tpl_pii_basic_tokenize_v1', NULL, '联系电话', 'PHONE', 'TOKENIZE', NULL, NULL, 0, 'phone', NULL, 'fail', 1, 13);

-- IDNO
INSERT OR IGNORE INTO template_rule(
  id, template_id, table_name, column_name, data_type, action,
  output_token_col, output_mask_col, keep_raw_col,
  normalize_profile, normalize_params, on_error, enabled, sort_order
)
VALUES
  ('tplr_idno_zh_身份证号', 'tpl_pii_basic_tokenize_v1', NULL, '身份证号', 'IDNO', 'TOKENIZE', NULL, NULL, 0, 'idno', NULL, 'fail', 1, 20),
  ('tplr_idno_zh_身份证',   'tpl_pii_basic_tokenize_v1', NULL, '身份证',   'IDNO', 'TOKENIZE', NULL, NULL, 0, 'idno', NULL, 'fail', 1, 21),
  ('tplr_idno_en_idno',     'tpl_pii_basic_tokenize_v1', NULL, 'idno',     'IDNO', 'TOKENIZE', NULL, NULL, 0, 'idno', NULL, 'fail', 1, 22),
  ('tplr_idno_en_idcard',   'tpl_pii_basic_tokenize_v1', NULL, 'id_card',  'IDNO', 'TOKENIZE', NULL, NULL, 0, 'idno', NULL, 'fail', 1, 23);

-- NAME
INSERT OR IGNORE INTO template_rule(
  id, template_id, table_name, column_name, data_type, action,
  output_token_col, output_mask_col, keep_raw_col,
  normalize_profile, normalize_params, on_error, enabled, sort_order
)
VALUES
  ('tplr_name_zh_姓名',     'tpl_pii_basic_tokenize_v1', NULL, '姓名',     'NAME', 'TOKENIZE', NULL, NULL, 0, 'name', NULL, 'fail', 1, 30),
  ('tplr_name_zh_客户姓名', 'tpl_pii_basic_tokenize_v1', NULL, '客户姓名', 'NAME', 'TOKENIZE', NULL, NULL, 0, 'name', NULL, 'fail', 1, 31),
  ('tplr_name_en_name',     'tpl_pii_basic_tokenize_v1', NULL, 'name',     'NAME', 'TOKENIZE', NULL, NULL, 0, 'name', NULL, 'fail', 1, 32);

-- ADDR
INSERT OR IGNORE INTO template_rule(
  id, template_id, table_name, column_name, data_type, action,
  output_token_col, output_mask_col, keep_raw_col,
  normalize_profile, normalize_params, on_error, enabled, sort_order
)
VALUES
  ('tplr_addr_zh_地址',       'tpl_pii_basic_tokenize_v1', NULL, '地址',       'ADDR', 'TOKENIZE', NULL, NULL, 0, 'addr', NULL, 'fail', 1, 40),
  ('tplr_addr_zh_详细地址',   'tpl_pii_basic_tokenize_v1', NULL, '详细地址',   'ADDR', 'TOKENIZE', NULL, NULL, 0, 'addr', NULL, 'fail', 1, 41),
  ('tplr_addr_zh_收货地址',   'tpl_pii_basic_tokenize_v1', NULL, '收货地址',   'ADDR', 'TOKENIZE', NULL, NULL, 0, 'addr', NULL, 'fail', 1, 42),
  ('tplr_addr_en_address',    'tpl_pii_basic_tokenize_v1', NULL, 'address',    'ADDR', 'TOKENIZE', NULL, NULL, 0, 'addr', NULL, 'fail', 1, 43);

COMMIT;

