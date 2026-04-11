# 脱敏体系（SQLite/本地库 工具版）落地设计说明

目标：这是一套“工具级实现、企业可迁移”的脱敏体系。当前落地于 SQLite / 本地文件，但设计上保留了迁移到公司大型系统（集中式 Vault/KMS、权限系统、审计系统）的接口与边界。

---

## 1. 总体架构（两类数据源 + 两个独立库）

### 1.1 三类数据库对象

1) **Source DB（源库）**  
用户导入 Excel 后生成的 SQLite（或用户打开的已有 SQLite）。原则：**尽量不改源表**，避免污染与误操作。

2) **Masked DB（脱敏输出库）**  
工具生成的脱敏库（新 db 文件），表结构来自源库，但敏感字段改为 token/masked 产物。原则：**可删可重建**。

3) **Vault DB（独立库）**  
只存：token 映射、审计、namespace 注册等。原则：**强隔离**（未来可直接迁移到企业 Vault/KMS）。

> 补充：规则仓库（Policy Repo）建议也是独立 SQLite（policy.db），与 Vault DB 分开，便于权限边界与备份策略。

---

## 2. Token 规范（A：可读型 Token）

### 2.1 格式

建议格式：

```
{TYPE}_{NS}_{BODY}_{CHK}
```

- TYPE：固定枚举（PHONE/IDNO/EMAIL/NAME/ADDR/BANK/…）
- NS：namespace（项目/工程短码，例如 P01/HR/FIN）
- BODY：主体随机串（推荐 Base32 / Crockford Base32），建议 10~16 位
- CHK：校验位（建议保留，用于快速识别输错/截断）

示例：

```
PHONE_P01_7K3M9D2QH6_5
IDNO_HR_1Z8N0F4C2P_A
```

### 2.2 约束与校验

- Token 必须可快速判断类型与 namespace（用于路由与安全兜底）
- Token 必须具备：**唯一性**（全局 UNIQUE(token)）
- 建议：加入 CHK（校验位）避免 token 在报表里被截断或被误填导致还原出错

---

## 3. 规范化（Normalized Value）规则

同值同 Token 的核心前提：**先规范化，再 fingerprint**。不同类型建议不同规范化策略（可迭代）：

### 3.1 PHONE（手机号）
- 去空格、去符号（`-`、空格、括号）
- 统一国家码（可选：+86）
- 仅保留数字（推荐存规范化结果用于 fingerprint 与加密保存）

### 3.2 IDNO（证件号/身份证）
- 去空格
- 字母统一大写（X）
- 可选：校验位合法性检查（不合法则走 on_error 策略）

### 3.3 EMAIL
- trim
- 域名部分小写
- 本地部分是否小写：按需求决定（工具版可默认不改本地部分）

### 3.4 NAME / ADDR
- trim
- 多空格合并为单空格
- 常见符号统一（后续迭代）

---

## 4. Fingerprint 与 GetOrCreateToken

### 4.1 Fingerprint（推荐 HMAC）

工具版可先用：
- `SHA256(namespace|type|normalized_value|secret)`

企业迁移推荐：
- `HMAC-SHA256(secret, namespace|type|normalized_value)`

> fingerprint 只作为查找 key：避免用明文做索引。

### 4.2 GetOrCreateToken 伪流程

1) `normalized = Normalize(type, raw_value)`
2) `fingerprint = HMAC(secret, namespace|type|normalized)`
3) 查 `token_map`：`SELECT token FROM token_map WHERE namespace=? AND type=? AND fingerprint=?`
4) 若存在：返回 token，并更新 `use_count/last_used_at`
5) 若不存在：
   - 生成 token：`TYPE_NS_BODY_CHK`
   - `enc_value = Encrypt(normalized 或 raw_value)`
   - 插入 `token_map`（依赖 UNIQUE(namespace,type,fingerprint) 与 UNIQUE(token) 防并发重复）
   - 返回 token

---

## 5. Masked DB（脱敏输出库）表结构策略

对源表 `T` 生成脱敏表 `T_masked`（建议表名后缀统一）：

### 5.1 字段策略（推荐默认）
- 非敏感字段：原样复制
- 敏感字段：默认 **不复制原字段**（或管理员模式可保留）
- 生成：
  - `{col}_token`
  - `{col}_masked`（可选：打星/部分保留）
- 增加 lineage：
  - `mask_job_id`
  - `policy_id`
  - `policy_version`
  - `masked_at`

---

## 6. 执行流程（SQLite 工具版建议固定 5 步）

1) **扫描字段**：读取表结构、字段类型、行数、抽样值  
2) **应用规则/模板**：生成“执行计划”（哪些表哪些列、如何处理）  
3) **抽样预览**：每列抽样 N 行，展示 `原值 → token/masked`（不落库）  
4) **生成 Masked DB / 表**：创建 `*_masked` 表结构  
5) **批量写入**：按行读取源表 → 处理敏感字段 → insert 到 masked 表（支持进度回调）

---

## 7. “是否脱敏输出”对查询/报表的影响（路由）

为每个查询/报表增加一个统一配置：

```
OutputMode = Masked | Raw
```

路由规则：
- Masked：数据源指向 Masked DB 的 `{table}_masked`
- Raw：数据源指向 Source DB 原表（需要权限/模式开关）

建议工具版先做两级：
- 普通用户：只能 Masked
- 管理员/授权角色：可 Raw + 可执行还原

---

## 8. 回读还原（DETOKENIZE）模板

### 8.1 输入与识别
- 导入 Excel/CSV → 转临时表进 SQLite
- 识别 token 列：
  - 列名后缀 `_token`；或
  - 值形态符合 `TYPE_NS_BODY_CHK`

### 8.2 还原逻辑
对 token 列批量反查：
1) `token -> enc_value`
2) `Decrypt(enc_value) -> 原值`
3) 输出方式：
   - 生成“内部明文报表”（新文件），或
   - 新增 `*_real` 列

### 8.3 权限与审计（强制）
DETOKENIZE 必须：
- 角色校验
- 填写用途/工单号（reason_ticket），否则拒绝
- 落 `audit_log`

---

## 9. 最小可迁移边界（为什么能迁移到企业）

1) Token 生成逻辑与 fingerprint 设计与企业一致（HMAC/KMS 可替换）  
2) Vault DB 的 token_map/audit_log 可以迁移到集中式 Vault/审计平台  
3) Policy Repo 的 policy/policy_rule/template 是通用“规则配置层”  
4) Source/Masked 与 Vault 物理隔离，权限边界清晰  

---

## 10. 本仓库相关产物

- `vault_db_schema.sql`：Vault DB 建表脚本（token_map + audit_log 等）
- `policy_repo_schema.sql`：规则/模板仓库建表脚本（policy + policy_rule + template 等）

