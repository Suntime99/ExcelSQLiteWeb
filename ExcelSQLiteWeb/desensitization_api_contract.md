# 脱敏模块消息协议（index.html ↔ Form1.cs）

目标：固定一套最小、稳定、可扩展的消息协议，避免 UI 与 C# 两边“改名/字段不一致”导致隐蔽故障。

## 1. 总体消息格式

### 1.1 Web → Host（JS 发给 C#）

```json
{
  "action": "string",
  "requestId": "string",
  "payload": { }
}
```

约定：
- `action`：必填，动作名（见第2节）
- `requestId`：建议必填，用于并发/追踪（前端生成 `xxx_${Date.now()}_${rand}`）
- `payload`：可空对象

### 1.2 Host → Web（C# 回给 JS）

```json
{
  "action": "string",
  "requestId": "string",
  "ok": true,
  "message": "string",
  "data": { }
}
```

约定：
- `ok=false` 必须同时给出 `message` 与 `errorCode`
- 推荐在 `data` 里回传统计信息（写入行数、耗时、生成库路径等）

---

## 2. Action 清单（第一阶段最小闭环）

### 2.1 Vault / Policy Repo 初始化

#### 2.1.1 `initVaultDb`
用途：创建/打开 vault.db，并执行 `vault_db_schema.sql`（如未初始化）。

Web → Host：
```json
{ "action":"initVaultDb", "requestId":"...", "payload": { "vaultDbPath":"C:/.../vault.db" } }
```

Host → Web：
```json
{ "action":"vaultDbReady", "requestId":"...", "ok":true, "data":{ "vaultDbPath":"...", "schemaVersion":1 } }
```

#### 2.1.2 `initPolicyRepoDb`
用途：创建/打开 policy.db，并执行 `policy_repo_schema.sql`。

Web → Host：
```json
{ "action":"initPolicyRepoDb", "requestId":"...", "payload": { "policyDbPath":"C:/.../policy.db" } }
```

Host → Web：
```json
{ "action":"policyRepoReady", "requestId":"...", "ok":true, "data":{ "policyDbPath":"..." } }
```

#### 2.1.3 `seedDefaultTemplates`
用途：导入默认模板（执行 `policy_repo_seed_default_templates.sql`）。

Web → Host：
```json
{ "action":"seedDefaultTemplates", "requestId":"...", "payload": { "namespace":"DEFAULT" } }
```

Host → Web：
```json
{ "action":"defaultTemplatesSeeded", "requestId":"...", "ok":true, "data":{ "templateCount":1, "ruleCount":12 } }
```

---

### 2.2 脱敏执行（Mask）

#### 2.2.1 `previewMaskSample`
用途：抽样预览（不落库）：展示 `raw -> token/masked`。

Web → Host：
```json
{
  "action":"previewMaskSample",
  "requestId":"...",
  "payload":{
    "namespace":"P01",
    "policyId":"...",
    "policyVersion":1,
    "sourceDbPath":"C:/.../source.db",
    "sourceTable":"Main",
    "limit":50
  }
}
```

Host → Web：
```json
{
  "action":"maskSamplePreviewed",
  "requestId":"...",
  "ok":true,
  "data":{
    "rows":[
      { "手机号":"138****0000", "手机号_token":"PHONE_P01_...." }
    ]
  }
}
```

#### 2.2.2 `executeMaskJob`
用途：生成 Masked DB，并批量写入 `{table}_masked`。

Web → Host：
```json
{
  "action":"executeMaskJob",
  "requestId":"...",
  "payload":{
    "namespace":"P01",
    "operator":"admin",
    "policyId":"...",
    "policyVersion":1,
    "sourceDbPath":"C:/.../source.db",
    "maskedDbPath":"C:/.../masked_P01_20260408.db",
    "plans":[
      {
        "sourceTable":"Main",
        "targetTable":"Main_masked",
        "rules":[
          { "columnName":"手机号", "dataType":"PHONE", "action":"TOKENIZE", "keepRawCol":0 }
        ]
      }
    ]
  }
}
```

Host → Web：
```json
{ "action":"maskJobProgress", "requestId":"...", "ok":true, "data":{ "percent":35, "processed":35000 } }
```

结束回包：
```json
{
  "action":"maskJobCompleted",
  "requestId":"...",
  "ok":true,
  "data":{
    "maskedDbPath":"...",
    "tables":[ { "sourceTable":"Main", "targetTable":"Main_masked", "rowCount":123456 } ],
    "auditLogId":"..."
  }
}
```

---

### 2.3 回读还原（Detokenize）

#### 2.3.1 `detokenizePreview`
用途：对导入的临时表/结果集做预览还原（不导出文件）。

Web → Host：
```json
{
  "action":"detokenizePreview",
  "requestId":"...",
  "payload":{
    "namespace":"P01",
    "operator":"admin",
    "role":"admin",
    "reasonTicket":"INC-20260408-001",
    "input":{
      "dbPath":"C:/.../work.db",
      "table":"ImportedReport"
    },
    "tokenColumns":["手机号_token"],
    "limit":50
  }
}
```

Host → Web：
```json
{
  "action":"detokenizePreviewed",
  "requestId":"...",
  "ok":true,
  "data":{
    "columnsAdded":["手机号_real"],
    "rows":[
      { "手机号_token":"PHONE_P01_....", "手机号_real":"13800138000" }
    ]
  }
}
```

#### 2.3.2 `detokenizeExport`
用途：生成“内部明文报表”（新文件）或写回到新表。

Web → Host：
```json
{
  "action":"detokenizeExport",
  "requestId":"...",
  "payload":{
    "namespace":"P01",
    "operator":"admin",
    "role":"admin",
    "reasonTicket":"INC-20260408-001",
    "inputRef":"fileHash:xxxx",
    "mode":"xlsx",
    "input":{
      "dbPath":"C:/.../work.db",
      "table":"ImportedReport"
    },
    "tokenColumns":["手机号_token"]
  }
}
```

Host → Web：
```json
{ "action":"detokenizeExported", "requestId":"...", "ok":true, "data":{ "outputPath":"C:/.../detokenized.xlsx", "auditLogId":"..." } }
```

---

## 3. 错误码（建议第一版就固定）

| errorCode | 含义 |
|---|---|
| E_BAD_REQUEST | 参数缺失/格式错误 |
| E_NOT_READY | Vault/PolicyRepo 未初始化 |
| E_UNAUTHORIZED | 未授权（如非 admin 却请求 DETOKENIZE） |
| E_REASON_REQUIRED | 未提供 reasonTicket |
| E_TOKEN_INVALID | token 格式不合法/校验失败 |
| E_TOKEN_NOT_FOUND | token 在 token_map 中不存在 |
| E_CRYPTO_FAIL | 加解密失败/密钥不可用 |
| E_SQLITE_FAIL | SQLite 执行失败 |
| E_POLICY_NOT_FOUND | policy/version 不存在 |
| E_RULE_INVALID | 规则不可用（字段不存在/类型不支持） |

---

## 4. C# 伪代码（关键函数）

### 4.1 GetOrCreateToken（同值同Token）

```csharp
string GetOrCreateToken(string ns, string type, string raw, PolicyMeta policy) {
  var normalized = Normalize(type, raw);
  if (string.IsNullOrEmpty(normalized)) return ""; // 由 on_error 决定

  var fp = HmacSha256Hex(secret, $"{ns}|{type}|{normalized}");
  var existed = QueryScalar("SELECT token FROM token_map WHERE namespace=? AND type=? AND fingerprint=?", ns, type, fp);
  if (!string.IsNullOrEmpty(existed)) {
    Exec("UPDATE token_map SET use_count=use_count+1, last_used_at=? WHERE token=?", Now(), existed);
    return existed;
  }

  for (int retry=0; retry<5; retry++) {
    var token = BuildReadableToken(type, ns); // TYPE_NS_BODY_CHK
    var enc = Encrypt(normalized);           // 或 Encrypt(raw)
    try {
      Exec(@"INSERT INTO token_map(id,namespace,type,fingerprint,token,enc_value,policy_id,policy_version,created_at,created_by,use_count,last_used_at)
             VALUES(?,?,?,?,?,?,?,?,?,?,1,?)",
           Guid(), ns, type, fp, token, enc, policy.Id, policy.Version, Now(), policy.Operator, Now());
      return token;
    } catch (SqliteConstraintException) {
      // 可能是 token UNIQUE 冲突或 fp UNIQUE 冲突；重新查一次 fp
      existed = QueryScalar("SELECT token FROM token_map WHERE namespace=? AND type=? AND fingerprint=?", ns, type, fp);
      if (!string.IsNullOrEmpty(existed)) return existed;
      continue;
    }
  }
  throw new Exception("E_SQLITE_FAIL: generate token failed");
}
```

### 4.2 Detokenize（token -> 明文）

```csharp
string Detokenize(string token) {
  var enc = QueryBlob("SELECT enc_value FROM token_map WHERE token=?", token);
  if (enc == null) throw new AppException("E_TOKEN_NOT_FOUND", "token 未登记");
  return Decrypt(enc);
}
```

---

## 5. 审计字段建议（第一版就落库）

MASK：
- action=MASK
- namespace/operator/policy_id/policy_version
- input_ref=sourceDbPath:table
- output_ref=maskedDbPath:table_masked
- row_count/col_count

DETOKENIZE：
- action=DETOKENIZE
- namespace/operator/role
- reason_ticket（必填）
- input_ref（文件 hash / 临时表）
- output_ref（输出文件/表）

