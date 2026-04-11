## v101 补丁应用说明（覆盖式更新）

本补丁包是“按原工程目录结构”的覆盖式文件集合，用于把本轮修复/增强一次性同步到你的工程。

### 1）如何应用
1. 备份你的工程目录（或先提交 git）
2. 解压 `v101_single_pack_latest.zip`
3. 将解压出来的文件 **按相同相对路径覆盖** 到你的工程根目录中（例如 `Services/QueryEngine.cs` 覆盖到工程的 `Services/QueryEngine.cs`）
4. 本地执行一次 Release 编译并按《v101 验收清单》回归

### 2）包含的文件（本轮有变更）
- `index.html`
- `app.js`
- `Form1.cs`
- `Services/QueryEngine.cs`
- `Services/SqliteManager.cs`
- `v101_acceptance_checklist.md`
- `v101_release_notes.md`
- `v101_apply_instructions.md`

### 3）建议的本地回归顺序（最短闭环）
1. 打开项目/保存项目 → 导入 SQLite
2. 文件分析（SQLite）→ 一键关系识别
3. 关系识别 → 批量推送（含组合键）→ 多表关联 Step2 校验 → 锁定 SQL
4. 多表查询/多表统计：受控模式 + WITH(CTE) 场景
5. 导出：文件分析/工作表分析/元数据扫描/关系识别

