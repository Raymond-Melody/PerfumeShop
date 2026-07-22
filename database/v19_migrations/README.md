# database/v19_migrations/

V19 版本数据库迁移管理目录，存放里程碑级别的备份、迁移和回滚脚本。

## 目录用途

- 执行里程碑迁移前，先创建带版本号的数据库备份
- 存放每个里程碑的回滚脚本，确保迁移失败时可快速恢复
- 所有脚本纳入 Git 版本管理，可追溯变更历史

## 脚本命名约定

| 类型 | 命名格式 | 示例 |
|------|----------|------|
| 备份脚本 | `backup_before_migration.ps1` | `backup_before_migration.ps1 -Milestone M1` |
| 迁移脚本 | `setup/v19_migrate_m{N}.sql` | `setup/v19_migrate_m1.sql` |
| 回滚脚本 | `rollback_m{N}.sql` | `rollback_m1.sql` |

> 迁移脚本放在 `setup/` 目录（与历史迁移脚本保持一致），回滚脚本放在本目录。

## 执行顺序

### 迁移流程（以 M1 为例）

1. **备份数据库**（必须先执行）
   ```powershell
   .\database\v19_migrations\backup_before_migration.ps1 -Milestone M1
   ```
   输出：`database/backups/PerfumeShop_PreM1_20260710.bak`

2. **执行迁移脚本**
   ```powershell
   sqlcmd -S "localhost\YOURPERFUME" -d PerfumeShop -E -C -i "setup\v19_migrate_m1.sql"
   ```

3. **验证迁移结果** — 检查所有 PRINT 输出，确认无错误

4. **启动应用并验证功能** — 确保 M1 相关功能正常工作

### 回滚流程（M1 验证失败时）

1. **停止所有应用程序**（IIS / Kestrel）
2. **执行回滚脚本**
   ```powershell
   sqlcmd -S "localhost\YOURPERFUME" -d PerfumeShop -E -C -i "database\v19_migrations\rollback_m1.sql"
   ```
3. **验证回滚结果** — 检查 PRINT 输出，确认所有对象已移除
4. **如需完全恢复**，可从 PreM1 备份还原：
   ```sql
   RESTORE DATABASE PerfumeShop FROM DISK = N'path\to\PerfumeShop_PreM1_20260710.bak' WITH REPLACE
   ```

## 幂等性保证

- 迁移脚本：每个语句用 `IF NOT EXISTS` 包裹，重复执行不报错
- 回滚脚本：每个语句用 `IF EXISTS` 包裹，重复执行不报错

## 环境要求

- SQL Server 实例：`localhost\YOURPERFUME`
- 数据库：`PerfumeShop`
- 认证：Windows 集成认证（`-E`）
- PowerShell 5.1+（备份脚本）
- sqlcmd 命令行工具（需在 PATH 中）

## 里程碑列表

| 里程碑 | 迁移脚本 | 回滚脚本 | 主要内容 |
|--------|----------|----------|----------|
| M1 | `setup/v19_migrate_m1.sql` | `rollback_m1.sql` | Orders 优惠券/积分字段、ProductReviews 扩展字段、ReviewImages 表 |
