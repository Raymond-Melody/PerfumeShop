# PerfumeShop 数据库权限管理完整指南

## 📋 目录
1. [权限概述](#权限概述)
2. [快速修复](#快速修复)
3. [详细解决方案](#详细解决方案)
4. [故障排除](#故障排除)
5. [最佳实践](#最佳实践)

---

## 权限概述

### 必需的数据库权限

| 权限角色 | 用途 | 是否必需 |
|---------|------|---------|
| `db_ddladmin` | 创建/修改/删除表、索引、视图等数据库对象 | ✅ 必需 |
| `db_datareader` | 读取所有用户表中的数据（SELECT） | ✅ 必需 |
| `db_datawriter` | 插入/更新/删除所有用户表中的数据 | ✅ 必需 |
| `BACKUP DATABASE` | 执行数据库完整备份 | ⚠️ 可选 |
| `BACKUP LOG` | 执行事务日志备份 | ⚠️ 可选 |

### 常见的 IIS 应用池身份

| 身份名称 | 说明 | 使用场景 |
|---------|------|---------|
| `NT AUTHORITY\IUSR` | IIS 默认匿名用户 | IIS 7.0+ 默认配置 |
| `IIS APPPOOL\DefaultAppPool` | IIS 默认应用池身份 | Classic 管道模式 |
| `NT AUTHORITY\NETWORK SERVICE` | Network Service 账户 | 旧版 IIS 配置 |

---

## 快速修复

### 方法 1：一键自动修复（推荐）

1. 在文件资源管理器中导航到：
   ```
   f:\网站制作\网站\网站二\setup\
   ```

2. 找到文件 `grant_ddl_permission.ps1`

3. **右键** → **以管理员身份运行**

4. 等待脚本执行完成（约 5-10 秒）

5. 刷新部署工具页面：http://localhost/setup/deploy.asp

### 方法 2：使用 SSMS 执行 SQL 脚本

1. 打开 **SQL Server Management Studio (SSMS)**

2. 连接到服务器：`localhost\YOURPERFUME`

3. 打开文件：`f:\网站制作\网站\网站二\setup\grant_full_permissions.sql`

4. 点击 **执行** (或按 F5)

5. 查看输出确认所有权限授予成功

6. 刷新部署工具页面验证

### 方法 3：重新运行部署工具

1. 访问：http://localhost/setup/deploy.asp?action=run

2. 点击"开始完整部署"

3. 部署工具会自动尝试授予权限

4. 如果权限授予失败，页面会显示详细的解决指南

---

## 详细解决方案

### 方案 A：使用 PowerShell 脚本

**文件**: `setup/grant_ddl_permission.ps1`

**功能**:
- 自动检测所有常见的 IIS 应用池身份
- 为每个身份授予完整的数据库权限
- 包含详细的进度输出和错误处理
- 自动验证权限授予结果

**执行步骤**:
```powershell
# 以管理员身份打开 PowerShell
# 导航到 setup 目录
cd "f:\网站制作\网站\网站二\setup"

# 执行脚本
.\grant_ddl_permission.ps1
```

**输出示例**:
```
========================================
PerfumeShop 数据库权限自动配置工具
========================================

[*] 当前 Windows 用户: YourUsername
[*] 将配置以下 IIS 身份:
    - NT AUTHORITY\IUSR
    - IIS APPPOOL\DefaultAppPool
    - NT AUTHORITY\NETWORK SERVICE

[*] 正在连接 SQL Server: localhost\YOURPERFUME

✅ 权限授予成功！

已授予的权限:
  • db_ddladmin    - DDL 操作（CREATE/ALTER/DROP 表、索引等）
  • db_datareader  - 数据读取（SELECT 所有用户表）
  • db_datawriter  - 数据写入（INSERT/UPDATE/DELETE）
  • BACKUP DATABASE - 数据库完整备份
  • BACKUP LOG     - 事务日志备份
```

### 方案 B：使用 SSMS 手动授权

**步骤 1：打开 SSMS 并连接**
- 服务器名称：`localhost\YOURPERFUME`
- 身份验证：Windows 身份验证

**步骤 2：切换到 PerfumeShop 数据库**
```sql
USE [PerfumeShop]
GO
```

**步骤 3：为 IIS 身份授予权限**

```sql
-- 为 NT AUTHORITY\IUSR 授予权限
IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = 'NT AUTHORITY\IUSR')
    CREATE USER [NT AUTHORITY\IUSR] FOR LOGIN [NT AUTHORITY\IUSR]
GO

ALTER ROLE db_ddladmin ADD MEMBER [NT AUTHORITY\IUSR]
ALTER ROLE db_datareader ADD MEMBER [NT AUTHORITY\IUSR]
ALTER ROLE db_datawriter ADD MEMBER [NT AUTHORITY\IUSR]
GRANT BACKUP DATABASE TO [NT AUTHORITY\IUSR]
GRANT BACKUP LOG TO [NT AUTHORITY\IUSR]
GO

-- 为 IIS APPPOOL\DefaultAppPool 授予权限
IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = 'IIS APPPOOL\DefaultAppPool')
    CREATE USER [IIS APPPOOL\DefaultAppPool] FOR LOGIN [IIS APPPOOL\DefaultAppPool]
GO

ALTER ROLE db_ddladmin ADD MEMBER [IIS APPPOOL\DefaultAppPool]
ALTER ROLE db_datareader ADD MEMBER [IIS APPPOOL\DefaultAppPool]
ALTER ROLE db_datawriter ADD MEMBER [IIS APPPOOL\DefaultAppPool]
GRANT BACKUP DATABASE TO [IIS APPPOOL\DefaultAppPool]
GRANT BACKUP LOG TO [IIS APPPOOL\DefaultAppPool]
GO

-- 为 NT AUTHORITY\NETWORK SERVICE 授予权限
IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = 'NT AUTHORITY\NETWORK SERVICE')
    CREATE USER [NT AUTHORITY\NETWORK SERVICE] FOR LOGIN [NT AUTHORITY\NETWORK SERVICE]
GO

ALTER ROLE db_ddladmin ADD MEMBER [NT AUTHORITY\NETWORK SERVICE]
ALTER ROLE db_datareader ADD MEMBER [NT AUTHORITY\NETWORK SERVICE]
ALTER ROLE db_datawriter ADD MEMBER [NT AUTHORITY\NETWORK SERVICE]
GRANT BACKUP DATABASE TO [NT AUTHORITY\NETWORK SERVICE]
GRANT BACKUP LOG TO [NT AUTHORITY\NETWORK SERVICE]
GO
```

**步骤 4：验证权限**
```sql
-- 查看所有数据库用户的角色
SELECT 
    dp.name AS UserName,
    dp2.name AS RoleName
FROM sys.database_principals dp
LEFT JOIN sys.database_role_members drm ON dp.principal_id = drm.member_principal_id
LEFT JOIN sys.database_principals dp2 ON drm.role_principal_id = dp2.principal_id
WHERE dp.name IN (
    'NT AUTHORITY\IUSR',
    'IIS APPPOOL\DefaultAppPool',
    'NT AUTHORITY\NETWORK SERVICE'
)
ORDER BY dp.name, dp2.name
```

### 方案 C：使用 T-SQL 命令（高级）

**适用于：需要为自定义应用池身份授予权限**

```sql
USE [PerfumeShop]
GO

-- 替换为你的应用池身份
DECLARE @AppPoolIdentity NVARCHAR(255) = 'IIS APPPOOL\YourCustomAppPool'

-- 创建数据库用户
IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = @AppPoolIdentity)
BEGIN
    CREATE USER [@AppPoolIdentity] FOR LOGIN [@AppPoolIdentity]
    PRINT '数据库用户创建成功'
END
GO

-- 授予角色
ALTER ROLE db_ddladmin ADD MEMBER [@AppPoolIdentity]
ALTER ROLE db_datareader ADD MEMBER [@AppPoolIdentity]
ALTER ROLE db_datawriter ADD MEMBER [@AppPoolIdentity]
GRANT BACKUP DATABASE TO [@AppPoolIdentity]
GRANT BACKUP LOG TO [@AppPoolIdentity]
GO

PRINT '权限授予完成'
```

---

## 故障排除

### 问题 1：部署时报错"拒绝了 CREATE TABLE 权限"

**症状**:
```
✗ LoginAlerts - 创建失败: 在数据库 'PerfumeShop' 中拒绝了 CREATE TABLE 权限。
```

**原因**: 当前 IIS 身份没有 `db_ddladmin` 角色

**解决方案**:
1. 运行 `setup/grant_ddl_permission.ps1`（以管理员身份）
2. 或在 SSMS 中执行：
   ```sql
   USE [PerfumeShop]
   ALTER ROLE db_ddladmin ADD MEMBER [NT AUTHORITY\IUSR]
   ```
3. 刷新部署页面重试

### 问题 2：PowerShell 脚本报错"无法连接到 SQL Server"

**症状**:
```
❌ 执行失败，错误信息:
  Sqlcmd: Error: Microsoft ODBC Driver 17 for SQL Server : 无法打开登录所请求的数据库
```

**原因**: SQL Server 服务未启动

**解决方案**:
1. 按 `Win + R`，输入 `services.msc`
2. 找到 **SQL Server (YOURPERFUME)** 服务
3. 右键 → **启动**
4. 重新运行 PowerShell 脚本

### 问题 3：SSMS 中执行脚本时提示"用户没有执行此操作的权限"

**症状**:
```
Msg 15247, Level 16, State 1, Procedure sp_addrolemember
用户没有执行此操作的权限。
```

**原因**: 当前登录 SSMS 的 Windows 用户没有 SQL Server `sysadmin` 角色

**解决方案**:
1. 确认当前 Windows 用户是 SQL Server 管理员
2. 或者使用 `sa` 账户登录 SSMS（如果已启用）
3. 或者联系数据库管理员授予权限

### 问题 4：权限授予成功但部署仍然失败

**症状**: PowerShell 脚本显示成功，但部署工具仍然报权限错误

**原因**: IIS 应用程序池需要重启才能识别新权限

**解决方案**:
1. 打开命令提示符（管理员）
2. 执行：`iisreset`
3. 或使用 IIS 管理器重启应用程序池
4. 重新运行部署工具

### 问题 5：不确定当前使用的是哪个 IIS 身份

**解决方案**:
1. 访问：http://localhost/setup/verify_permissions.asp
2. 查看"当前身份信息"部分
3. 或创建临时 ASP 文件：
   ```asp
   <% Response.Write "当前登录: " & SUSER_SNAME() %>
   ```

### 问题 6：权限授予后某些表仍然创建失败

**原因**: 表可能已存在但结构不匹配

**解决方案**:
1. 在 SSMS 中执行：
   ```sql
   USE [PerfumeShop]
   SELECT name, create_date FROM sys.tables ORDER BY create_date DESC
   ```
2. 检查失败表的状态
3. 如果需要重建，在部署工具 URL 后添加 `&force=1` 参数

---

## 最佳实践

### 1. 定期验证权限

每月运行一次权限验证：
```
http://localhost/setup/verify_permissions.asp
```

### 2. 使用专用应用池

为 PerfumeShop 创建专用的 IIS 应用池：
1. 打开 IIS 管理器
2. 右键"应用程序池" → "添加应用程序池"
3. 名称：`PerfumeShopPool`
4. .NET CLR 版本：`无托管代码`（Classic ASP）
5. 托管管道模式：`经典`

### 3. 最小权限原则

- 生产环境：仅授予 `db_datareader` 和 `db_datawriter`
- 开发/部署环境：额外授予 `db_ddladmin`
- 备份权限：仅在需要备份功能时授予

### 4. 权限变更日志

记录所有权限变更：
```sql
-- 在 master 数据库中创建审计表
USE [master]
GO
CREATE TABLE PermissionAudit (
    AuditID INT IDENTITY(1,1) PRIMARY KEY,
    ChangeDate DATETIME DEFAULT GETDATE(),
    UserName NVARCHAR(255),
    RoleName NVARCHAR(100),
    Action NVARCHAR(50), -- GRANT/REVOKE
    ChangedBy NVARCHAR(255) DEFAULT SUSER_SNAME()
)
GO
```

### 5. 自动化权限检查

创建定时任务每周检查权限：

**PowerShell 脚本** (`check_permissions.ps1`):
```powershell
$sqlCmd = @"
USE [PerfumeShop]
SELECT 
    dp.name AS UserName,
    dp2.name AS RoleName
FROM sys.database_principals dp
LEFT JOIN sys.database_role_members drm ON dp.principal_id = drm.member_principal_id
LEFT JOIN sys.database_principals dp2 ON drm.role_principal_id = dp2.principal_id
WHERE dp.name LIKE 'NT AUTHORITY%' OR dp.name LIKE 'IIS APPPOOL%'
"@

$output = sqlcmd -S localhost\YOURPERFUME -E -Q $sqlCmd
$output | Out-File "f:\网站制作\网站\网站二\logs\permission_check_$(Get-Date -Format 'yyyyMMdd').txt"
```

### 6. 文档维护

- 更新此指南时，同步更新部署工具中的错误提示
- 保持 PowerShell 脚本和 SQL 脚本的权限列表一致
- 定期测试所有修复方案的有效性

---

## 相关工具

| 工具 | 路径 | 用途 |
|-----|------|------|
| 部署工具 | `setup/deploy.asp` | 完整数据库部署 |
| 权限验证 | `setup/verify_permissions.asp` | 可视化权限检查 |
| PowerShell 脚本 | `setup/grant_ddl_permission.ps1` | 一键授予权限 |
| SQL 脚本 | `setup/grant_full_permissions.sql` | SSMS 手动执行 |
| 备份权限脚本 | `database/grant_backup_permission.sql` | 仅授予备份权限 |

---

## 获取帮助

如果遇到本指南未涵盖的问题：

1. 查看 SQL Server 错误日志
2. 运行环境诊断：http://localhost/setup/env_check.asp
3. 检查 IIS 日志：`C:\inetpub\logs\LogFiles`
4. 查看 Windows 事件查看器 → Windows 日志 → 应用程序

---

**文档版本**: 2.0  
**最后更新**: 2026-06-10  
**适用版本**: PerfumeShop V10.4+
