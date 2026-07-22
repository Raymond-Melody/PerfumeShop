# V18 Classic ASP 退役清单

> **版本**: V18 → V19 全站迁移  
> **日期**: 2026-07-10  
> **状态**: 待执行  

---

## 一、前置检查清单

在执行退役操作前，必须确认以下事项全部通过：

- [ ] **M6 验证通过**: V19 全站功能验收测试全部通过
- [ ] **数据迁移完成**: `setup/v19_final_migrate.ps1` 执行成功，4 张核心表数据 COUNT 一致
- [ ] **性能基线达标**: V19 页面加载时间 ≤ V18 基线
- [ ] **灰度运行稳定**: V19 已独立运行 ≥ 7 天无 P0/P1 级故障
- [ ] **回滚脚本就绪**: `setup/v19_rollback.ps1` 已测试可用
- [ ] **数据库备份完成**: 退役前最终全量备份已执行

---

## 二、IIS 操作步骤

### 2.1 停止 V18 Classic ASP 虚拟目录

```powershell
# 停止 V18 Classic ASP 应用池
Stop-WebAppPool -Name "ClassicASP-Pool"

# 禁用 V18 虚拟目录（不立即删除，保留 30 天）
Set-WebConfigurationProperty -pspath "IIS:\Sites\Default Web Site\PerfumeShop" `
    -filter "system.webServer/security/access" -name "accessPolicy" -value "None"

# 验证 V18 已不可访问
Invoke-WebRequest -Uri "http://localhost/index.asp" -UseBasicParsing | Select-Object StatusCode
# 预期: 403 或 404
```

### 2.2 验证 V19 独立运行

```powershell
# 确认 V19 ASP.NET Core 应用正在运行
Get-Process -Name "dotnet" | Where-Object { $_.CommandLine -like "*PerfumeShop*" }

# 验证 V19 首页正常
$response = Invoke-WebRequest -Uri "http://localhost/" -UseBasicParsing
$response.StatusCode  # 预期: 200

# 验证 V19 API
Invoke-RestMethod -Uri "http://localhost/api/health" | Select-Object status

# 验证 V19 Admin
$response = Invoke-WebRequest -Uri "http://localhost/admin/" -UseBasicParsing
$response.StatusCode  # 预期: 200
```

---

## 三、代码归档

### 3.1 V18 代码打包

```powershell
# 创建归档目录
$archiveDir = "F:\archive\V18_ClassicASP"
New-Item -ItemType Directory -Path $archiveDir -Force

# 打包 V18 代码（排除 node_modules、.git、bin、obj）
$timestamp = Get-Date -Format "yyyyMMdd"
$archiveFile = Join-Path $archiveDir "V18_ClassicASP_${timestamp}.zip"

Compress-Archive -Path "F:\网站制作\网站\网站二\*" `
    -DestinationPath $archiveFile `
    -Force

Write-Host "归档完成: $archiveFile"
```

### 3.2 保留策略

| 项目 | 保留期 | 存储位置 | 负责人 |
|------|--------|----------|--------|
| V18 代码压缩包 | **90 天** | `F:\archive\V18_ClassicASP\` | 运维团队 |
| V18 数据库备份 | **180 天** | `F:\archive\V18_ClassicASP\db\` | DBA |
| V18 运行日志 | **60 天** | `F:\archive\V18_ClassicASP\logs\` | 运维团队 |

> 90 天后如无回滚需求，可删除 V18 代码归档。

---

## 四、数据库备份

### 4.1 退役前最终全量备份

```powershell
# 执行最终全量备份
.\database\backup_database.ps1 `
    -BackupDir "F:\archive\V18_ClassicASP\db" `
    -RetentionDays 180 `
    -ServerInstance "localhost\YOURPERFUME" `
    -DatabaseName "PerfumeShop"
```

### 4.2 备份验证

```powershell
# 验证备份文件完整性
sqlcmd -S "localhost\YOURPERFUME" -E -C -Q "RESTORE VERIFYONLY FROM DISK = N'F:\archive\V18_ClassicASP\db\PerfumeShop_XXXXXXXX_XXXXXX.bak'"
```

---

## 五、Flask AI 服务关闭

V18 依赖的 Python Flask AI 服务不再需要：

```powershell
# 查找并停止 ai-service 进程
Get-Process -Name "python" | Where-Object {
    $_.Path -like "*ai-service*" -or $_.CommandLine -like "*app.py*"
} | Stop-Process -Force

# 验证已停止
Get-Process -Name "python" -ErrorAction SilentlyContinue

# 如通过 nssm 注册为 Windows 服务
nssm stop "PerfumeShop-AI-Service"
nssm remove "PerfumeShop-AI-Service" confirm
```

> **注意**: V19 使用 ASP.NET Core 内置的 AI 集成模块，无需独立 Flask 服务。

---

## 六、运维文档更新

更新以下文档引用，将 V18 相关引用替换为 V19：

| 文档 | 更新内容 |
|------|----------|
| `docs/V19_使用手册.md` | 确认为当前主文档 |
| `docs/V19_完整归档.md` | 确认包含全部 V19 架构说明 |
| `docs/V18.0_使用手册.md` | 标记为 `[已废弃]`，添加指向 V19 的引用 |
| `docs/V18.0_完整归档.md` | 标记为 `[已归档]`，仅供历史参考 |

### 文档标记示例

在 V18 文档顶部添加：

```markdown
> **[已废弃]** 本文档适用于 V18 Classic ASP 版本，已不再维护。  
> 请参阅 [V19 使用手册](./V19_使用手册.md) 获取当前版本文档。
```

---

## 七、CI/CD 更新

### 7.1 GitHub Actions 流水线

更新 `.github/workflows/` 下的流水线配置：

```yaml
# 移除 V18 Classic ASP 构建/部署步骤
# 仅保留 V19 ASP.NET Core 构建/部署

name: V19 Deploy
on:
  push:
    branches: [main]

jobs:
  build-and-deploy:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-dotnet@v4
        with:
          dotnet-version: '8.0.x'
      - run: dotnet restore src/PerfumeShop.sln
      - run: dotnet build src/PerfumeShop.sln -c Release
      - run: dotnet publish src/PerfumeShop.Api -c Release -o ./publish
      # 部署到 IIS ...
```

### 7.2 清理 V18 相关流水线

- 删除 `deploy-v18-classic.yml`（如存在）
- 删除 V18 相关的部署脚本和密钥配置
- 更新仓库 README 中的构建状态徽章

---

## 八、DNS / 监控更新

### 8.1 DNS 检查

```powershell
# 确认域名解析指向 V19 服务器
nslookup yourperfume.com
# 确保 A 记录指向 V19 服务器 IP
```

### 8.2 监控配置更新

| 监控项 | V18 (移除) | V19 (保留/新增) |
|--------|-----------|----------------|
| 健康检查 | `/api/health_check.asp` | `/api/health` |
| 首页监控 | `/index.asp` | `/` |
| Admin 监控 | `/admin/login.asp` | `/admin/` |
| 端口监控 | 80 (IIS Classic) | 5000 (API), 5207 (Admin) |
| 进程监控 | `w3wp.exe` (Classic ASP) | `dotnet` (ASP.NET Core) |

### 8.3 告警规则更新

- 移除 V18 Classic ASP 应用池的告警规则
- 添加 V19 ASP.NET Core 进程存活性告警
- 更新 SSL 证书到期告警（确保指向 V19）

---

## 九、回滚期限

### 9.1 回滚策略

- **回滚窗口**: 退役后 **30 天** 内保留一键回滚能力
- **回滚脚本**: `setup/v19_rollback.ps1`
- **回滚前提**: V18 数据库备份可用、V18 代码归档完整

### 9.2 回滚操作

```powershell
# 一键回滚到 V18
.\setup\v19_rollback.ps1
```

### 9.3 回滚到期处理

30 天回滚窗口到期后（**2026-08-09**）：

1. 删除 V18 代码归档（`F:\archive\V18_ClassicASP\*.zip`）
2. 停止 V18 IIS 应用池（如尚未停止）
3. 移除 V18 虚拟目录配置
4. 保留数据库备份至 180 天期满
5. 更新本文档状态为 `[已完成]`

---

## 十、执行签署

| 步骤 | 执行人 | 日期 | 状态 |
|------|--------|------|------|
| 前置检查 | | | |
| IIS 操作 | | | |
| 代码归档 | | | |
| 数据库备份 | | | |
| AI 服务关闭 | | | |
| 文档更新 | | | |
| CI/CD 更新 | | | |
| DNS/监控更新 | | | |
| 回滚验证 | | | |

---

> **审批人**: ________________  
> **执行日期**: ________________  
> **备注**: ________________
