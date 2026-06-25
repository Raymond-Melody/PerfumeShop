# PerfumeShop V17 - SQL Server TCP/IP 协议启用脚本
# 用于启用 MSOLEDBSQL 驱动所需的 TCP/IP 协议
# 用途: 启用 localhost\YOURPERFUME 实例的 TCP/IP 协议
# 以管理员身份运行: PowerShell -ExecutionPolicy Bypass -File enable_tcpip_sqlserver.ps1

param(
    [string]$InstanceName = "YOURPERFUME",
    [int]$TcpPort = 1433,
    [switch]$SkipRestart
)

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  PerfumeShop SQL Server TCP/IP 启用工具" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 检查是否以管理员身份运行
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "[WARNING] 建议以管理员身份运行此脚本" -ForegroundColor Yellow
    Write-Host "  请右键 PowerShell -> 以管理员身份运行" -ForegroundColor Yellow
    Write-Host ""
}

Write-Host "目标实例: $InstanceName" -ForegroundColor Yellow
Write-Host "目标端口: $TcpPort" -ForegroundColor Yellow
Write-Host ""

# 1. 检查 SQL Server 配置管理器是否可用
Write-Host "[Step 1/4] 检查 SQL Server 服务状态..." -ForegroundColor Yellow
$serviceName = "MSSQL`$$InstanceName"
$sqlService = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

if ($null -eq $sqlService) {
    Write-Host "[FAIL] 找不到 SQL Server 实例服务: $serviceName" -ForegroundColor Red
    Write-Host "请确认实例名称是否正确"
    exit 1
}
Write-Host "[OK] SQL Server 服务: $($sqlService.DisplayName) - $($sqlService.Status)" -ForegroundColor Green

# 2. 使用 SQL Server 配置管理器启用 TCP/IP (通过 WMI)
Write-Host "[Step 2/4] 启用 TCP/IP 协议..." -ForegroundColor Yellow

$wmiPath = "ROOT\Microsoft\SqlServer\ComputerManagement15"
$instancePath = "MSSQL`$$InstanceName"

try {
    # 尝试 SQL Server 2017 (ComputerManagement15)
    $tcpProtocol = Get-WmiObject -Namespace $wmiPath -Class "ServerNetworkProtocol" -Filter "InstanceName='$InstanceName' AND ProtocolName='Tcp'" -ErrorAction SilentlyContinue
    
    if ($null -eq $tcpProtocol) {
        # 尝试 SQL Server 2016 (ComputerManagement14)
        $wmiPath = "ROOT\Microsoft\SqlServer\ComputerManagement14"
        $tcpProtocol = Get-WmiObject -Namespace $wmiPath -Class "ServerNetworkProtocol" -Filter "InstanceName='$InstanceName' AND ProtocolName='Tcp'" -ErrorAction SilentlyContinue
    }
    if ($null -eq $tcpProtocol) {
        # 尝试 SQL Server 2014 (ComputerManagement13)
        $wmiPath = "ROOT\Microsoft\SqlServer\ComputerManagement13"
        $tcpProtocol = Get-WmiObject -Namespace $wmiPath -Class "ServerNetworkProtocol" -Filter "InstanceName='$InstanceName' AND ProtocolName='Tcp'" -ErrorAction SilentlyContinue
    }
    if ($null -eq $tcpProtocol) {
        # 尝试 SQL Server 2012 (ComputerManagement12)
        $wmiPath = "ROOT\Microsoft\SqlServer\ComputerManagement12"
        $tcpProtocol = Get-WmiObject -Namespace $wmiPath -Class "ServerNetworkProtocol" -Filter "InstanceName='$InstanceName' AND ProtocolName='Tcp'" -ErrorAction SilentlyContinue
    }
    if ($null -eq $tcpProtocol) {
        Write-Host "[WARNING] 无法通过 WMI 自动配置 TCP/IP 协议" -ForegroundColor Yellow
        Write-Host "请手动启用 TCP/IP:"
        Write-Host "  1. 打开 'SQL Server 配置管理器'" -ForegroundColor Cyan
        Write-Host "  2. 展开 'SQL Server 网络配置'" -ForegroundColor Cyan
        Write-Host "  3. 选择 '$InstanceName 的协议'" -ForegroundColor Cyan
        Write-Host "  4. 右键 'TCP/IP' -> '启用'" -ForegroundColor Cyan
        Write-Host "  5. 右键 'TCP/IP' -> '属性' -> 'IP 地址' 页" -ForegroundColor Cyan
        Write-Host "  6. 在 'IPAll' 下设置 TCP 端口为 $TcpPort" -ForegroundColor Cyan
        Write-Host "  7. 重启 SQL Server 服务" -ForegroundColor Cyan
    } else {
        $tcpProtocol.SetEnable()
        Write-Host "[OK] TCP/IP 协议已启用" -ForegroundColor Green
        
        # 设置 TCP 端口
        $tcpIpProps = Get-WmiObject -Namespace $wmiPath -Class "ServerNetworkProtocolProperty" -Filter "InstanceName='$InstanceName' AND ProtocolName='Tcp' AND IPAddressName='IPAll'"
        if ($null -ne $tcpIpProps) {
            foreach ($prop in $tcpIpProps) {
                if ($prop.PropertyName -eq "TcpPort") {
                    $prop.SetStringValue($TcpPort.ToString())
                    Write-Host "[OK] TCP 端口已设置为: $TcpPort" -ForegroundColor Green
                }
                if ($prop.PropertyName -eq "TcpDynamicPorts") {
                    $prop.SetStringValue("")
                    Write-Host "[OK] 动态端口已禁用" -ForegroundColor Green
                }
            }
        }
    }
} catch {
    Write-Host "[WARNING] WMI 配置失败: $_" -ForegroundColor Yellow
    Write-Host "请参考上方手动配置说明进行操作。" -ForegroundColor Yellow
}

# 3. 添加防火墙规则（如果需要）
Write-Host "[Step 3/4] 检查 Windows 防火墙规则..." -ForegroundColor Yellow
$ruleName = "SQL Server (TCP Port $TcpPort)"
$existingRule = Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue
if ($null -eq $existingRule) {
    try {
        New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Protocol TCP -LocalPort $TcpPort -Action Allow -ErrorAction SilentlyContinue | Out-Null
        Write-Host "[OK] 已添加入站防火墙规则: TCP/$TcpPort" -ForegroundColor Green
    } catch {
        Write-Host "[INFO] 防火墙规则添加跳过（可能需要管理员权限）" -ForegroundColor DarkGray
    }
} else {
    Write-Host "[OK] 防火墙规则已存在" -ForegroundColor Green
}

# 4. 重启 SQL Server 服务
if (-not $SkipRestart) {
    Write-Host "[Step 4/4] 重启 SQL Server 服务..." -ForegroundColor Yellow
    try {
        Restart-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
        Write-Host "[OK] SQL Server 服务已重启" -ForegroundColor Green
        
        # 等待服务完全启动
        Start-Sleep -Seconds 5
        $sqlService.Refresh()
        if ($sqlService.Status -eq "Running") {
            Write-Host "[OK] SQL Server 服务运行正常" -ForegroundColor Green
        } else {
            Write-Host "[WARNING] SQL Server 服务状态: $($sqlService.Status)" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "[WARNING] 服务重启失败: $_" -ForegroundColor Yellow
        Write-Host "请手动重启 SQL Server ($serviceName) 服务" -ForegroundColor Yellow
    }
} else {
    Write-Host "[Step 4/4] 跳过服务重启 (-SkipRestart 参数)" -ForegroundColor DarkGray
    Write-Host "请手动重启 SQL Server ($serviceName) 服务以使变更生效" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  操作完成" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "下一步:" -ForegroundColor Yellow
Write-Host "  1. 确认 SQL Server 已启用 TCP/IP" -ForegroundColor Yellow
Write-Host "  2. 在 config.asp 中设置 FEATURE_MSOLEDBSQL = True" -ForegroundColor Yellow
Write-Host "  3. 重启网站应用程序池" -ForegroundColor Yellow
Write-Host "  4. 访问 /api/system_diag.asp 确认连接状态" -ForegroundColor Yellow
