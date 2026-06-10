$conn = New-Object -ComObject ADODB.Connection
$conn.Open("Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;")

$settings = @{
    "EnableAlipay" = "1"
    "EnableWechatPay" = "1"
    "EnablePayPal" = "1"
    "EnableCOD" = "1"
    "EnableBankTransfer" = "1"
}

foreach ($key in $settings.Keys) {
    $val = $settings[$key]
    # Check if exists
    $rs = $conn.Execute("SELECT COUNT(*) FROM SiteSettings WHERE SettingKey='$key'")
    $count = $rs.Fields(0).Value
    $rs.Close()
    
    if ($count -eq 0) {
        $conn.Execute("INSERT INTO SiteSettings (SettingKey, SettingValue) VALUES ('$key', '$val')")
        Write-Host "Inserted: $key = $val"
    } else {
        $conn.Execute("UPDATE SiteSettings SET SettingValue='$val' WHERE SettingKey='$key'")
        Write-Host "Updated: $key = $val"
    }
}

Write-Host "`n=== All SiteSettings now ==="
$rs = $conn.Execute("SELECT SettingKey, SettingValue FROM SiteSettings")
while (-not $rs.EOF) {
    Write-Host ("{0} = {1}" -f $rs.Fields(0).Value, $rs.Fields(1).Value)
    $rs.MoveNext()
}
$rs.Close()

$conn.Close()
Write-Host "`nPayment settings fixed!"
