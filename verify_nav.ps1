$r = Invoke-WebRequest 'http://localhost/admin/operation/index.asp' -UseBasicParsing
$checks = @{
    'admin_hamburger' = 'adminHamburger'
    'admin_sidebar' = 'adminSidebar'
    'sidebar_overlay' = 'sidebarOverlay'
    'sidebar_toggle_script' = 'openSidebar'
    'active_link_script' = 'setActiveLink'
}
foreach ($key in $checks.Keys) {
    if ($r.Content -match [regex]::Escape($checks[$key])) {
        Write-Host "OK: $key found"
    } else {
        Write-Host "MISSING: $key not found"
    }
}
