$file = Join-Path $PSScriptRoot "purchase_orders.asp"
$content = Get-Content $file -Raw -Encoding UTF8

# Get the exact anchor text from the file for Edit 2
# Find the unique line containing "fas fa-arrow-left"
$lines = $content -split "`r`n"
foreach ($line in $lines) {
    if ($line -match "fas fa-arrow-left") {
        $anchorLine = $line
        break
    }
}

# Build old2 and new2 using the exact line found
$old2Btn = @"
                    <a href="purchase_orders.asp" class="btn btn-secondary">
$anchorLine
                    </a>
"@

$copyBtnLine = $anchorLine -replace "arrow-left", "copy" -replace "[^<]+$", "复制订单"

$new2Btn = @"
                    <a href="purchase_orders.asp" class="btn btn-secondary">
$anchorLine
                    </a>
                    <button type="button" class="btn btn-secondary" style="margin-left:8px;" onclick="copyOrderFromView(<%= viewOrderID %>)">
                        <i class="fas fa-copy"></i> 复制订单
                    </button>
"@

if ($content.Contains($old2Btn)) {
    $content = $content.Replace($old2Btn, $new2Btn)
    Write-Host "Edit 2 (copy button): SUCCESS"
} else {
    Write-Host "Edit 2: target not found in file"
}

# Edit 1: Add copyOrderFromView function
$old1 = '    </script>'
$new1 = @'
        function copyOrderFromView(orderId) {
            if (!confirm('确定要复制此订单吗？将创建一个新的草稿订单。')) return;
            var form = document.createElement('form');
            form.method = 'POST';
            form.style.display = 'none';
            var input1 = document.createElement('input');
            input1.name = 'action';
            input1.value = 'copy';
            form.appendChild(input1);
            var input2 = document.createElement('input');
            input2.name = 'purchase_id';
            input2.value = orderId;
            form.appendChild(input2);
            document.body.appendChild(form);
            form.submit();
        }
    </script>
'@

if ($content.Contains($old1)) {
    $content = $content.Replace($old1, $new1)
    Write-Host "Edit 1 (JS function): SUCCESS"
} else {
    Write-Host "Edit 1: target not found"
}

[System.IO.File]::WriteAllText($file, $content, [System.Text.UTF8Encoding]::new($false))
Write-Host "File saved"
