$file = 'f:\网站制作\网站\网站二\admin\purchase\purchase_orders.asp'
$content = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)

# Edit 1: Add copyOrderFromView JS function before </script>
$scriptTag = '    </script>'
$newScript = @'
        function copyOrderFromView(orderId) {
            if (!confirm('\u786e\u5b9a\u8981\u590d\u5236\u6b64\u8ba2\u5355\u5417\uff1f\u5c06\u521b\u5efa\u4e00\u4e2a\u65b0\u7684\u8349\u7a3f\u8ba2\u5355\u3002')) return;
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
$content = $content.Replace($scriptTag, $newScript)

# Edit 2: Add copy button in view-mode header
$oldBtn = '                    <a href="purchase_orders.asp" class="btn btn-secondary">
                        <i class="fas fa-arrow-left"></i> \u8fd4\u56de\u5217\u8868
                    </a>'
$newBtn = '                    <a href="purchase_orders.asp" class="btn btn-secondary">
                        <i class="fas fa-arrow-left"></i> \u8fd4\u56de\u5217\u8868
                    </a>
                    <button type="button" class="btn btn-secondary" style="margin-left:8px;" onclick="copyOrderFromView(<%= viewOrderID %>)">
                        <i class="fas fa-copy"></i> \u590d\u5236\u8ba2\u5355
                    </button>'
$content = $content.Replace($oldBtn, $newBtn)

[System.IO.File]::WriteAllText($file, $content, [System.Text.Encoding]::UTF8)
Write-Host "All edits completed successfully"
