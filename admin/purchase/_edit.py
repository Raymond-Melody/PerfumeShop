# -*- coding: utf-8 -*-
import os

file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'purchase_orders.asp')

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Edit 1: Add copyOrderFromView JS function before </script>
old_script = '    </script>'
new_script = '''        function copyOrderFromView(orderId) {
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
    </script>'''

if old_script in content:
    content = content.replace(old_script, new_script)
    print("Edit 1 applied (copyOrderFromView function added)")
else:
    print("WARNING: Edit 1 target not found!")

# Edit 2: Add copy button in view-mode header
old_btn = '                    <a href="purchase_orders.asp" class="btn btn-secondary">\n                        <i class="fas fa-arrow-left"></i> 返回列表\n                    </a>'
new_btn = '                    <a href="purchase_orders.asp" class="btn btn-secondary">\n                        <i class="fas fa-arrow-left"></i> 返回列表\n                    </a>\n                    <button type="button" class="btn btn-secondary" style="margin-left:8px;" onclick="copyOrderFromView(<%= viewOrderID %>)">\n                        <i class="fas fa-copy"></i> 复制订单\n                    </button>'

if old_btn in content:
    content = content.replace(old_btn, new_btn)
    print("Edit 2 applied (copy button added)")
else:
    print("WARNING: Edit 2 target not found!")

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("File saved successfully")
