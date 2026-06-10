const fs = require('fs');
const file = 'f:/网站制作/网站/网站二/admin/purchase/purchase_orders.asp';
let content = fs.readFileSync(file, 'utf8');

// Fix 1: Find and fix garbled confirm message
const pattern1 = "if (!confirm('";
const idx1 = content.indexOf(pattern1);
if (idx1 >= 0) {
    // End marker: ')) return; (single quote, two closing parens for confirm and if)
    const endMarker = "')) return;";
    const endIdx = content.indexOf(endMarker, idx1);
    if (endIdx >= 0) {
        const before = content.substring(0, idx1);
        const after = content.substring(endIdx + endMarker.length);
        const correctText = '确定要复制此订单吗？将创建一个新的草稿订单。';
        content = before + pattern1 + correctText + endMarker + after;
        console.log('Fix 1: confirm message fixed');
    } else {
        console.log('Fix 1: end marker not found');
    }
} else {
    console.log('Fix 1: pattern not found');
}

fs.writeFileSync(file, content, 'utf8');
console.log('All fixes applied');
