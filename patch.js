const fs = require('fs');
const p = 'f:/网站制作/网站/网站二/includes/cache_manager.asp';
let c = fs.readFileSync(p, 'utf8');
const marker = "' 初始化统计";
const insert = "' V18: 缓存扩展\r\n<!--#include file=\"cache_v18_ext.asp\"-->\r\n\r\n' 初始化统计";
c = c.replace(marker, insert);
fs.writeFileSync(p, c);
console.log('OK - cache_manager.asp patched');
