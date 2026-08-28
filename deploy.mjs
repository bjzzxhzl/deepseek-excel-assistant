// 部署地址切换：同步更新云端与本机清单中的页面/图标 URL（免本地服务器）
// 用法: node deploy.mjs https://你的用户名.github.io/deepseek-excel-assistant
//       node deploy.mjs --local   （还原为本地服务器模式）
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const dir = import.meta.dirname;
const arg = process.argv[2] || '';
if (!arg) {
  console.log('用法: node deploy.mjs <公网HTTPS地址>   或   node deploy.mjs --local');
  process.exit(1);
}
const v = 39;
const base = arg === '--local' ? 'http://localhost:8090' : arg.replace(/\/+$/, '');
let parsed;
try { parsed = new URL(base); }
catch { console.error('地址无效：请输入完整的 https:// URL，或使用 --local'); process.exit(1); }
if (arg !== '--local' && parsed.protocol !== 'https:') {
  console.error('公网部署必须使用 HTTPS 地址。');
  process.exit(1);
}

function updateManifest(fileName) {
  const manifestPath = join(dir, fileName);
  if (!existsSync(manifestPath)) throw new Error('清单不存在：' + fileName);
  let manifest = readFileSync(manifestPath, 'utf8');
  manifest = manifest.replace(/https?:\/\/[^"?\s<]+\/taskpane\.html(?:\?[^"\s<]*)?/g, base + '/taskpane.html?v=' + v);
  for (const size of [16, 32, 64, 80]) {
    const pattern = new RegExp('https?:\\/\\/[^"?\\s<]+\\/icon' + size + '\\.png(?:\\?[^"\\s<]*)?', 'g');
    manifest = manifest.replace(pattern, base + '/icon' + size + '.png?v=' + v);
  }
  manifest = manifest.replace(/<SupportUrl DefaultValue="[^"]*"\/>/, '<SupportUrl DefaultValue="' + base + '/"/>');
  manifest = manifest.replace(/<AppDomain>[^<]*<\/AppDomain>/g, '<AppDomain>' + parsed.origin + '</AppDomain>');
  writeFileSync(manifestPath, manifest, 'utf8');
  console.log('✓ ' + fileName + ' 已指向: ' + base);
}

for (const fileName of ['manifest.xml', 'manifest-standalone.xml']) updateManifest(fileName);
console.log('下一步：把以下文件上传到该地址对应的静态空间（与地址同目录）：');
console.log('  taskpane.html, app.js, app.css, icon16.png, icon32.png, icon64.png, icon80.png, manifest.xml');
console.log('然后：完全重启 Excel（本机安装使用 manifest-standalone.xml，管理员部署使用 manifest.xml）。');
