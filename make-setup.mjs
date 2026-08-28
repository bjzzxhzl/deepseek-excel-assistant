// 生成自解压安装 exe（调用 Windows 自带 IExpress）
// 中文路径会导致 IExpress 读 SED 失败，故先复制到纯英文临时目录再打包
import { readdirSync, statSync, writeFileSync, mkdirSync, copyFileSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { execFileSync } from 'node:child_process';

const dir = import.meta.dirname;
const staging = join(tmpdir(), 'dsx-setup');
rmSync(staging, { recursive: true, force: true });
mkdirSync(staging, { recursive: true });

const exclude = new Set(['test', 'setup.sed', 'DeepSeekExcelAssistantSetup.exe', 'ExcelAISetup.exe', '分发部署.md', '启动服务器.bat', '卸载.bat', '安装注册.bat', 'make-setup.mjs']);
const files = readdirSync(dir)
  .filter(f => {
    if (f.startsWith('~$')) return false; // 排除 Excel 锁文件
    if (exclude.has(f)) return false;
    try { return statSync(join(dir, f)).isFile(); } catch (e) { return false; }
  })
  .sort();

for (const f of files) copyFileSync(join(dir, f), join(staging, f));

const sed = [
  '[Version]', 'Class=IEXPRESS', 'SEDVersion=3',
  '[Options]',
  'PackagePurpose=InstallApp',
  'ShowInstallProgramWindow=1',
  'HideExtractAnimation=1',
  'UseLongFileName=1',
  'InsideCompressed=0',
  'CAB_FixedSize=0',
  'CAB_ResvCodeSigning=0',
  'RebootMode=N',
  'InstallPrompt=',
  'DisplayLicense=',
  'FinishMessage=Installation complete. Launch the desktop shortcut to start.',
  'TargetName=ExcelAISetup.exe',
  'FriendlyName=ExcelAI',
  'AppLaunched=install.bat',
  'PostInstallCmd=<None>',
  'AdminQuietInstCmd=',
  'UserQuietInstCmd=',
  ...files.map((f, i) => `FILE${i}=${f}`),
  '[Strings]',
  '[SourceFiles]',
  'SourceFiles0=' + staging,
  '[SourceFiles0]',
  ...files
].join('\r\n') + '\r\n';

writeFileSync(join(staging, 'setup.sed'), sed, 'utf8');
console.log('打包文件（' + files.length + ' 个）: ' + files.join(', '));
try {
  execFileSync('C:\\Windows\\System32\\iexpress.exe', ['/N', join(staging, 'setup.sed')], { stdio: 'inherit' });
  copyFileSync(join(staging, 'ExcelAISetup.exe'), join(dir, 'ExcelAISetup.exe'));
  console.log('✓ 安装包已生成到 excel-addin\\ExcelAISetup.exe');
} catch (e) {
  console.log('✗ IExpress 生成失败：' + (e.message || e));
}
