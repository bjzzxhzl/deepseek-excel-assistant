// Compatibility wrapper. The real icon pipeline resizes the generated
// transparent master and packages the Windows multi-size ICO with PowerShell.
import { spawnSync } from 'node:child_process';
import { join } from 'node:path';

const result = spawnSync(
  'powershell.exe',
  ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', join(import.meta.dirname, 'generate-icon-assets.ps1')],
  { cwd: import.meta.dirname, stdio: 'inherit' },
);

if (result.error) throw result.error;
process.exitCode = result.status ?? 1;
