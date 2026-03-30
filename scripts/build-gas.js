const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const rootDir = process.cwd();
const outDir = path.join(rootDir, 'out');
const manifestPath = path.join(rootDir, 'appsscript.json');
const outFilePath = path.join(outDir, 'Code.js');

const gasWrappers = [
  'function SendMonthlyReport() { return globalThis.__SendMonthlyReportImpl(); }',
  'function EnsureThisMonth() { return globalThis.__EnsureThisMonthImpl(); }',
  'function EnsureNextMonth() { return globalThis.__EnsureNextMonthImpl(); }',
  'function UpdateVolunteers() { return globalThis.__UpdateVolunteersImpl(); }',
  'function CopyLatestWaiverToAttendance() { return globalThis.__CopyLatestWaiverToAttendanceImpl(); }',
  // Keep backward compatibility for existing trigger/library calls with the original typo.
  'function CopyLaatestWaiverToAttendance() { return globalThis.__CopyLatestWaiverToAttendanceImpl(); }'
].join('\n');

async function build() {
  fs.rmSync(outDir, { recursive: true, force: true });
  fs.mkdirSync(outDir, { recursive: true });

  const esbuildArgs = [
    'exec',
    '--yes',
    'esbuild',
    '--',
    path.join(rootDir, 'main.ts'),
    '--bundle',
    '--platform=browser',
    '--format=iife',
    '--target=es2017',
    `--outfile=${outFilePath}`
  ];

  const result = process.platform === 'win32'
    ? spawnSync('cmd.exe', ['/d', '/s', '/c', 'npm', ...esbuildArgs], { stdio: 'inherit' })
    : spawnSync('npm', esbuildArgs, { stdio: 'inherit' });

  if (result.error) {
    throw result.error;
  }

  if (result.status !== 0) {
    throw new Error(`esbuild bundling failed (exit ${result.status})`);
  }

  const bundled = fs.readFileSync(outFilePath, 'utf8');
  fs.writeFileSync(outFilePath, `${bundled}\n\n${gasWrappers}\n`, 'utf8');

  fs.copyFileSync(manifestPath, path.join(outDir, 'appsscript.json'));
  console.log('Build complete: out/Code.js and out/appsscript.json');
}

build().catch((error) => {
  console.error('Build failed:', error);
  process.exit(1);
});
