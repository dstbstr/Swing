const fs = require('fs');
const path = require('path');
const src = path.join(process.cwd(), 'appsscript.json');
const destDir = path.join(process.cwd(), 'out');
const dest = path.join(destDir, 'appsscript.json');
try {
  fs.mkdirSync(destDir, { recursive: true });
  fs.copyFileSync(src, dest);
  console.log(`Copied ${src} -> ${dest}`);
} catch (e) {
  console.error('Failed to copy appsscript.json:', e);
  process.exit(1);
}
