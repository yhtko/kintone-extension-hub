import fs from 'fs';
import path from 'path';

const rootDir = process.cwd();
const srcDir = path.join(rootDir, 'src');
const distDir = path.join(rootDir, 'dist');

const EXCLUDE_FILES = new Set([
  'dev-tools.js'
]);

const EXCLUDE_DIRS = new Set([
  '.git',
  'node_modules',
  'dist',
  'keys'
]);

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function resetDir(dirPath) {
  fs.rmSync(dirPath, { recursive: true, force: true });
  fs.mkdirSync(dirPath, { recursive: true });
}

function copyRecursive(src, dest) {
  const stat = fs.statSync(src);

  if (stat.isDirectory()) {
    const dirName = path.basename(src);
    if (EXCLUDE_DIRS.has(dirName)) return;

    ensureDir(dest);
    for (const entry of fs.readdirSync(src)) {
      copyRecursive(
        path.join(src, entry),
        path.join(dest, entry)
      );
    }
    return;
  }

  const fileName = path.basename(src);
  if (EXCLUDE_FILES.has(fileName)) return;

  ensureDir(path.dirname(dest));
  fs.copyFileSync(src, dest);
}

function patchManifest() {
  const manifestPath = path.join(distDir, 'manifest.json');
  const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));

  delete manifest.plugbits_dev_tools;

  if (Array.isArray(manifest.web_accessible_resources)) {
    manifest.web_accessible_resources = manifest.web_accessible_resources
      .map((item) => {
        if (!item || !Array.isArray(item.resources)) return item;
        return {
          ...item,
          resources: item.resources.filter((r) => r !== 'dev-tools.js')
        };
      })
      .filter((item) => Array.isArray(item.resources) && item.resources.length > 0);
  }

  fs.writeFileSync(
    manifestPath,
    JSON.stringify(manifest, null, 2),
    'utf8'
  );
}

if (!fs.existsSync(srcDir)) {
  console.error('src folder not found:', srcDir);
  process.exit(1);
}

resetDir(distDir);
copyRecursive(srcDir, distDir);
patchManifest();

console.log('[OK] build completed');
console.log('src :', srcDir);
console.log('dist:', distDir);