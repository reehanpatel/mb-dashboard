import { execSync } from 'child_process';

console.log('[Postinstall] Running post-installation setup...');

try {
  console.log('[Postinstall] Installing client dependencies...');
  execSync('npm --prefix client install', { stdio: 'inherit' });
} catch (e) {
  console.warn('[Postinstall] Client install warning:', e.message);
}

// Attempt to ensure Python dependencies for PPT converter
const pipCandidates = [
  'python3 -m pip install --user python-pptx Pillow',
  'pip3 install --user python-pptx Pillow',
  'python -m pip install --user python-pptx Pillow',
  'pip install --user python-pptx Pillow'
];

for (const cmd of pipCandidates) {
  try {
    console.log(`[Postinstall] Attempting: ${cmd}`);
    execSync(cmd, { stdio: 'inherit', timeout: 60000 });
    console.log('[Postinstall] Python packages (python-pptx, Pillow) installed successfully.');
    break;
  } catch (e) {
    // Silently continue to next candidate if not available
  }
}
