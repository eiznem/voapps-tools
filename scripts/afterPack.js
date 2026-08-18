const { execSync } = require('child_process');
const path = require('path');

// Strip macOS extended attributes (resource forks, quarantine flags, Finder metadata)
// from the packaged app before codesign runs. Without this, binaries downloaded from
// the internet carry xattrs that codesign rejects with "detritus not allowed".
exports.default = async function afterPack(context) {
  if (context.electronPlatformName !== 'darwin') return;
  const appPath = path.join(
    context.appOutDir,
    `${context.packager.appInfo.productFilename}.app`
  );
  console.log(`  • stripping xattrs from ${appPath}`);
  execSync(`xattr -cr "${appPath}"`, { stdio: 'inherit' });
};
