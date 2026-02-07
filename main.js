const { app, BrowserWindow, ipcMain, shell } = require('electron');
const path = require('path');
const { startServer, stopServer, getLastArtifacts, getDatabaseStats } = require('./server');
const { VERSION } = require('./version');
const https = require('https');

// Suppress macOS CoreText font warnings
app.commandLine.appendSwitch('disable-features', 'CalculateNativeWinOcclusion');
app.commandLine.appendSwitch('disable-font-subpixel-positioning');

let mainWindow = null;
let serverUrl = null;

// Update checker configuration
const GITHUB_REPO = 'eiznem/voapps-tools';

const UPDATE_CHECK_INTERVAL = 24 * 60 * 60 * 1000; // 24 hours

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
      disableHtmlFullscreenWindowResize: true
    },
    title: `VoApps Tools v${VERSION}`,
    show: false,
    backgroundColor: '#ffffff'
  });

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
  });

  if (serverUrl) {
    mainWindow.loadURL(serverUrl);
  }

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.whenReady().then(async () => {
  try {
    const result = await startServer();
    serverUrl = result.url;
    console.log(`Server started at ${serverUrl}`);
    createWindow();
  } catch (error) {
    console.error('Failed to start server:', error);
    app.quit();
  }

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('before-quit', async () => {
  await stopServer();
});

// Update Checker
function checkForUpdates() {
  return new Promise((resolve) => {
    const options = {
      hostname: 'api.github.com',
      path: `/repos/${GITHUB_REPO}/releases/latest`,
      method: 'GET',
      headers: {
        'User-Agent': 'VoApps-Tools',
        'Accept': 'application/vnd.github.v3+json'
      }
    };

    const req = https.request(options, (res) => {
      let data = '';

      res.on('data', (chunk) => {
        data += chunk;
      });

      res.on('end', () => {
        try {
          if (res.statusCode === 200) {
            const release = JSON.parse(data);
            const latestVersion = release.tag_name.replace('v', '');
            const currentVersion = VERSION;

            // Simple version comparison (works for semantic versioning)
            const isNewer = compareVersions(latestVersion, currentVersion) > 0;

            if (isNewer) {
              // Find DMG asset
              const dmgAsset = release.assets.find(asset => 
                asset.name.endsWith('.dmg') || asset.name.endsWith('-arm64.dmg')
              );

              resolve({
                updateAvailable: true,
                latestVersion,
                currentVersion,
                releaseUrl: release.html_url,
                downloadUrl: dmgAsset ? dmgAsset.browser_download_url : release.html_url,
                releaseNotes: release.body,
                releaseName: release.name
              });
            } else {
              resolve({
                updateAvailable: false,
                latestVersion,
                currentVersion
              });
            }
          } else {
            resolve({ updateAvailable: false, currentVersion: VERSION, error: `HTTP ${res.statusCode}` });
          }
        } catch (error) {
          resolve({ updateAvailable: false, currentVersion: VERSION, error: error.message });
        }
      });
    });

    req.on('error', (error) => {
      resolve({ updateAvailable: false, currentVersion: VERSION, error: error.message });
    });

    req.setTimeout(10000, () => {
      req.destroy();
      resolve({ updateAvailable: false, currentVersion: VERSION, error: 'Timeout' });
    });

    req.end();
  });
}

function compareVersions(v1, v2) {
  const parts1 = v1.split('.').map(Number);
  const parts2 = v2.split('.').map(Number);
  
  for (let i = 0; i < Math.max(parts1.length, parts2.length); i++) {
    const num1 = parts1[i] || 0;
    const num2 = parts2[i] || 0;
    
    if (num1 > num2) return 1;
    if (num1 < num2) return -1;
  }
  
  return 0;
}

// IPC Handlers
ipcMain.handle('get-last-artifacts', async () => {
  try {
    const artifacts = getLastArtifacts();
    return { ok: true, artifacts };
  } catch (error) {
    return { ok: false, error: error.message };
  }
});

ipcMain.handle('open-path', async (event, filePath) => {
  try {
    await shell.openPath(filePath);
    return { ok: true };
  } catch (error) {
    return { ok: false, error: error.message };
  }
});

ipcMain.handle('check-for-updates', async () => {
  try {
    const result = await checkForUpdates();
    return { ok: true, ...result };
  } catch (error) {
    return { ok: false, error: error.message };
  }
});

ipcMain.handle('open-update-url', async (event, url) => {
  try {
    await shell.openExternal(url);
    return { ok: true };
  } catch (error) {
    return { ok: false, error: error.message };
  }
});

ipcMain.handle('get-database-stats', async () => {
  try {
    const stats = await getDatabaseStats();
    return { ok: true, stats };
  } catch (error) {
    return { ok: false, error: error.message };
  }
});