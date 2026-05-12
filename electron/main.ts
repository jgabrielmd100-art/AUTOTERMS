import { app, BrowserWindow, ipcMain } from 'electron';
import path from 'path';
import { fileURLToPath } from 'url';
import { autoUpdater } from 'electron-updater';
import log from 'electron-log';
import isDev from 'electron-is-dev';
import { spawn } from 'child_process';

// Configure logging
log.transports.file.level = 'info';
autoUpdater.logger = log;

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

let mainWindow: BrowserWindow | null = null;
let serverProcess: any = null;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
    icon: path.join(__dirname, '../public/logo.png'), // Ensure you have a logo.png or similar
    title: 'AutoTermos',
  });

  if (isDev) {
    mainWindow.loadURL('http://localhost:3000');
    mainWindow.webContents.openDevTools();
  } else {
    mainWindow.loadFile(path.join(__dirname, '../dist/index.html'));
  }

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// Start the Express server
function startServer() {
  if (isDev) {
    // In dev, the server is usually started separately via npm run dev:all
    // But we can also spawn it here if needed.
    return;
  }

  // In production, we run the bundled server
  const serverPath = path.join(__dirname, '../dist-server/index.mjs');
  serverProcess = spawn('node', [serverPath], {
    env: { ...process.env, PORT: '3001' }
  });

  serverProcess.stdout.on('data', (data: any) => {
    log.info(`Server: ${data}`);
  });

  serverProcess.stderr.on('data', (data: any) => {
    log.error(`Server Error: ${data}`);
  });
}

app.on('ready', () => {
  createWindow();
  startServer();
  
  if (!isDev) {
    autoUpdater.checkForUpdatesAndNotify();
  }
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
  if (serverProcess) {
    serverProcess.kill();
  }
});

app.on('activate', () => {
  if (mainWindow === null) {
    createWindow();
  }
});

// IPC handlers for Auto-Updater
ipcMain.on('check-for-updates', () => {
  autoUpdater.checkForUpdates();
});

ipcMain.on('restart-app', () => {
  autoUpdater.quitAndInstall();
});

autoUpdater.on('update-available', () => {
  mainWindow?.webContents.send('update_available');
});

autoUpdater.on('update-downloaded', () => {
  mainWindow?.webContents.send('update_downloaded');
});

autoUpdater.on('error', (err) => {
  log.error('Update error:', err);
  mainWindow?.webContents.send('update_error', err.message);
});
