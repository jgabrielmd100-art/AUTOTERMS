import { contextBridge, ipcRenderer } from 'electron';

console.log('Preload script loading...');

contextBridge.exposeInMainWorld('electronAPI', {
  isElectron: true,
  checkForUpdates: () => ipcRenderer.send('check-for-updates'),
  restartApp: () => ipcRenderer.send('restart-app'),
  onUpdateAvailable: (callback: () => void) => ipcRenderer.on('update_available', callback),
  onUpdateDownloaded: (callback: () => void) => ipcRenderer.on('update_downloaded', callback),
  onUpdateError: (callback: (message: string) => void) => ipcRenderer.on('update_error', (_event, message) => callback(message)),
});
