const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  imap: {
    getConfig: () => ipcRenderer.invoke('imap:getConfig'),
    saveConfig: (config) => ipcRenderer.invoke('imap:saveConfig', config),
    fetch: (limit) => ipcRenderer.invoke('imap:fetch', limit),
    fetchOne: (uid) => ipcRenderer.invoke('imap:fetchOne', uid),
    test: (config) => ipcRenderer.invoke('imap:test', config)
  }
});
