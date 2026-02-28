const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  imap: {
    getConfig: () => ipcRenderer.invoke('imap:getConfig'),
    saveConfig: (config) => ipcRenderer.invoke('imap:saveConfig', config),
    fetch: (limit) => ipcRenderer.invoke('imap:fetch', limit),
    fetchMore: (limit, beforeUid) => ipcRenderer.invoke('imap:fetchMore', limit, beforeUid),
    fetchOne: (uid) => ipcRenderer.invoke('imap:fetchOne', uid),
    fetchBodies: (uids) => ipcRenderer.invoke('imap:fetchBodies', uids),
    fetchThreadHeaders: (uids) => ipcRenderer.invoke('imap:fetchThreadHeaders', uids),
    test: (config) => ipcRenderer.invoke('imap:test', config)
  },
  smtp: {
    send: (opts) => ipcRenderer.invoke('smtp:send', opts)
  },
  embeddings: {
    compute: (emails) => ipcRenderer.invoke('embeddings:compute', emails),
    query: (query) => ipcRenderer.invoke('embeddings:query', query),
    pca: (embeddings, emailIds) => ipcRenderer.invoke('embeddings:pca', embeddings, emailIds),
    cluster: (embeddings, emailIds) => ipcRenderer.invoke('embeddings:cluster', embeddings, emailIds),
    promptCluster: (prompt, embeddings, emailIds, threshold) =>
      ipcRenderer.invoke('embeddings:promptCluster', prompt, embeddings, emailIds, threshold),
    promptClusterScored: (prompt, embeddings, emailIds) =>
      ipcRenderer.invoke('embeddings:promptClusterScored', prompt, embeddings, emailIds),
    onProgress: (cb) => {
      ipcRenderer.on('embeddings:progress', (_, p) => cb(p));
    }
  },
  search: {
    hybrid: (query, emails, embeddings) => ipcRenderer.invoke('search:hybrid', query, emails, embeddings)
  }
});
