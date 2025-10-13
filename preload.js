const { contextBridge, ipcRenderer } = require('electron');

// Expose protected methods that allow the renderer process to use
// the ipcRenderer without exposing the entire object
contextBridge.exposeInMainWorld('electronAPI', {
  onConnectionInfo: (callback) => ipcRenderer.on('connection-info', callback),
  xmlaRequest: (endpoint, soapBody) => ipcRenderer.invoke('xmla-request', { endpoint, soapBody }),
  onMainLog: (callback) => ipcRenderer.on('main-log', callback),
  openExternal: async (url) => {
    console.log('[Preload] openExternal called with URL:', url);
    try {
      const result = await ipcRenderer.invoke('open-external', url);
      console.log('[Preload] openExternal succeeded');
      return result;
    } catch (error) {
      console.error('[Preload] openExternal failed:', error);
      throw error;
    }
  }
});
