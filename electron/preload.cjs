const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('desktopApi', {
  chooseDirectory: (title) => ipcRenderer.invoke('choose-directory', title),
  importCloud: (source, destination) => ipcRenderer.invoke('import-cloud', { source, destination }),
  prepareLocalFolders: (folders) => ipcRenderer.invoke('prepare-local-folders', { folders }),
  runWorker: (job, workerPath, pythonPath) => ipcRenderer.invoke('run-worker', { job, workerPath, pythonPath }),
  cancelWorker: () => ipcRenderer.invoke('cancel-worker'),
  onWorkerEvent: (callback) => {
    const listener = (_, event) => callback(event)
    ipcRenderer.on('worker-event', listener)
    return () => ipcRenderer.removeListener('worker-event', listener)
  },
})