import { contextBridge, ipcRenderer, webFrame } from 'electron'
import { electronAPI } from '@electron-toolkit/preload'

// Custom APIs for renderer
const api = {
  saveMarkdown: (content: string, defaultName: string) => ipcRenderer.invoke('save-markdown', content, defaultName),
  selectDirectory: () => ipcRenderer.invoke('select-directory'),
  listDirectory: (path: string) => ipcRenderer.invoke('list-directory', path),
  readFile: (path: string) => ipcRenderer.invoke('read-file', path),
  readBuffer: (path: string) => ipcRenderer.invoke('read-buffer', path),
  writeFile: (path: string, content: string) => ipcRenderer.invoke('write-file', path, content),
  moveFile: (oldPath: string, newPath: string) => ipcRenderer.invoke('move-file', oldPath, newPath)
}

// Use `contextBridge` APIs to expose Electron APIs to
// renderer only if context isolation is enabled, otherwise
// just add to the DOM global.
if (process.contextIsolated) {
  try {
    contextBridge.exposeInMainWorld('electron', electronAPI)
    contextBridge.exposeInMainWorld('api', api)
  } catch (error) {
    console.error(error)
  }
} else {
  // @ts-ignore (define in dts)
  window.electron = electronAPI
  // @ts-ignore (define in dts)
  window.api = api
}