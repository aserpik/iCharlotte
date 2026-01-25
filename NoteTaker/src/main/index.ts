import { app, shell, BrowserWindow, ipcMain, dialog } from 'electron'
import { join } from 'path'
import { electronApp, optimizer, is } from '@electron-toolkit/utils'
import { writeFile, readFile, readdir, stat } from 'fs/promises'

function createWindow(): void {
  // Create the browser window.
  const mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    show: false,
    autoHideMenuBar: true,
    ...(process.platform === 'linux' ? { icon: join(__dirname, '../../build/icon.png') } : {}),
    webPreferences: {
      preload: join(__dirname, '../preload/index.js'),
      sandbox: false,
      nodeIntegration: true,
      contextIsolation: true,
      webSecurity: false // Allow loading local resources (file://)
    }
  })

  mainWindow.on('ready-to-show', () => {
    mainWindow.show()
  })

  mainWindow.webContents.setWindowOpenHandler((details) => {
    shell.openExternal(details.url)
    return { action: 'deny' }
  })

  // HMR for renderer base on electron-vite cli.
  // Load the remote URL for development or the local html file for production.
  if (is.dev && process.env['ELECTRON_RENDERER_URL']) {
    mainWindow.loadURL(process.env['ELECTRON_RENDERER_URL'])
  } else {
    mainWindow.loadFile(join(__dirname, '../renderer/index.html'))
  }
}

app.whenReady().then(() => {
  // Set app user model id for windows
  electronApp.setAppUserModelId('com.electron')

  app.on('browser-window-created', (_, window) => {
    optimizer.watchWindowShortcuts(window)
  })

  // IPC Handlers
  ipcMain.handle('save-markdown', async (_, content: string, defaultName: string) => {
    const { canceled, filePath } = await dialog.showSaveDialog({
      defaultPath: defaultName || 'notes.md',
      filters: [{ name: 'Markdown', extensions: ['md'] }]
    })

    if (canceled || !filePath) return false

    await writeFile(filePath, content, 'utf-8')
    return true
  })

  ipcMain.handle('select-directory', async () => {
    const { canceled, filePaths } = await dialog.showOpenDialog({
      properties: ['openDirectory']
    })
    if (canceled) return null
    return filePaths[0]
  })

  ipcMain.handle('select-pdf', async () => {
    const { canceled, filePaths } = await dialog.showOpenDialog({
      properties: ['openFile'],
      filters: [{ name: 'PDF Documents', extensions: ['pdf'] }]
    })
    if (canceled) return null
    return filePaths[0]
  })

  ipcMain.handle('list-directory', async (_, dirPath: string) => {
    try {
      const files = await readdir(dirPath)
      const fileStats = await Promise.all(
        files.map(async (file) => {
          try {
            const stats = await stat(join(dirPath, file))
            return {
              name: file,
              isDirectory: stats.isDirectory(),
              path: join(dirPath, file),
              size: stats.size,
              mtime: stats.mtimeMs
            }
          } catch {
            return null
          }
        })
      )
      return fileStats.filter((f) => f !== null)
    } catch (e) {
      console.error(e)
      return []
    }
  })

  ipcMain.handle('read-file', async (_, path: string) => {
    try {
      return await readFile(path, 'utf-8')
    } catch {
      return null
    }
  })

  ipcMain.handle('read-buffer', async (_, path: string) => {
    try {
      return await readFile(path) // Returns Buffer
    } catch (e) {
      console.error(e)
      return null
    }
  })

  ipcMain.handle('save-app-state', async (_, state: any) => {
    try {
      const userDataPath = app.getPath('userData')
      const statePath = join(userDataPath, 'app-state.json')
      await writeFile(statePath, JSON.stringify(state), 'utf-8')
      return true
    } catch (e) {
      console.error('Failed to save state:', e)
      return false
    }
  })

  ipcMain.handle('load-app-state', async () => {
    try {
      const userDataPath = app.getPath('userData')
      const statePath = join(userDataPath, 'app-state.json')
      const data = await readFile(statePath, 'utf-8')
      return JSON.parse(data)
    } catch (e) {
      // File likely doesn't exist on first run
      return null
    }
  })

  ipcMain.handle('move-file', async (_, oldPath: string, newPath: string) => {
    try {
      const fs = require('fs/promises')
      await fs.rename(oldPath, newPath)
      return true
    } catch (e) {
      console.error(e)
      return false
    }
  })

  createWindow()

  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})
