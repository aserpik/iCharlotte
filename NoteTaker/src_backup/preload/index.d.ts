/// <reference types="vite/client" />
/// <reference types="electron" />

declare interface Window {
  electron: import('@electron-toolkit/preload').ElectronAPI
  api: {
    saveMarkdown: (content: string, defaultName: string) => Promise<boolean>
    selectDirectory: () => Promise<string | null>
    listDirectory: (path: string) => Promise<Array<{ name: string; isDirectory: boolean; path: string }>>
    readFile: (path: string) => Promise<string | null>
    readBuffer: (path: string) => Promise<Buffer>
    writeFile: (path: string, content: string) => Promise<boolean>
  }
}