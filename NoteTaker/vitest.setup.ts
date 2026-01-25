import { vi } from 'vitest'

// Mock window.api
Object.defineProperty(window, 'api', {
  value: {
    saveMarkdown: vi.fn(),
    readFile: vi.fn(),
    writeFile: vi.fn(),
    selectDirectory: vi.fn(),
    listDirectory: vi.fn()
  },
  writable: true
})

// Mock ResizeObserver which is used by react-resizable-panels or TipTap
class ResizeObserver {
  observe() {}
  unobserve() {}
  disconnect() {}
}
window.ResizeObserver = ResizeObserver

// Mock getClientRects for ProseMirror/TipTap
Element.prototype.getClientRects = vi.fn(() => ({
  length: 0,
  item: () => null,
  [Symbol.iterator]: function* () {}
})) as any

Range.prototype.getClientRects = vi.fn(() => ({
  length: 0,
  item: () => null,
  [Symbol.iterator]: function* () {}
})) as any

Range.prototype.getBoundingClientRect = vi.fn(() => ({
  width: 0,
  height: 0,
  top: 0,
  left: 0,
  bottom: 0,
  right: 0,
})) as any
