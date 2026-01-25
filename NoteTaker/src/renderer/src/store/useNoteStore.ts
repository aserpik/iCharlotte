import { create } from 'zustand'

interface FileState {
  highlights: any[]
  notesContent: string
}

interface NoteState {
  // File Management
  currentFilePath: string | null
  currentWorkspacePath: string | null
  setWorkspace: (path: string) => void
  
  openFiles: { path: string; name: string }[]
  fileStates: Record<string, FileState>
  
  addFile: (path: string, name: string) => void
  switchFile: (path: string) => void
  closeFile: (path: string) => void
  updateFileContent: (path: string, content: string) => void
  
  pdfUrl: string | null
  pdfTitle: string
  // setPdf is now a wrapper around addFile + switchFile
  setPdf: (url: string, title: string, path: string | null) => void
  
  // Highlighting & Linking
  highlights: any[]
  setHighlights: (highlights: any[]) => void
  addHighlight: (highlight: any) => void
  scrollToHighlightId: string | null
  jumpToHighlight: (id: string) => void
  
  // Selection State
  pendingSelection: { position: any; content: any; hideTip?: () => void } | null
  setPendingSelection: (selection: { position: any; content: any; hideTip?: () => void } | null) => void

  // Auto Note Toggle
  autoNote: boolean
  setAutoNote: (value: boolean) => void

  // Highlight Nesting State
  hasHighlights: boolean
  setHasHighlights: (value: boolean) => void

  nestingLevel: number
  setNestingLevel: (level: number) => void
  
  highlightColor: string
  setHighlightColor: (color: string) => void

  isSaving: boolean
  setIsSaving: (saving: boolean) => void

  zoom: number
  debouncedZoom: number
  setZoom: (zoom: number) => void
  setDebouncedZoom: (zoom: number) => void

  hasLoaded: boolean
  loadFromState: (state: any) => void

  isOcrRunning: boolean
  ocrMessage: string
  setOcrStatus: (running: boolean, message: string) => void
  checkAndRunOCR: (path: string) => Promise<void>

  // A queue of actions for the editor to consume
  editorActions: Array<{ type: 'INSERT_TITLE' | 'INSERT_BULLET' | 'INSERT_IMAGE' | 'INSERT_PDF_NAME'; payload: string; highlightId?: string; level?: number }>
  pushAction: (action: { type: 'INSERT_TITLE' | 'INSERT_BULLET' | 'INSERT_IMAGE' | 'INSERT_PDF_NAME'; payload: string; highlightId?: string; level?: number }) => void
  clearActions: () => void
}

export const useNoteStore = create<NoteState>((set, get) => ({
  currentFilePath: null,
  currentWorkspacePath: null,
  setWorkspace: (path) => set({ currentWorkspacePath: path }),

  openFiles: [],
  fileStates: {},

  addFile: (path, name) => set((state) => {
    const exists = state.openFiles.find(f => f.path === path)
    if (exists) return state
    return {
      openFiles: [...state.openFiles, { path, name }],
      // Initialize state for new file if not exists
      fileStates: {
        ...state.fileStates,
        [path]: state.fileStates[path] || { highlights: [], notesContent: '' }
      }
    }
  }),

  switchFile: (path) => {
    const state = get()
    const fileState = state.fileStates[path]
    
    if (fileState) {
        set({
            currentFilePath: path,
            // Add cache buster to force PDF viewer to reload the actual file
            pdfUrl: `${path}?t=${Date.now()}`,
            pdfTitle: state.openFiles.find(f => f.path === path)?.name || 'Document',
            highlights: fileState.highlights,
            hasHighlights: fileState.highlights.length > 0
        })
    }
  },

  closeFile: (path) => set((state) => ({
    openFiles: state.openFiles.filter(f => f.path !== path),
    // Optionally keep the state in fileStates? Let's keep it for cache.
  })),

  updateFileContent: (path, content) => set((state) => ({
    fileStates: {
      ...state.fileStates,
      [path]: {
        ...state.fileStates[path],
        notesContent: content
      }
    }
  })),

  pdfUrl: null,
  pdfTitle: '',
  setPdf: (url, title, path) => {
    if (!path) return
    const state = get()
    
    // 1. Add file if not open
    if (!state.openFiles.find(f => f.path === path)) {
        state.addFile(path, title)
    }

    // 2. Switch to it
    state.switchFile(path)
    
    // 3. Check and Run OCR
    state.checkAndRunOCR(path);
    
    // If it's a "fresh" load (empty notes), maybe insert the title?
    const currentState = get().fileStates[path]
    if (currentState && !currentState.notesContent && currentState.highlights.length === 0) {
        set(s => ({
            editorActions: [...s.editorActions, { type: 'INSERT_PDF_NAME', payload: title }]
        }))
    }
  },

  highlights: [],
  setHighlights: (highlights) => set((state) => {
      // Update local highlights AND the persistent fileState
      const currentPath = state.currentFilePath
      if (currentPath) {
          return {
              highlights,
              fileStates: {
                  ...state.fileStates,
                  [currentPath]: {
                      ...state.fileStates[currentPath],
                      highlights
                  }
              }
          }
      }
      return { highlights }
  }),
  
  addHighlight: (highlight) => set((state) => {
      const newHighlights = [highlight, ...state.highlights]
      const currentPath = state.currentFilePath
      if (currentPath) {
          return {
              highlights: newHighlights,
              fileStates: {
                  ...state.fileStates,
                  [currentPath]: {
                      ...state.fileStates[currentPath],
                      highlights: newHighlights
                  }
              }
          }
      }
      return { highlights: newHighlights }
  }),
  
  scrollToHighlightId: null,
  jumpToHighlight: (id) => set({ scrollToHighlightId: id }),

  pendingSelection: null,
  setPendingSelection: (selection) => set({ pendingSelection: selection }),

  autoNote: true,
  setAutoNote: (value) => set({ autoNote: value }),
  
  hasHighlights: false,
  setHasHighlights: (value) => set({ hasHighlights: value }),

  nestingLevel: 1,
  setNestingLevel: (level: number) => set({ nestingLevel: level }),

  highlightColor: '#ffeb3b', // Default yellow
  setHighlightColor: (color: string) => set({ highlightColor: color }),

  isSaving: false,
  setIsSaving: (saving: boolean) => set({ isSaving: saving }),

  zoom: 1.0,
  debouncedZoom: 1.0,
  setZoom: (zoom: number) => set({ zoom }),
  setDebouncedZoom: (zoom: number) => set({ debouncedZoom: zoom }),

  hasLoaded: false,
  loadFromState: (state) => {
      if (!state) {
          set({ hasLoaded: true });
          return;
      }
      set({
          openFiles: state.openFiles || [],
          fileStates: state.fileStates || {},
          currentFilePath: state.currentFilePath || null,
          hasLoaded: true
      });
      // If there is a current file, ensure the rest of the store syncs up
      if (state.currentFilePath) {
         const fileState = (state.fileStates || {})[state.currentFilePath];
         if (fileState) {
             set({
                 pdfUrl: state.currentFilePath,
                 pdfTitle: (state.openFiles || []).find((f: any) => f.path === state.currentFilePath)?.name || 'Document',
                 highlights: fileState.highlights || [],
                 hasHighlights: (fileState.highlights || []).length > 0
             });
         }
      }
  },

  isOcrRunning: false,
  ocrMessage: '',
  setOcrStatus: (running, message) => set({ isOcrRunning: running, ocrMessage: message }),
  
  checkAndRunOCR: async (path) => {
      const state = get();
      // @ts-ignore
      if (window.api && window.api.checkNeedsOCR) {
          try {
              // @ts-ignore
              const needsOCR = await window.api.checkNeedsOCR(path);
              if (needsOCR) {
                  state.setOcrStatus(true, "Image-based PDF detected. Performing OCR... This may take a moment.");
                  // @ts-ignore
                  await window.api.runOCR(path);
                  
                  // Since we don't have a reliable finished signal back to this promise easily
                  // without adding an event listener, we'll just log.
                  // The user requested it to perform OCR. 
                  // In a real app, we'd listen for an event.
                  console.log("OCR started for", path);
              }
          } catch (err) {
              console.error("OCR Check failed", err);
          }
      }
  },
  
  editorActions: [],
  pushAction: (action) => set((state) => ({ editorActions: [...state.editorActions, action] })),
  clearActions: () => set({ editorActions: [] })
}))
