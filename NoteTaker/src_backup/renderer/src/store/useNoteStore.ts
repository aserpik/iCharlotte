import { create } from 'zustand'

interface NoteState {
  // File Management
  currentFilePath: string | null
  currentWorkspacePath: string | null
  setWorkspace: (path: string) => void
  
  pdfUrl: string | null
  pdfTitle: string
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

  // A queue of actions for the editor to consume
  editorActions: Array<{ type: 'INSERT_TITLE' | 'INSERT_BULLET' | 'INSERT_IMAGE' | 'INSERT_PDF_NAME'; payload: string; highlightId?: string }>
  pushAction: (action: { type: 'INSERT_TITLE' | 'INSERT_BULLET' | 'INSERT_IMAGE' | 'INSERT_PDF_NAME'; payload: string; highlightId?: string }) => void
  clearActions: () => void
}

export const useNoteStore = create<NoteState>((set) => ({
  currentFilePath: null,
  currentWorkspacePath: null,
  setWorkspace: (path) => set({ currentWorkspacePath: path }),

  pdfUrl: null,
  pdfTitle: '',
  setPdf: (url, title, path) => set((state) => ({ 
    pdfUrl: url, 
    pdfTitle: title, 
    currentFilePath: path,
    highlights: [], // Clear highlights for the new PDF
    editorActions: [...state.editorActions, { type: 'INSERT_PDF_NAME', payload: title }],
    hasHighlights: false
  })),

  highlights: [],
  setHighlights: (highlights) => set({ highlights }),
  addHighlight: (highlight) => set((state) => ({ highlights: [highlight, ...state.highlights] })),
  
  scrollToHighlightId: null,
  jumpToHighlight: (id) => set({ scrollToHighlightId: id }),

  pendingSelection: null,
  setPendingSelection: (selection) => set({ pendingSelection: selection }),

  autoNote: false,
  setAutoNote: (value) => set({ autoNote: value }),
  
  hasHighlights: false,
  setHasHighlights: (value) => set({ hasHighlights: value }),
  
  editorActions: [],
  pushAction: (action) => set((state) => ({ editorActions: [...state.editorActions, action] })),
  clearActions: () => set({ editorActions: [] })
}))
