import React from 'react'
import { render, screen } from '@testing-library/react'
import { vi, describe, it, expect, beforeEach } from 'vitest'
import PDFViewer from './PDFViewer'
import { useNoteStore } from '../store/useNoteStore'

// Mock the store
vi.mock('../store/useNoteStore', () => ({
  useNoteStore: vi.fn()
}))

// Mock the react-pdf-viewer components
vi.mock('@react-pdf-viewer/core', () => ({
  Viewer: ({ fileUrl }) => <div data-testid="pdf-viewer" data-url={fileUrl}>PDF Viewer</div>,
  Worker: ({ children }) => <div>{children}</div>,
  SpecialZoomLevel: {
    PageWidth: 'PageWidth'
  }
}))

vi.mock('@react-pdf-viewer/zoom', () => ({
  zoomPlugin: () => ({
    zoomTo: vi.fn()
  })
}))

vi.mock('@react-pdf-viewer/highlight', () => ({
  highlightPlugin: () => ({}),
  Trigger: {
    TextSelection: 'TextSelection'
  }
}))

describe('PDFViewer', () => {
  beforeEach(() => {
    vi.clearAllMocks()
  })

  it('renders a message when no PDF is selected', () => {
    // @ts-ignore
    useNoteStore.mockReturnValue({
      pdfUrl: null,
      setAutoNote: vi.fn(),
      autoNote: false,
      highlights: [],
      setPdf: vi.fn()
    })

    render(<PDFViewer />)
    expect(screen.getByText(/Select a PDF/)).toBeDefined()
  })

  it('renders the viewer when a PDF URL is provided', async () => {
    // @ts-ignore
    useNoteStore.mockReturnValue({
      pdfUrl: 'C:/test.pdf',
      setAutoNote: vi.fn(),
      autoNote: false,
      highlights: [],
      setPdf: vi.fn()
    })

    render(<PDFViewer />)
    
    // It should show "Loading PDF data..." initially or the viewer if it's fast
    // Since we use local-resource scheme, it should set pdfBlobUrl immediately
    const viewer = await screen.findByTestId('pdf-viewer')
    expect(viewer).toBeDefined()
    expect(viewer.getAttribute('data-url')).toContain('local-resource://')
  })
})