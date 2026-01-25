import React from 'react'
import { render, screen, act } from '@testing-library/react'
import { vi, describe, it, expect, beforeEach } from 'vitest'
import PDFViewer from './PDFViewer'
import { useNoteStore } from '../store/useNoteStore'

// Mock the store
vi.mock('../store/useNoteStore', () => ({
  useNoteStore: vi.fn()
}))

// Mock the react-pdf-viewer components
let capturedRenderHighlightTarget: any = null;

vi.mock('@react-pdf-viewer/core', () => ({
  Viewer: (props: any) => {
    // Capture the renderHighlightTarget from the highlight plugin
    const highlightPlugin = props.plugins.find((p: any) => p.renderHighlightTarget);
    if (highlightPlugin) {
        capturedRenderHighlightTarget = highlightPlugin.renderHighlightTarget;
    }
    return <div data-testid="pdf-viewer">PDF Viewer</div>;
  },
  Worker: ({ children }: any) => <div>{children}</div>,
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
  highlightPlugin: (props: any) => ({
    renderHighlightTarget: props.renderHighlightTarget,
    renderHighlights: props.renderHighlights,
  }),
  Trigger: {
    TextSelection: 'TextSelection'
  }
}))

describe('PDFViewer Selection Flow', () => {
  beforeEach(() => {
    vi.clearAllMocks()
  })

  it('renders the HighlightLogic when renderHighlightTarget is called', async () => {
    const pushAction = vi.fn();
    const addHighlight = vi.fn();
    
    // @ts-ignore
    useNoteStore.mockReturnValue({
      pdfUrl: 'test.pdf',
      setAutoNote: vi.fn(),
      autoNote: false,
      highlights: [],
      setPdf: vi.fn(),
      pushAction,
      addHighlight
    })

    render(<PDFViewer />)
    
    // Wait for the viewer to render and capture the target
    await screen.findByTestId('pdf-viewer')
    expect(capturedRenderHighlightTarget).toBeDefined()

    // Simulate calling renderHighlightTarget
    const props = {
        selectedText: 'Test Selection',
        selectionRegion: { top: 10, left: 10, width: 100, height: 20 },
        highlightAreas: [],
        toggle: vi.fn()
    }

    render(capturedRenderHighlightTarget(props))

    // Check if the "Add Note" button appears
    expect(screen.getByText('Add Note')).toBeDefined()
  })
})