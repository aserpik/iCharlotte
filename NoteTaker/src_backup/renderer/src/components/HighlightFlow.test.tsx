import { describe, it, expect, vi, beforeEach } from 'vitest'
import { render, screen, act, fireEvent } from '@testing-library/react'
import PDFViewer from './PDFViewer'
import { useNoteStore } from '../store/useNoteStore'
import React from 'react'

// Mock useNoteStore
vi.mock('../store/useNoteStore', () => ({
    useNoteStore: vi.fn()
}));

// Mock react-pdf-viewer
let capturedRenderHighlightTarget: any = null;

vi.mock('@react-pdf-viewer/core', () => ({
    Worker: ({ children }: any) => <div>{children}</div>,
    Viewer: (props: any) => {
        const highlightPlugin = props.plugins.find((p: any) => p.renderHighlightTarget);
        if (highlightPlugin) {
            capturedRenderHighlightTarget = highlightPlugin.renderHighlightTarget;
        }
        return <div data-testid="pdf-viewer">PDF Viewer</div>;
    },
    SpecialZoomLevel: { PageWidth: 'PageWidth' }
}));

vi.mock('@react-pdf-viewer/highlight', () => ({
    highlightPlugin: (props: any) => ({
        renderHighlightTarget: props.renderHighlightTarget,
        renderHighlights: props.renderHighlights,
    }),
    Trigger: { TextSelection: 'TextSelection' }
}));

vi.mock('@react-pdf-viewer/zoom', () => ({
    zoomPlugin: () => ({ zoomTo: vi.fn() })
}));

describe('Highlight to Note Flow', () => {
    beforeEach(() => {
        vi.clearAllMocks();
        // Reset the store manually if needed, but here we just mock useNoteStore
    });

    it('triggers highlight and note insertion when Add Note is clicked', async () => {
        const pushAction = vi.fn();
        const addHighlight = vi.fn();

        // @ts-ignore
        useNoteStore.mockReturnValue({
            pdfUrl: 'test.pdf',
            autoNote: false, // Manual note for this test
            setAutoNote: vi.fn(),
            highlights: [],
            setPdf: vi.fn(),
            pushAction,
            addHighlight
        })

        render(<PDFViewer />);

        // 1. Capture the target renderer
        await screen.findByTestId('pdf-viewer');
        expect(capturedRenderHighlightTarget).toBeDefined();

        // 2. Simulate selection by calling renderHighlightTarget
        const props = {
            selectedText: 'Test Highlighted Text',
            selectionRegion: { top: 10, left: 10, width: 100, height: 20 },
            highlightAreas: [{ pageIndex: 0, left: 10, top: 10, width: 100, height: 20 }],
            toggle: vi.fn()
        };

        const { getByText } = render(capturedRenderHighlightTarget(props));

        // 3. Click "Add Note"
        const addNoteBtn = getByText('Add Note');
        fireEvent.click(addNoteBtn);

        // 4. Verify actions
        expect(addHighlight).toHaveBeenCalled();
        expect(pushAction).toHaveBeenCalledWith(expect.objectContaining({
            type: 'INSERT_BULLET',
            payload: 'Test Highlighted Text'
        }));
        expect(props.toggle).toHaveBeenCalled();
    });

    it('auto-notes when autoNote is true', async () => {
        const pushAction = vi.fn();
        const addHighlight = vi.fn();

        // @ts-ignore
        useNoteStore.mockReturnValue({
            pdfUrl: 'test.pdf',
            autoNote: true, // Auto note
            setAutoNote: vi.fn(),
            highlights: [],
            setPdf: vi.fn(),
            pushAction,
            addHighlight
        })

        render(<PDFViewer />);

        await screen.findByTestId('pdf-viewer');

        const props = {
            selectedText: 'Auto Note Text',
            selectionRegion: { top: 10, left: 10, width: 100, height: 20 },
            highlightAreas: [{ pageIndex: 0, left: 10, top: 10, width: 100, height: 20 }],
            toggle: vi.fn()
        };

        // Render the HighlightLogic (which has the useEffect for auto-noting)
        render(capturedRenderHighlightTarget(props));

        // 3. Verify actions (should be called via useEffect)
        expect(addHighlight).toHaveBeenCalled();
        expect(pushAction).toHaveBeenCalledWith(expect.objectContaining({
            type: 'INSERT_BULLET',
            payload: 'Auto Note Text'
        }));
    });
});