import React, { useEffect } from 'react'
import { Panel, PanelGroup, PanelResizeHandle } from 'react-resizable-panels'
import Editor from './components/Editor'
import PDFViewer from './components/PDFViewer'
import Sidebar from './components/Sidebar'
import ErrorBoundary from './components/ErrorBoundary'
import { useNoteStore } from './store/useNoteStore'
import { Worker } from '@react-pdf-viewer/core'
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.js?url'

function App(): JSX.Element {
  const { setPdf } = useNoteStore()

  useEffect(() => {
    // Expose setPdf to window for iCharlotte integration
    // @ts-ignore
    window.setPdfExternal = (path: string) => {
      if (!path || typeof path !== 'string') {
          console.error("setPdfExternal received invalid path:", path);
          return;
      }
      console.log("App: setPdfExternal called with:", path);
      const parts = path.split(/[\\\/]/);
      const fileName = (parts && parts.length > 0) ? (parts.pop() || 'document.pdf') : 'document.pdf';
      try {
        setPdf(path, fileName, path)
      } catch (e) {
        console.error("App: Error in setPdf:", e);
      }
    }
    // @ts-ignore
    window.loadPdfExternal = window.setPdfExternal
  }, [setPdf])

  return (
    <Worker workerUrl={pdfWorkerUrl}>
      <div className="h-screen w-screen overflow-hidden flex bg-gray-900 text-white">
        <ErrorBoundary name="Sidebar">
          <Sidebar />
        </ErrorBoundary>
        <div className="flex-1 flex flex-col h-full overflow-hidden bg-white text-black">
           <PanelGroup direction="horizontal">
            <Panel defaultSize={50} minSize={20} className="bg-white">
              <ErrorBoundary name="Editor">
                <Editor />
              </ErrorBoundary>
            </Panel>
            <PanelResizeHandle className="w-2 bg-gray-200 hover:bg-gray-400 transition-colors cursor-col-resize flex items-center justify-center border-l border-r border-gray-300">
              <div className="h-4 w-1 bg-gray-400 rounded"></div>
            </PanelResizeHandle>
            <Panel defaultSize={50} minSize={20} className="bg-gray-100">
              <ErrorBoundary name="PDFViewer">
                <PDFViewer />
              </ErrorBoundary>
            </Panel>
          </PanelGroup>
        </div>
      </div>
    </Worker>
  )
}

export default App