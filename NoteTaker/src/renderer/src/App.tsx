import React, { useEffect, useState, useRef } from 'react'
import { Panel, PanelGroup, PanelResizeHandle } from 'react-resizable-panels'
import Editor from './components/Editor'
import PDFViewer from './components/PDFViewer'
import Sidebar from './components/Sidebar'
import ErrorBoundary from './components/ErrorBoundary'
import { useNoteStore } from './store/useNoteStore'
import clsx from 'clsx'

function App(): JSX.Element {
  const { setPdf, loadFromState, openFiles, fileStates, currentFilePath, hasLoaded, setOcrStatus, isSaving, setIsSaving } = useNoteStore()
  const [bridgeStatus, setBridgeStatus] = useState<'waiting' | 'ready' | 'missing'>(
    (window as any).api ? 'ready' : 'waiting'
  )

  // Load state on startup
  useEffect(() => {
      const loadState = async () => {
          // @ts-ignore
          if (bridgeStatus === 'ready' && window.api && window.api.loadAppState && !hasLoaded) {
              try {
                  console.log("App: Attempting to load state from disk...");
                  // @ts-ignore
                  const state = await window.api.loadAppState();
                  if (state) {
                      console.log("App: Loaded state successfully", state);
                      loadFromState(state);
                  } else {
                      console.log("App: No saved state found, starting fresh.");
                      loadFromState(null); // Mark as loaded even if null
                  }
              } catch (err) {
                  console.error("App: Failed to load state", err);
                  loadFromState(null);
              }
          }
      }
      loadState();
  }, [bridgeStatus, loadFromState, hasLoaded]);

  // Save state on changes (debounced)
  const saveTimeoutRef = useRef<NodeJS.Timeout | null>(null);
  useEffect(() => {
      if (!hasLoaded) return; // DON'T save until we've at least tried to load

      if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current);
      
      saveTimeoutRef.current = setTimeout(async () => {
          // @ts-ignore
          if (window.api && window.api.saveAppState) {
              setIsSaving(true);
              const stateToSave = { openFiles, fileStates, currentFilePath };
              console.log("App: Saving state to disk", stateToSave);
              try {
                  // @ts-ignore
                  await window.api.saveAppState(stateToSave);
              } finally {
                  // Artificial delay for UX visibility if it's too fast
                  setTimeout(() => setIsSaving(false), 500);
              }
          }
      }, 1000);

      return () => {
          if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current);
      }
  }, [openFiles, fileStates, currentFilePath, hasLoaded, setIsSaving]);

  useEffect(() => {
    console.log("App: Initializing integration APIs");
    
    // @ts-ignore
    window.setPdfExternal = (path: string) => {
      if (!path) return;
      console.log("App: setPdfExternal ->", path);
      const parts = path.split(/[\\\/]/);
      const fileName = (parts && parts.length > 0) ? (parts.pop() || 'document.pdf') : 'document.pdf';
      setPdf(path, fileName, path);
    }
    
    // @ts-ignore
    window.loadPdfExternal = window.setPdfExternal;

    const handleBridgeReady = () => {
      console.log("App: bridge-ready event received");
      setBridgeStatus('ready');
    };

    const handleOcrFinished = (e: any) => {
        const { path, success, message } = e.detail;
        console.log("App: OCR Finished event:", path, success);
        setOcrStatus(false, "");
        if (success && path === currentFilePath) {
            // Force reload by resetting current PDF
            // This is a bit hacky but it works.
            const title = openFiles.find(f => f.path === path)?.name || 'Document';
            setPdf(path, title, path);
        }
    };

    window.addEventListener('ocr-finished', handleOcrFinished);

    if (!(window as any).api) {
      window.addEventListener('bridge-ready', handleBridgeReady);
      
      const timeout = setTimeout(() => {
        if (!(window as any).api && !(window as any).bridge_ready) {
            console.warn("App: Bridge initialization timed out");
            setBridgeStatus('missing');
        } else {
            setBridgeStatus('ready');
        }
      }, 10000);

      return () => {
        window.removeEventListener('bridge-ready', handleBridgeReady);
        window.removeEventListener('ocr-finished', handleOcrFinished);
        clearTimeout(timeout);
      };
    } else {
      setBridgeStatus('ready');
    }

    return () => {
        window.removeEventListener('ocr-finished', handleOcrFinished);
    }
  }, [setPdf, currentFilePath, openFiles, setOcrStatus])

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    
    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
        const file = files[0];
        // Electron exposes 'path' on File objects
        // @ts-ignore
        const path = file.path;
        
        if (path && (path.toLowerCase().endsWith('.pdf') || file.type === 'application/pdf')) {
             console.log("App: Dropped file ->", path);
             // @ts-ignore
             if (window.setPdfExternal) {
                 // @ts-ignore
                 window.setPdfExternal(path);
             }
        }
    }
  };

  return (
    <div 
        className="h-screen w-screen overflow-hidden flex flex-col bg-gray-900 text-white"
        onDragOver={handleDragOver}
        onDrop={handleDrop}
    >
      {bridgeStatus !== 'ready' && (
        <div className={clsx(
            "px-4 py-1 text-[10px] text-center font-medium uppercase tracking-widest",
            bridgeStatus === 'waiting' ? "bg-yellow-600 animate-pulse" : "bg-red-600"
        )}>
            {bridgeStatus === 'waiting' ? "Connecting to iCharlotte..." : "iCharlotte Connection Offline"}
        </div>
      )}
      {isSaving && (
          <div className="absolute top-2 right-4 z-50 bg-green-600/80 text-white text-[9px] px-2 py-0.5 rounded-full font-bold uppercase tracking-widest animate-pulse pointer-events-none">
              Saving...
          </div>
      )}
      <div className="flex-1 flex overflow-hidden">
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
    </div>
  )
}

export default App
