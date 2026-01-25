import React, { useState, useEffect, useRef, useCallback, useMemo } from 'react'
import { Viewer, SpecialZoomLevel } from '@react-pdf-viewer/core'
import { zoomPlugin } from '@react-pdf-viewer/zoom'
import { highlightPlugin, Trigger, RenderHighlightTargetProps, RenderHighlightsProps } from '@react-pdf-viewer/highlight'
import '@react-pdf-viewer/highlight/lib/styles/index.css'
import { useNoteStore } from '../store/useNoteStore'
import clsx from 'clsx'

// Component to handle the logic within the highlight target render prop
const HighlightLogic: React.FC<RenderHighlightTargetProps & { autoNoteRef: React.MutableRefObject<boolean> }> = ({ 
    selectionRegion, 
    highlightAreas,
    selectedText, 
    toggle, 
    autoNoteRef 
}) => {
    const { addHighlight, pushAction } = useNoteStore()

    useEffect(() => {
        if (autoNoteRef.current && selectedText) {
            console.log("PDFViewer: Auto-noting selected text:", selectedText.substring(0, 20))
            const id = typeof crypto.randomUUID === 'function' ? crypto.randomUUID() : Math.random().toString(36).substring(2)
            const newHighlight = {
                id,
                content: selectedText,
                highlightAreas: highlightAreas || [],
                quote: selectedText
            }
            
            // 1. Persist highlight
            addHighlight(newHighlight)
            
            // 2. Add to editor
            pushAction({
                type: 'INSERT_BULLET',
                payload: selectedText,
                highlightId: id
            })
            
            // 3. Clear selection
            if (toggle) {
                try {
                    toggle();
                } catch (e) {
                    console.error("PDFViewer: Error calling toggle() in useEffect:", e);
                }
            }
        }
    }, [selectedText, selectionRegion, highlightAreas, toggle, addHighlight, pushAction, autoNoteRef])

    if (autoNoteRef.current) return null
    if (!selectionRegion) return null

    return (
        <div
            style={{
                background: '#fff',
                border: '1px solid rgba(0, 0, 0, .3)',
                borderRadius: '2px',
                padding: '8px',
                position: 'absolute',
                left: `${selectionRegion.left}%`,
                top: `${selectionRegion.top + (selectionRegion.height || 0)}%`,
                zIndex: 1,
            }}
        >
            <button 
                className="bg-blue-600 text-white px-2 py-1 rounded text-sm hover:bg-blue-700"
                onClick={() => {
                    const id = typeof crypto.randomUUID === 'function' ? crypto.randomUUID() : Math.random().toString(36).substring(2)
                    const newHighlight = {
                        id,
                        content: selectedText,
                        highlightAreas: highlightAreas || [],
                        quote: selectedText
                    }
                    addHighlight(newHighlight)
                    pushAction({
                        type: 'INSERT_BULLET',
                        payload: selectedText,
                        highlightId: id
                    })
                    if (toggle) {
                        try {
                            toggle();
                        } catch (e) {
                            console.error("PDFViewer: Error calling toggle() in button:", e);
                        }
                    }
                }}
            >
                Add Note
            </button>
        </div>
    )
}

const PDFViewer: React.FC = () => {
    console.log("PDFViewer: Render cycle started")
    const { pdfUrl, setPdf, autoNote, setAutoNote, highlights, addHighlight, pushAction } = useNoteStore()
    const [pdfBlobUrl, setPdfBlobUrl] = useState<string | null>(null)
    const pdfBlobUrlRef = useRef<string | null>(null)
    const [error, setError] = useState<string | null>(null)
    const [isBridgeReady, setIsBridgeReady] = useState(!!(window as any).api)
    
    // Refs for stable plugin access
    const autoNoteRef = useRef(autoNote)
    const highlightsRef = useRef(highlights)

    useEffect(() => {
        autoNoteRef.current = autoNote
    }, [autoNote])

    useEffect(() => {
        highlightsRef.current = highlights
    }, [highlights])

    // Track the current scale to allow continuous zooming
    const currentScaleRef = useRef<number>(1)
    const optimisticScaleRef = useRef<number | null>(null)
    const resetOptimisticTimeoutRef = useRef<NodeJS.Timeout | null>(null)
    
    // Initialize zoom plugin
    const zoomPluginInstance = useMemo(() => zoomPlugin({
        onZoom: (e) => {
            currentScaleRef.current = e.scale
        }
    }), [])
    const { zoomTo } = zoomPluginInstance

    // Initialize highlight plugin
    const renderHighlightTarget = useCallback((props: RenderHighlightTargetProps) => {
        return <HighlightLogic {...props} autoNoteRef={autoNoteRef} />
    }, [])

    const renderHighlights = useCallback((props: RenderHighlightsProps) => {
        try {
            console.log("PDFViewer: renderHighlights called for page", props.pageIndex)
            if (!highlightsRef.current) {
                console.warn("PDFViewer: highlightsRef.current is undefined")
                return <div />
            }
            if (!Array.isArray(highlightsRef.current)) {
                console.warn("PDFViewer: highlightsRef.current is not an array:", highlightsRef.current)
                return <div />
            }

            const pageHighlights = (highlightsRef.current || []).filter((highlight) => {
                if (!highlight) return false
                if (!highlight.highlightAreas || !Array.isArray(highlight.highlightAreas)) {
                    console.warn("PDFViewer: Highlight missing highlightAreas:", highlight)
                    return false
                }
                return highlight.highlightAreas.some((area: any) => area && area.pageIndex === props.pageIndex)
            });

            console.log(`PDFViewer: Found ${(pageHighlights || []).length} highlights for page ${props.pageIndex}`);

            return (
                <div>
                    {(pageHighlights || []).map((highlight) => {
                        if (!highlight || typeof highlight !== 'object' || !highlight.id) {
                            console.warn("PDFViewer: Invalid highlight object:", highlight);
                            return null;
                        }
                        return (
                            <React.Fragment key={highlight.id}>
                                {(highlight.highlightAreas || [])
                                    .filter((area: any) => area && typeof area === 'object' && area.pageIndex === props.pageIndex)
                                    .map((area: any, idx: number) => {
                                        if (!area) return null;
                                        try {
                                            const cssProps = props.getCssProperties(area, props.rotation);
                                            return (
                                                <div
                                                    key={`${highlight.id}-${idx}`}
                                                    style={Object.assign(
                                                        {},
                                                        {
                                                            background: 'yellow',
                                                            opacity: 0.4,
                                                        },
                                                        cssProps
                                                    )}
                                                />
                                            );
                                        } catch (err) {
                                            console.error("PDFViewer: Error getting CSS properties for highlight area:", err, area);
                                            return null;
                                        }
                                    })}
                            </React.Fragment>
                        );
                    })}
                </div>
            )
        } catch (err) {
            console.error("PDFViewer: Fatal error in renderHighlights:", err);
            return <div />;
        }
    }, [])

    const highlightPluginInstance = useMemo(() => highlightPlugin({
        trigger: Trigger.TextSelection,
        renderHighlightTarget,
        renderHighlights,
    }), [renderHighlightTarget, renderHighlights])

    // Update ref when state changes
    useEffect(() => {
        pdfBlobUrlRef.current = pdfBlobUrl
    }, [pdfBlobUrl])

    // Listen for bridge readiness
    useEffect(() => {
        const handleBridgeReady = () => {
            console.log("PDFViewer: Bridge ready event received")
            setIsBridgeReady(true)
        }
        window.addEventListener('bridge-ready', handleBridgeReady)
        return () => window.removeEventListener('bridge-ready', handleBridgeReady)
    }, [])

    // Handle Ctrl + Wheel Zoom
    const containerRef = useRef<HTMLDivElement>(null)
    useEffect(() => {
        const container = containerRef.current
        if (!container) return

        const handleWheel = (e: WheelEvent) => {
            if (e.ctrlKey) {
                e.preventDefault()
                
                // Clear any pending reset of the optimistic scale
                if (resetOptimisticTimeoutRef.current) {
                    clearTimeout(resetOptimisticTimeoutRef.current)
                }

                const zoomStep = 0.2
                // Use optimistic scale if valid (during scroll burst), otherwise fallback to confirmed scale
                const baseScale = optimisticScaleRef.current !== null ? optimisticScaleRef.current : currentScaleRef.current
                
                let nextScale = e.deltaY < 0 
                    ? baseScale + zoomStep 
                    : baseScale - zoomStep
                
                // Clamp
                nextScale = Math.max(0.1, Math.min(5.0, nextScale)) // Increased max zoom to 500%

                // Update optimistic ref IMMEDIATELY
                optimisticScaleRef.current = nextScale
                
                zoomTo(nextScale)

                // Reset optimistic ref after scroll stops (100ms) to resync with real state
                resetOptimisticTimeoutRef.current = setTimeout(() => {
                    optimisticScaleRef.current = null
                }, 100)
            }
        }

        container.addEventListener('wheel', handleWheel, { passive: false })
        return () => {
            container.removeEventListener('wheel', handleWheel)
            if (resetOptimisticTimeoutRef.current) clearTimeout(resetOptimisticTimeoutRef.current)
        }
    }, [zoomTo])

    // Load PDF Data
    useEffect(() => {
        if (!pdfUrl) {
            console.log("PDFViewer: No pdfUrl provided, clearing data.")
            if (pdfBlobUrlRef.current && typeof pdfBlobUrlRef.current === 'string' && pdfBlobUrlRef.current.startsWith('blob:')) {
                URL.revokeObjectURL(pdfBlobUrlRef.current)
            }
            setPdfBlobUrl(null)
            setError(null)
            return
        }
        
        let active = true;

        const loadData = async () => {
            setError(null)
            try {
                console.log("PDFViewer: Starting load process for:", pdfUrl)
                if (!pdfUrl) {
                    console.log("PDFViewer: pdfUrl is empty, skipping load")
                    return;
                }
                
                // PRIORITIZE local-resource scheme for performance (streaming)
                if (typeof pdfUrl === 'string') {
                    console.log("PDFViewer: Using local-resource scheme for string path:", pdfUrl)
                    
                    const normalizedPath = (pdfUrl || "").replace(/\\/g, '/');
                    const segments = normalizedPath.split('/').filter(s => s !== undefined && s !== null);
                    console.log("PDFViewer: Path segments count:", segments ? segments.length : 'undefined')
                    
                    if (!segments || !Array.isArray(segments)) {
                        throw new Error("Failed to parse path segments");
                    }

                    const safePath = segments.map(s => {
                        try {
                            return encodeURIComponent(s || "");
                        } catch (e) {
                            console.error("PDFViewer: Failed to encode segment:", s, e);
                            return s || "";
                        }
                    }).join('/')
                    
                    console.log("PDFViewer: Constructed safePath length:", safePath ? safePath.length : 'undefined')
                    const resourceUrl = `local-resource:///${safePath}`
                    console.log("PDFViewer: Final resourceUrl:", resourceUrl)
                    
                    if (pdfBlobUrlRef.current && typeof pdfBlobUrlRef.current === 'string' && pdfBlobUrlRef.current.startsWith('blob:')) {
                        console.log("PDFViewer: Revoking old blob URL:", pdfBlobUrlRef.current)
                        URL.revokeObjectURL(pdfBlobUrlRef.current)
                    }
                    setPdfBlobUrl(resourceUrl)
                }
                else {
                    console.warn("PDFViewer: pdfUrl is not a string, type is:", typeof pdfUrl, pdfUrl)
                    const win = window as any
                    if (win.api && win.api.readBuffer) {
                        console.log("PDFViewer: Reading buffer via bridge for non-string pdfUrl")
                        const buffer = await win.api.readBuffer(pdfUrl)
                        if (!active) return
                        if (buffer && buffer.byteLength !== undefined) {
                            console.log("PDFViewer: Buffer received, byteLength:", buffer.byteLength)
                            if (buffer.byteLength > 0) {
                                const blob = new Blob([buffer], { type: 'application/pdf' })
                                const blobUrl = URL.createObjectURL(blob)
                                
                                if (pdfBlobUrlRef.current && typeof pdfBlobUrlRef.current === 'string' && pdfBlobUrlRef.current.startsWith('blob:')) {
                                    URL.revokeObjectURL(pdfBlobUrlRef.current)
                                }
                                setPdfBlobUrl(blobUrl)
                            } else {
                                throw new Error("Buffer byteLength is 0")
                            }
                        } else {
                            throw new Error(`Failed to read file buffer (result is ${buffer ? 'empty' : 'null'}).`)
                        }
                    } else {
                        throw new Error(`Bridge API not ready for: ${pdfUrl}`)
                    }
                }
            } catch (e) {
                if (!active) return
                const msg = `Exception loading PDF: ${e}`
                console.error("PDFViewer: " + msg)
                setError(msg)
            }
        }

        loadData()
        
        return () => {
            active = false
        }
    }, [pdfUrl, isBridgeReady])

    const handleDocumentLoad = useCallback(() => {
        console.log("PDFViewer: Document loaded successfully")
        setError(null)
    }, [])

    const handleDocumentError = useCallback((err: any) => {
        console.error("PDFViewer: Document load error:", err)
        const errorMsg = err ? (err.message || String(err)) : "Unknown error"
        setError(`Failed to load PDF: ${errorMsg}`)
    }, [])

    const handleDrop = (e: React.DragEvent) => {
        e.preventDefault()
        try {
            if (e.dataTransfer && e.dataTransfer.files) {
                const files = e.dataTransfer.files;
                if (files.length > 0) {
                    const file = files[0]
                    if (file && file.name && (file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf'))) {
                        // @ts-ignore
                        const path = file.path
                        if (path) {
                            console.log("PDFViewer: File dropped, path:", path);
                            setPdf(path, file.name, path)
                        } else {
                            console.warn("PDFViewer: Dropped file has no path property. This might be a browser security restriction.");
                        }
                    }
                }
            }
        } catch (err) {
            console.error("PDFViewer: Error handling drop:", err);
        }
    }

    // Cleanup on UNMOUNT
    useEffect(() => {
        return () => {
            if (pdfBlobUrlRef.current && pdfBlobUrlRef.current.startsWith('blob:')) {
                URL.revokeObjectURL(pdfBlobUrlRef.current)
            }
        }
    }, [])

    return (
        <div 
            ref={containerRef}
            className="h-full w-full bg-gray-100 relative flex flex-col"
            onDrop={handleDrop}
            onDragOver={(e) => e.preventDefault()}
        >
             <div className="bg-white border-b px-4 py-2 flex items-center justify-between shadow-sm z-10">
                <div className="flex items-center gap-2">
                    <span className="text-sm font-medium text-gray-700">Auto Note</span>
                    <button 
                        onClick={() => setAutoNote(!autoNote)}
                        className={clsx(
                            "w-10 h-5 rounded-full transition-colors relative focus:outline-none focus:ring-2 focus:ring-offset-1 focus:ring-blue-500",
                            autoNote ? "bg-blue-600" : "bg-gray-300"
                        )}
                    >
                        <div className={clsx(
                            "w-4 h-4 bg-white rounded-full absolute top-0.5 transition-transform shadow-sm",
                            autoNote ? "left-5" : "left-0.5"
                        )} />
                    </button>
                </div>
             </div>

             {!pdfUrl ? (
                <div className="flex flex-col items-center justify-center h-full text-gray-400">
                    <p className="text-lg">Select a PDF from the sidebar or Drag & Drop here</p>
                </div>
             ) : error ? (
                <div className="flex flex-col items-center justify-center h-full text-red-500 p-4 text-center">
                    <p className="text-lg font-semibold mb-2">Error Loading PDF</p>
                    <p className="text-sm font-mono whitespace-pre-wrap">{error}</p>
                </div>
             ) : (
                <div className="flex-1 overflow-auto relative">
                    {pdfBlobUrl ? (
                        <Viewer 
                            key={`${pdfUrl}-${pdfBlobUrl}`}
                            fileUrl={pdfBlobUrl} 
                            onDocumentLoad={handleDocumentLoad}
                            onDocumentError={handleDocumentError}
                            defaultScale={SpecialZoomLevel.PageWidth}
                            plugins={[zoomPluginInstance, highlightPluginInstance]}
                        />
                    ) : (
                        <div className="flex items-center justify-center h-full text-gray-500">
                            Loading PDF data...
                        </div>
                    )}
                </div>
             )}
        </div>
    )
}

export default PDFViewer