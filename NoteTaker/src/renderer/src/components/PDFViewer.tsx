import React, { useState, useEffect, useCallback, useRef, useMemo } from 'react'
import { 
  PdfLoader, 
  PdfHighlighter, 
  Highlight, 
  Popup, 
  AreaHighlight 
} from 'react-pdf-highlighter'
import { useNoteStore } from '../store/useNoteStore'
import { constructLocalPdfUrl } from '../utils/pdfUrlUtils'
import { Search, ChevronDown, ChevronUp, X } from 'lucide-react'
import clsx from 'clsx'

const getNextId = () => String(Math.random()).slice(2)

const parseIdFromHash = () => document.location.hash.slice("#highlight-".length)

const resetHash = () => {
  document.location.hash = ""
}

const HighlightPopup = ({ comment, }: { comment: { text: string; emoji: string } }) => (
  comment.text ? (
    <div className="absolute bg-white border rounded p-2 shadow-lg text-xs z-50 transform -translate-y-full -mt-2">
      {comment.emoji} {comment.text}
    </div>
  ) : null
)

const PDFViewer: React.FC = () => {
  const { 
    pdfUrl, 
    addHighlight, 
    pushAction, 
    autoNote, 
    setAutoNote, 
    highlights, 
    jumpToHighlight, 
    scrollToHighlightId,
    isOcrRunning,
    ocrMessage,
    highlightColor,
    zoom,
    setZoom,
    debouncedZoom,
    setDebouncedZoom
  } = useNoteStore()
  const [error, setError] = useState<string | null>(null)
  const viewerContainerRef = useRef<HTMLDivElement>(null)

  // Search State
  const [searchActive, setSearchActive] = useState(false)
  const [searchValue, setSearchValue] = useState('')
  const [searchResults, setSearchResults] = useState<any[]>([])
  const [currentSearchResultIndex, setCurrentSearchResultIndex] = useState(-1)
  const [isSearching, setIsSearching] = useState(false)

  // Zoom Debounce
  useEffect(() => {
    const timer = setTimeout(() => {
      setDebouncedZoom(zoom);
    }, 100);
    return () => clearTimeout(timer);
  }, [zoom, setDebouncedZoom]);

  // Ref to access current highlights in callbacks without dependency cycles
  const highlightsRef = useRef(highlights)
  
  // Sync ref with store
  useEffect(() => {
    highlightsRef.current = highlights
  }, [highlights])

  const scrollViewerTo = useRef<(highlight: any) => void>(() => {})

  const scrollToHighlightFromHash = useCallback(() => {
    const highlight = highlightsRef.current.find((h) => h.id === parseIdFromHash())
    if (highlight) {
      scrollViewerTo.current(highlight)
    }
  }, [])

  useEffect(() => {
    window.addEventListener("hashchange", scrollToHighlightFromHash)
    return () => {
      window.removeEventListener("hashchange", scrollToHighlightFromHash)
    }
  }, [scrollToHighlightFromHash])

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
        if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
            e.preventDefault();
            setSearchActive(prev => !prev);
        }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, []);

  useEffect(() => {
    const handleWheel = (e: WheelEvent) => {
      if (e.ctrlKey || e.metaKey) {
        e.preventDefault();
        const currentZoom = useNoteStore.getState().zoom;
        const delta = e.deltaY > 0 ? -0.1 : 0.1;
        const newZoom = Math.min(Math.max(currentZoom + delta, 0.5), 3.0);
        setZoom(parseFloat(newZoom.toFixed(2)));
      }
    };

    window.addEventListener('wheel', handleWheel, { passive: false });
    return () => window.removeEventListener('wheel', handleWheel);
  }, [setZoom]);

  const executeSearch = useCallback(async (text: string, pdfDocument: any) => {
    if (!text || text.length < 3 || !pdfDocument) {
      setSearchResults([]);
      setCurrentSearchResultIndex(-1);
      return;
    }

    setIsSearching(true);
    const results: any[] = [];
    
    try {
        for (let i = 1; i <= pdfDocument.numPages; i++) {
            const page = await pdfDocument.getPage(i);
            const textContent = await page.getTextContent();
            const pageText = textContent.items.map((item: any) => item.str).join(" ");
            
            if (pageText.toLowerCase().includes(text.toLowerCase())) {
                results.push({
                    pageNumber: i,
                    id: `search-result-${i}-${Math.random()}`,
                    content: { text: `Match on page ${i}` },
                    position: {
                        pageNumber: i,
                        boundingRect: { x1: 0, y1: 0, x2: 0, y2: 0, width: 0, height: 0 },
                        rects: []
                    }
                });
            }
        }
        setSearchResults(results);
        if (results.length > 0) {
            setCurrentSearchResultIndex(0);
            scrollViewerTo.current(results[0]);
        } else {
            setCurrentSearchResultIndex(-1);
        }
    } catch (err) {
        console.error("Search failed", err);
    } finally {
        setIsSearching(false);
    }
  }, []);

  const nextSearchResult = () => {
      if (searchResults.length === 0) return;
      const nextIndex = (currentSearchResultIndex + 1) % searchResults.length;
      setCurrentSearchResultIndex(nextIndex);
      scrollViewerTo.current(searchResults[nextIndex]);
  };

  const prevSearchResult = () => {
      if (searchResults.length === 0) return;
      const nextIndex = (currentSearchResultIndex - 1 + searchResults.length) % searchResults.length;
      setCurrentSearchResultIndex(nextIndex);
      scrollViewerTo.current(searchResults[nextIndex]);
  };
  
  useEffect(() => {
    if (scrollToHighlightId) {
        const highlight = highlightsRef.current.find(h => h.id === scrollToHighlightId)
        if (highlight) {
            scrollViewerTo.current(highlight)
        }
    }
  }, [scrollToHighlightId])

  // Custom "Add Note" Popup
  const renderTip = useCallback(({ position, content, onConfirm, onOpen, onUpdate }) => {
     const onSummarize = async () => {
         if (!content.text) return;
         
         const actionId = getNextId();
         pushAction({
             type: 'INSERT_BULLET',
             payload: `<em>Summarizing: "${content.text.substring(0, 50)}..."</em>`,
             highlightId: actionId
         });

         onConfirm({ text: "Summary pending...", emoji: "🤖" });

         if (window.api && window.api.summarizeText) {
             try {
                 const summary = await window.api.summarizeText(content.text);
                 pushAction({
                     type: 'INSERT_BULLET',
                     payload: `<strong>AI Summary:</strong> ${summary}`,
                     highlightId: actionId
                 });
             } catch (err) {
                 console.error("Summarization failed", err);
             }
         }
     }

     if (autoNote) {
         setTimeout(() => {
             onConfirm({ text: "", emoji: "" })
         }, 0)
         return null
     }

     return (
         <div className="absolute bg-white border border-gray-300 rounded shadow-sm px-2 py-1 flex gap-2 transform -translate-y-full -mt-2" style={{ left: 0, top: 0 }}>
             <button 
                 className="text-black text-xs px-2 py-1 rounded hover:opacity-80 whitespace-nowrap font-medium shadow-sm"
                 style={{ backgroundColor: highlightColor }}
                 onClick={() => onConfirm({ text: "", emoji: "" })}
             >
                 Add Note
             </button>
             <button 
                 className="bg-purple-600 text-white text-xs px-2 py-1 rounded hover:bg-purple-700 whitespace-nowrap font-medium shadow-sm flex items-center gap-1"
                 onClick={onSummarize}
             >
                 <span className="text-[10px]">🤖</span> Summarize
             </button>
         </div>
     )
  }, [autoNote, highlightColor, pushAction]);

  // Process Selection
  const addHighlightEntry = useCallback((highlight: any) => {
    console.log("Saving highlight", highlight)
    const highlightWithColor = { ...highlight, color: highlightColor };
    addHighlight(highlightWithColor)
    
    if (highlight.content && highlight.content.text) {
        pushAction({
            type: 'INSERT_BULLET',
            payload: highlight.content.text,
            highlightId: highlight.id
        })
    } else if (highlight.content && highlight.content.image) {
        pushAction({
            type: 'INSERT_IMAGE',
            payload: highlight.content.image,
            highlightId: highlight.id
        })
    }
  }, [highlightColor, addHighlight, pushAction]);

  const prevSearchValueRef = useRef('');
  const pdfDocRef = useRef<any>(null);

  useEffect(() => {
      if (searchValue !== prevSearchValueRef.current && searchValue.length >= 3 && pdfDocRef.current) {
          const timeout = setTimeout(() => {
              executeSearch(searchValue, pdfDocRef.current);
              prevSearchValueRef.current = searchValue;
          }, 500);
          return () => clearTimeout(timeout);
      }
  }, [searchValue, executeSearch]);

  const highlighterRef = useRef<any>(null);

  // Apply Zoom manually to avoid the "jump"
  useEffect(() => {
    if (highlighterRef.current && highlighterRef.current.viewer) {
        try {
            console.log("Applying zoom:", debouncedZoom);
            highlighterRef.current.viewer.currentScaleValue = debouncedZoom.toString();
        } catch (e) {
            console.error("Failed to set zoom", e);
        }
    }
  }, [debouncedZoom]);

  const highlightTransform = useCallback((
    highlight,
    index,
    setTip,
    hideTip,
    viewportToScaled,
    screenshot,
    isScrolledTo
  ) => {
      const isTextHighlight = !highlight.content.image

      const component = isTextHighlight ? (
          <div style={{ '--highlight-color': highlight.color } as any}>
              <Highlight
                  isScrolledTo={isScrolledTo}
                  position={highlight.position}
                  comment={highlight.comment}
              />
          </div>
      ) : (
          <AreaHighlight
              isScrolledTo={isScrolledTo}
              highlight={highlight}
              onChange={() => {}}
          />
      )

      return (
          <Popup
              popupContent={<HighlightPopup {...highlight} />}
              onMouseOver={(popupContent) =>
                  setTip(highlight, () => popupContent)
              }
              onMouseOut={hideTip}
              key={index}
              children={component}
          />
      )
  }, [])

  const onSelectionFinished = useCallback((
    position,
    content,
    hideTipAndSelection
  ) => (
      renderTip({
          position,
          content,
          onOpen: () => {},
          onConfirm: (comment) => {
              addHighlightEntry({ content, position, comment, id: getNextId() })
              hideTipAndSelection()
          },
          onUpdate: () => {}
      })
  ), [renderTip, addHighlightEntry])

  if (!pdfUrl) {
     return (
        <div className="flex flex-col items-center justify-center h-full text-gray-400">
            <p className="text-lg">Select a PDF from the sidebar</p>
        </div>
     )
  }
  
  let finalPdfUrl = constructLocalPdfUrl(pdfUrl);

  const renderPdfContent = useCallback((pdfDocument) => {
      pdfDocRef.current = pdfDocument;
      return (
          <PdfHighlighter
              ref={highlighterRef}
              key={highlightColor} 
              pdfDocument={pdfDocument}
              enableAreaSelection={(event) => event.altKey}
              onScrollChange={() => {}}
              scrollRef={(scrollTo) => {
                  scrollViewerTo.current = scrollTo
                  scrollToHighlightFromHash()
              }}
              onSelectionFinished={onSelectionFinished}
              highlightTransform={highlightTransform}
              highlights={highlights}
          />
      )
  }, [highlightColor, highlights, onSelectionFinished, highlightTransform, scrollToHighlightFromHash]);

  return (
    <div className="h-full w-full relative flex flex-col bg-gray-100" ref={viewerContainerRef}>
         <div className="bg-white border-b px-4 py-2 flex items-center justify-between shadow-sm z-10 shrink-0">
            <div className="flex items-center gap-4">
                <div className="flex items-center gap-2">
                    <span className="text-xs font-medium text-gray-500 uppercase tracking-wider">Auto Note</span>
                    <button 
                        onClick={() => setAutoNote(!autoNote)}
                        className={clsx(
                            "w-8 h-4 rounded-full transition-colors relative focus:outline-none focus:ring-1 focus:ring-offset-1 focus:ring-blue-500",
                            autoNote ? "bg-blue-600" : "bg-gray-300"
                        )}
                    >
                        <div className={clsx(
                            "w-3 h-3 bg-white rounded-full absolute top-0.5 transition-transform shadow-sm",
                            autoNote ? "left-5" : "left-0.5"
                        )} />
                    </button>
                </div>
                
                <button 
                    onClick={() => setSearchActive(!searchActive)}
                    className={clsx(
                        "p-1.5 rounded hover:bg-gray-100 transition-colors",
                        searchActive ? "text-blue-600 bg-blue-50" : "text-gray-500"
                    )}
                    title="Find in document (Ctrl+F)"
                >
                    <Search size={16} />
                </button>

                <div className="flex items-center bg-gray-100 px-2 py-1 rounded border border-gray-200 ml-2">
                    <span className="text-[10px] font-bold text-gray-500 font-mono w-10 text-center">
                        {Math.round(zoom * 100)}%
                    </span>
                </div>
            </div>

            {searchActive && (
                <div className="flex-1 max-w-sm mx-4 flex items-center gap-1 bg-gray-100 px-2 py-1 rounded border border-gray-200">
                    <Search size={12} className={clsx("text-gray-400", isSearching && "animate-pulse text-blue-500")} />
                    <input 
                        type="text"
                        placeholder="Find (3+ chars)..."
                        value={searchValue}
                        onChange={(e) => setSearchValue(e.target.value)}
                        onKeyDown={(e) => {
                            if (e.key === 'Enter') {
                                if (searchResults.length > 0) nextSearchResult();
                            }
                        }}
                        autoFocus
                        className="bg-transparent border-none outline-none text-xs flex-1"
                    />
                    {searchResults.length > 0 && (
                        <div className="flex items-center gap-1 border-l pl-1 ml-1">
                            <span className="text-[10px] text-gray-500 whitespace-nowrap">
                                {currentSearchResultIndex + 1} / {searchResults.length}
                            </span>
                            <button onClick={prevSearchResult} className="p-0.5 hover:bg-gray-200 rounded text-gray-600">
                                <ChevronUp size={12} />
                            </button>
                            <button onClick={nextSearchResult} className="p-0.5 hover:bg-gray-200 rounded text-gray-600">
                                <ChevronDown size={12} />
                            </button>
                        </div>
                    )}
                    <button onClick={() => {
                        setSearchActive(false);
                        setSearchValue('');
                        setSearchResults([]);
                    }} className="text-gray-400 hover:text-gray-600 ml-1">
                        <X size={12} />
                    </button>
                </div>
            )}

            {isOcrRunning && (
                <div className="flex items-center gap-2 bg-yellow-50 px-3 py-1 rounded border border-yellow-200 animate-pulse">
                    <span className="text-[10px] font-bold text-yellow-700 uppercase tracking-tighter">{ocrMessage}</span>
                </div>
            )}
         </div>

        <div className="flex-1 relative overflow-hidden">
             <PdfLoader url={finalPdfUrl} beforeLoad={<div className="p-4 text-gray-500">Loading PDF...</div>}> 
                {renderPdfContent}
            </PdfLoader>
        </div>
    </div>
  )
}

export default PDFViewer