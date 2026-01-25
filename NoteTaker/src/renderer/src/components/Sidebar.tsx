import React from 'react'
import { useNoteStore } from '../store/useNoteStore'
import { FileText, X, Plus, Trash2, Layers } from 'lucide-react'
import clsx from 'clsx'

const Sidebar = () => {
  const { 
    openFiles, 
    currentFilePath, 
    switchFile, 
    closeFile, 
    nestingLevel, 
    setNestingLevel,
    highlightColor,
    setHighlightColor
  } = useNoteStore()

  const colors = [
    { name: 'Yellow', value: '#ffeb3b' },
    { name: 'Green', value: '#ccff90' },
    { name: 'Blue', value: '#a1d9ff' },
    { name: 'Red', value: '#ff8a80' },
    { name: 'Purple', value: '#d1c4e9' },
  ]

  const handleLoadPdf = async () => {
    // @ts-ignore
    if (window.api && window.api.selectPdf) {
      // @ts-ignore
      const path = await window.api.selectPdf()
      if (path) {
        // @ts-ignore
        if (window.setPdfExternal) {
            // @ts-ignore
            window.setPdfExternal(path)
        }
      }
    }
  }

  const handleClearAll = () => {
    if (confirm("Close all files and clear all notes? This cannot be undone.")) {
        useNoteStore.setState({ openFiles: [], fileStates: {}, currentFilePath: null, pdfUrl: null });
    }
  }

  return (
    <div className="h-full flex flex-col bg-gray-900 text-gray-300 border-r border-gray-700 w-64 shrink-0">
      <div className="p-4 border-b border-gray-700 bg-gray-800 flex items-center justify-between">
        <h2 className="font-semibold text-xs uppercase tracking-widest truncate">
            Open Files
        </h2>
        <div className="flex gap-2">
            <button 
                onClick={handleClearAll}
                className="p-1 hover:bg-red-900/30 rounded text-gray-500 hover:text-red-400 transition-colors"
                title="Clear All"
            >
                <Trash2 size={14} />
            </button>
            <button 
                onClick={handleLoadPdf}
                className="p-1 hover:bg-gray-700 rounded text-gray-400 hover:text-white transition-colors"
                title="Load PDF"
            >
                <Plus size={16} />
            </button>
        </div>
      </div>

      <div className="flex-1 overflow-y-auto p-2">
        {openFiles.length === 0 ? (
            <div className="text-center mt-10 p-4">
                <p className="text-xs text-gray-500 mb-4">No files open.</p>
                <p className="text-[10px] text-gray-600">Drag and drop PDFs here or use the Load PDF button.</p>
            </div>
        ) : (
            <div className="space-y-0.5">
                {openFiles.map((file) => (
                    <div
                        key={file.path}
                        className={clsx(
                            "w-full flex items-center gap-2 px-3 py-1.5 rounded text-xs group transition-colors cursor-pointer",
                            currentFilePath === file.path 
                                ? "bg-blue-600 text-white" 
                                : "hover:bg-gray-800 text-gray-400 hover:text-gray-200"
                        )}
                        onClick={() => switchFile(file.path)}
                    >
                        <FileText size={14} className={currentFilePath === file.path ? "text-white" : "text-blue-400"} />
                        <span className="truncate flex-1" title={file.name}>{file.name}</span>
                        <button 
                            onClick={(e) => {
                                e.stopPropagation();
                                closeFile(file.path);
                            }}
                            className="opacity-0 group-hover:opacity-100 p-0.5 hover:bg-white/20 rounded"
                        >
                            <X size={12} />
                        </button>
                    </div>
                ))}
            </div>
        )}
      </div>

      <div className="p-4 border-t border-gray-700 bg-gray-800/50">
          <div className="flex items-center gap-2 mb-3">
              <Layers size={14} className="text-gray-500" />
              <span className="text-[10px] uppercase font-bold tracking-tighter text-gray-400">Nesting Level</span>
          </div>
          <div className="flex gap-1 p-1 bg-gray-900 rounded-lg">
              {[1, 2, 3].map(level => (
                  <button
                    key={level}
                    onClick={() => setNestingLevel(level)}
                    className={clsx(
                        "flex-1 py-1.5 rounded text-[10px] font-bold transition-all",
                        nestingLevel === level 
                            ? "bg-blue-600 text-white shadow-lg" 
                            : "text-gray-500 hover:bg-gray-800 hover:text-gray-300"
                    )}
                  >
                      LVL {level}
                  </button>
              ))}
          </div>
      </div>

      <div className="p-4 border-t border-gray-700 bg-gray-800/50">
          <div className="flex items-center gap-2 mb-3">
              <span className="text-[10px] uppercase font-bold tracking-tighter text-gray-400">Highlight Color</span>
          </div>
          <div className="flex gap-2 justify-between px-1">
              {colors.map(color => (
                  <button
                    key={color.value}
                    onClick={() => setHighlightColor(color.value)}
                    className={clsx(
                        "w-6 h-6 rounded-full border-2 transition-all",
                        highlightColor === color.value 
                            ? "border-white scale-110 shadow-lg" 
                            : "border-transparent hover:scale-105"
                    )}
                    style={{ backgroundColor: color.value }}
                    title={color.name}
                  />
              ))}
          </div>
      </div>
    </div>
  )
}

export default Sidebar