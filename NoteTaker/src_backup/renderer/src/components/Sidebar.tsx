import React, { useState, useEffect } from 'react'
import { useNoteStore } from '../store/useNoteStore'
import { FolderOpen, FileText, File, Folder, ChevronLeft } from 'lucide-react'
import clsx from 'clsx'

const Sidebar = () => {
  const { currentWorkspacePath, setWorkspace, setPdf, currentFilePath, pushAction } = useNoteStore()
  const [currentPath, setCurrentPath] = useState<string | null>(null)
  const [files, setFiles] = useState<Array<{ name: string; isDirectory: boolean; path: string; size?: number; mtime?: number }>>([])
  const [activeTab, setActiveTab] = useState<'workspace' | 'case'>('workspace')
  const [sortConfig, setSortConfig] = useState<{ key: 'name' | 'size' | 'mtime'; direction: 'asc' | 'desc' }>({
    key: 'name',
    direction: 'asc'
  })

  const handleOpenFolder = async () => {
    // @ts-ignore
    if (window.api && window.api.selectDirectory) {
      // @ts-ignore
      const path = await window.api.selectDirectory()
      if (path) {
        setWorkspace(path)
        setCurrentPath(path)
      }
    } else {
        console.error("API not ready")
    }
  }

  const refreshFiles = () => {
    if (currentPath) {
      // @ts-ignore
      if (window.api && window.api.listDirectory) {
          // @ts-ignore
          window.api.listDirectory(currentPath).then((allFiles) => {
            console.log("Sidebar: listDirectory result:", allFiles);
            if (!allFiles || !Array.isArray(allFiles)) {
                console.warn("Sidebar: listDirectory returned invalid result:", allFiles);
                setFiles([]);
                return;
            }
            // Filter PDFs and Directories
            const filtered = allFiles.filter((f) => f && (f.isDirectory || (f.name && f.name.toLowerCase().endsWith('.pdf'))))
            setFiles(filtered)
          }).catch(err => {
              console.error("Sidebar: listDirectory failed:", err);
              setFiles([]);
          })
      }
    }
  }

  useEffect(() => {
    if (currentWorkspacePath && !currentPath) {
        setCurrentPath(currentWorkspacePath)
    }
  }, [currentWorkspacePath])

  useEffect(() => {
    refreshFiles()
  }, [currentPath])

  const handleFileClick = (file: { name: string; path: string; isDirectory: boolean }) => {
    if (file.isDirectory) {
        setCurrentPath(file.path)
    } else {
        setPdf(file.path, file.name, file.path)
    }
  }

  const handleGoUp = () => {
    if (currentPath && currentPath !== currentWorkspacePath) {
        // Simple way to go up one level
        const parts = currentPath.split(/[\\\/]/)
        if (parts && Array.isArray(parts) && parts.length > 1) {
            parts.pop()
            setCurrentPath(parts.join('/'))
        }
    }
  }

  const handleDragStart = (e: React.DragEvent, file: any) => {
    if (file.isDirectory) return
    e.dataTransfer.setData('application/json', JSON.stringify(file))
    e.dataTransfer.effectAllowed = 'move'
  }

  const handleDragOver = (e: React.DragEvent, folder: any) => {
    if (!folder.isDirectory) return
    e.preventDefault()
    e.dataTransfer.dropEffect = 'move'
  }

  const handleDrop = async (e: React.DragEvent, targetFolder: any) => {
    if (!targetFolder.isDirectory) return
    e.preventDefault()
    const data = e.dataTransfer.getData('application/json')
    if (!data) return

    const draggedFile = JSON.parse(data)
    
    // Construct target path
    let targetPath = `${targetFolder.path}/${draggedFile.name}`.replace(/\\/g, '/')
    
    // Ensure Windows paths are correct (remove leading slash if it exists and looks like /C:/...)
    if (targetPath.match(/^\/[a-zA-Z]:/)) {
        targetPath = targetPath.substring(1)
    }

    let sourcePath = draggedFile.path.replace(/\\/g, '/')
    if (sourcePath.match(/^\/[a-zA-Z]:/)) {
        sourcePath = sourcePath.substring(1)
    }
    
    // @ts-ignore
    const success = await window.api.moveFile(sourcePath, targetPath)
    if (success) {
        refreshFiles()
    }
  }

  const formatSize = (bytes?: number) => {
    if (bytes === undefined) return '-'
    if (bytes === 0) return '0 B'
    const k = 1024
    const sizes = ['B', 'KB', 'MB', 'GB', 'TB']
    const i = Math.floor(Math.log(bytes) / Math.log(k))
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i]
  }

  const formatDate = (mtime?: number) => {
    if (!mtime) return '-'
    return new Date(mtime).toLocaleDateString()
  }

  const sortedFiles = (files || []).slice().sort((a, b) => {
    // Directories always first
    if (a.isDirectory && !b.isDirectory) return -1
    if (!a.isDirectory && b.isDirectory) return 1

    const key = sortConfig.key
    const aValue = a[key] ?? 0
    const bValue = b[key] ?? 0
    
    if (aValue < bValue) return sortConfig.direction === 'asc' ? -1 : 1
    if (aValue > bValue) return sortConfig.direction === 'asc' ? 1 : -1
    return 0
  })

  const requestSort = (key: 'name' | 'size' | 'mtime') => {
    let direction: 'asc' | 'desc' = 'asc'
    if (sortConfig.key === key && sortConfig.direction === 'asc') {
      direction = 'desc'
    }
    setSortConfig({ key, direction })
  }

  return (
    <div className="h-full flex flex-col bg-gray-900 text-gray-300 border-r border-gray-700 w-80">
      <div className="p-4 border-b border-gray-700 flex items-center justify-between">
        <h2 className="font-semibold text-sm uppercase tracking-wider truncate" title={currentPath || 'Note Taker'}>
            {currentPath ? currentPath.split(/[\\\/]/).pop() : 'Note Taker'}
        </h2>
        <div className="flex items-center gap-2">
            {currentPath !== currentWorkspacePath && currentPath && (
                <button onClick={handleGoUp} className="hover:text-white" title="Go Up">
                    <ChevronLeft size={18} />
                </button>
            )}
            <button onClick={handleOpenFolder} className="hover:text-white" title="Open Folder">
                <FolderOpen size={18} />
            </button>
        </div>
      </div>

      <div className="flex border-b border-gray-700">
        <button 
            onClick={() => setActiveTab('workspace')}
            className={clsx(
                "flex-1 py-2 text-xs font-medium uppercase tracking-wider transition-colors",
                activeTab === 'workspace' ? "bg-gray-800 text-white border-b-2 border-blue-500" : "hover:bg-gray-800"
            )}
        >
            Workspace
        </button>
        <button 
            onClick={() => setActiveTab('case')}
            className={clsx(
                "flex-1 py-2 text-xs font-medium uppercase tracking-wider transition-colors",
                activeTab === 'case' ? "bg-gray-800 text-white border-b-2 border-blue-500" : "hover:bg-gray-800"
            )}
        >
            Case View
        </button>
      </div>
      
      <div className="flex-1 overflow-y-auto">
        {!currentWorkspacePath ? (
            <div className="text-center mt-10 text-gray-500 text-sm">
                <p>No folder open.</p>
                <button onClick={handleOpenFolder} className="mt-2 text-blue-400 hover:underline">Open Folder</button>
            </div>
        ) : activeTab === 'workspace' ? (
            <div className="p-2 space-y-1">
                {sortedFiles.map((file) => (
                    <button
                        key={file.path}
                        onClick={() => handleFileClick(file)}
                        draggable={!file.isDirectory}
                        onDragStart={(e) => handleDragStart(e, file)}
                        onDragOver={(e) => handleDragOver(e, file)}
                        onDrop={(e) => handleDrop(e, file)}
                        className={clsx(
                            "w-full text-left px-3 py-2 rounded text-sm flex items-center gap-2 truncate",
                            currentFilePath === file.path 
                                ? "bg-blue-600 text-white" 
                                : "hover:bg-gray-800",
                            file.isDirectory ? "text-yellow-500" : ""
                        )}
                    >
                        {file.isDirectory ? <Folder size={16} className="shrink-0" /> : <FileText size={16} className="shrink-0" />}
                        <span className="truncate">{file.name}</span>
                    </button>
                ))}
                {(files || []).length === 0 && <p className="text-gray-500 text-sm px-3 pt-2">No PDFs or folders found.</p>}
            </div>
        ) : (
            <div className="p-0">
                <table className="w-full text-xs text-left border-collapse">
                    <thead className="sticky top-0 bg-gray-900 shadow-sm">
                        <tr>
                            <th 
                                className="px-2 py-2 cursor-pointer hover:bg-gray-800"
                                onClick={() => requestSort('name')}
                            >
                                Name {sortConfig.key === 'name' && (sortConfig.direction === 'asc' ? '↑' : '↓')}
                            </th>
                            <th 
                                className="px-2 py-2 cursor-pointer hover:bg-gray-800 w-16"
                                onClick={() => requestSort('mtime')}
                            >
                                Date {sortConfig.key === 'mtime' && (sortConfig.direction === 'asc' ? '↑' : '↓')}
                            </th>
                            <th 
                                className="px-2 py-2 cursor-pointer hover:bg-gray-800 w-16"
                                onClick={() => requestSort('size')}
                            >
                                Size {sortConfig.key === 'size' && (sortConfig.direction === 'asc' ? '↑' : '↓')}
                            </th>
                        </tr>
                    </thead>
                    <tbody>
                        {sortedFiles.map((file) => (
                            <tr 
                                key={file.path}
                                onClick={() => handleFileClick(file)}
                                draggable={!file.isDirectory}
                                onDragStart={(e) => handleDragStart(e, file)}
                                onDragOver={(e) => handleDragOver(e, file)}
                                onDrop={(e) => handleDrop(e, file)}
                                className={clsx(
                                    "cursor-pointer border-b border-gray-800 group",
                                    currentFilePath === file.path ? "bg-blue-900/50" : "hover:bg-gray-800"
                                )}
                            >
                                <td className="px-2 py-2 truncate max-w-[120px]" title={file.name}>
                                    <div className="flex items-center gap-1">
                                        {file.isDirectory ? (
                                            <Folder size={12} className="shrink-0 text-yellow-500" />
                                        ) : (
                                            <FileText size={12} className="shrink-0 text-gray-500 group-hover:text-gray-300" />
                                        )}
                                        <span className={clsx("truncate", file.isDirectory ? "text-yellow-500 font-medium" : "")}>
                                            {file.name}
                                        </span>
                                    </div>
                                </td>
                                <td className="px-2 py-2 whitespace-nowrap text-gray-500">
                                    {formatDate(file.mtime)}
                                </td>
                                <td className="px-2 py-2 whitespace-nowrap text-gray-500 text-right">
                                    {formatSize(file.size)}
                                </td>
                            </tr>
                        ))}
                    </tbody>
                </table>
                {(files || []).length === 0 && <p className="text-gray-500 text-sm px-3 pt-4">No PDFs found.</p>}
            </div>
        )}
      </div>
    </div>
  )
}

export default Sidebar
