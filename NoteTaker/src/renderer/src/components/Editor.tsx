import React, { useEffect, useRef } from 'react'
import { useEditor, EditorContent } from '@tiptap/react'
import StarterKit from '@tiptap/starter-kit'
import Placeholder from '@tiptap/extension-placeholder'
import Image from '@tiptap/extension-image'
import { useNoteStore } from '../store/useNoteStore'
import { HighlightLink } from './HighlightExtension'
import TurndownService from 'turndown'

const Editor = () => {
  const { 
    editorActions, 
    clearActions, 
    jumpToHighlight, 
    pdfTitle, 
    hasHighlights, 
    setHasHighlights,
    currentFilePath,
    updateFileContent,
    fileStates
  } = useNoteStore()

  // Track the last file path to detect switches
  const previousFilePathRef = useRef<string | null>(null)

  const editor = useEditor({
    extensions: [
      StarterKit,
      Placeholder.configure({
        placeholder: 'Start taking notes...',
      }),
      HighlightLink.configure({
          HTMLAttributes: {
              class: 'cursor-pointer bg-blue-100 text-blue-800 text-xs px-1 rounded ml-1 align-super font-mono border border-blue-200 hover:bg-blue-200'
          }
      }),
      Image
    ],
    content: '',
    editorProps: {
      attributes: {
        class: 'prose prose-sm sm:prose lg:prose-lg xl:prose-2xl focus:outline-none max-w-none w-full text-left',
      },
      handleClick: (view, pos, event) => {
        const target = event.target as HTMLElement
        const id = target.getAttribute('data-highlight-id')
        if (id) {
           jumpToHighlight(id)
           return true
        }
        return false
      }
    },
    onUpdate: ({ editor }) => {
        if (currentFilePath) {
            updateFileContent(currentFilePath, editor.getHTML())
        }
    }
  })

  // Effect to load content when switching files
  useEffect(() => {
    if (editor && currentFilePath && currentFilePath !== previousFilePathRef.current) {
        const content = fileStates[currentFilePath]?.notesContent || ''
        
        // Only update if content is actually different to avoid cursor jumps on initial load?
        // Actually, for a file switch, we always want to swap content.
        console.log(`Editor: Switching to ${currentFilePath}, loading content...`)
        editor.commands.setContent(content)
        
        previousFilePathRef.current = currentFilePath
    }
  }, [currentFilePath, editor, fileStates])


  // Effect to process actions from the store
  useEffect(() => {
    if (!editor || !editorActions || !Array.isArray(editorActions) || (editorActions || []).length === 0) {
        if (editorActions && !Array.isArray(editorActions)) {
            console.error("Editor: editorActions is not an array:", editorActions);
        }
        return
    }

    console.log(`Editor: Processing ${(editorActions || []).length} actions`);

    (editorActions || []).forEach((action) => {
      if (!action) return;
      console.log(`Editor: Action type=${action.type}`);
      if (action.type === 'INSERT_TITLE') {
        editor
          .chain()
          .focus()
          .insertContent(`<h1>${action.payload}</h1>`)
          .run()
      } else if (action.type === 'INSERT_PDF_NAME') {
         editor
          .chain()
          .focus('end')
          .insertContent(`<ul><li><strong>${action.payload}</strong></li></ul>`)
          .run()
      } else if (action.type === 'INSERT_BULLET') {
        console.log(`Editor: Inserting bullet content: ${(action.payload || "").substring(0, 20)}...`);

        try {
            const targetLevel = action.level || useNoteStore.getState().nestingLevel || 1;
            
            // 1. Force focus to end of doc
            editor.commands.focus('end')
            
            // 2. If we are not in a list (e.g. cursor is after the </ul>), 
            // we must move INTO the last list item to split it.
            if (!editor.isActive('bulletList')) {
                editor.commands.selectNodeBackward()
            }

            // 3. Perform the chain:
            // - Split (creates new item at same level as PDF Name or previous highlight)
            // - Lift everything to root (Level 0)
            // - Sink to exactly targetLevel
            editor.chain()
                .splitListItem('listItem')
                .liftListItem('listItem') // Lift as much as possible
                .liftListItem('listItem')
                .liftListItem('listItem')
                .run()
                
            // Now sink to target
            for (let i = 0; i < targetLevel; i++) {
                editor.commands.sinkListItem('listItem')
            }

            // 4. Insert the actual content
            editor.chain().insertContent(action.payload).run()

            // Insert the reference link node if ID exists
            if (action.highlightId) {
                console.log(`Editor: Inserting highlight link node for ID: ${action.highlightId}`);
                editor.commands.setHighlightLink({ id: action.highlightId })
            }

        } catch (err) {
            console.error("Editor: Error during insertion:", err);
            // Fallback: just append text
            editor.commands.insertContent(action.payload)
        }

        if (!hasHighlights) {
            setHasHighlights(true)
        }
      } else if (action.type === 'INSERT_IMAGE') {
        editor.chain().focus().setImage({ src: action.payload }).run()
      }
    })

    clearActions()
  }, [editorActions, editor, clearActions])

  const handleSave = async () => {
    if (!editor) return
    const html = editor.getHTML()
    const turndownService = new TurndownService()
    const markdown = turndownService.turndown(html)
    
    // @ts-ignore - Defined in d.ts
    if (window.api && window.api.saveMarkdown) {
        await window.api.saveMarkdown(markdown, pdfTitle ? `${pdfTitle}.md` : 'notes.md')
    } else {
        console.warn("Editor: Save markdown API not available")
    }
  }

  if (!editor) {
    return null
  }

  return (
    <div className="h-full w-full overflow-y-auto bg-white flex flex-col">
       <div className="border-b p-2 flex gap-2 bg-gray-50 items-center justify-between sticky top-0 z-10">
        <div className="flex gap-2">
            <button onClick={() => editor.chain().focus().toggleBold().run()} className="px-2 py-1 border rounded hover:bg-gray-200 font-bold text-sm">B</button>
            <button onClick={() => editor.chain().focus().toggleItalic().run()} className="px-2 py-1 border rounded hover:bg-gray-200 italic text-sm">I</button>
            <button onClick={() => editor.chain().focus().toggleBulletList().run()} className="px-2 py-1 border rounded hover:bg-gray-200 text-sm">List</button>
            <button onClick={() => editor.chain().focus().clearContent().run()} className="px-2 py-1 border rounded hover:bg-red-50 text-red-600 text-sm border-red-200">Clear</button>
        </div>
        <button onClick={handleSave} className="px-3 py-1 bg-blue-600 text-white rounded hover:bg-blue-700 text-sm font-medium">
            Save Notes
        </button>
      </div>
      <div className="p-4 flex-1">
          <EditorContent editor={editor} className="min-h-[500px]" />
      </div>
    </div>
  )
}

export default Editor