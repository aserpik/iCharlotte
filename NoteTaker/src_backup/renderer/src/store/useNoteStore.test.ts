import { describe, it, expect, beforeEach } from 'vitest'
import { useNoteStore } from './useNoteStore'

describe('NoteStore', () => {
  beforeEach(() => {
    useNoteStore.setState({
      highlights: [],
      editorActions: [],
      pdfUrl: null,
      pdfTitle: ''
    })
  })

  it('should add a highlight', () => {
    const highlight = { id: '1', content: { text: 'test' } }
    useNoteStore.getState().addHighlight(highlight)
    expect(useNoteStore.getState().highlights).toHaveLength(1)
    expect(useNoteStore.getState().highlights[0].id).toBe('1')
  })

  it('should push and clear editor actions', () => {
    const action = { type: 'INSERT_TITLE' as const, payload: 'My PDF' }
    useNoteStore.getState().pushAction(action)
    
    expect(useNoteStore.getState().editorActions).toHaveLength(1)
    expect(useNoteStore.getState().editorActions[0].payload).toBe('My PDF')

    useNoteStore.getState().clearActions()
    expect(useNoteStore.getState().editorActions).toHaveLength(0)
  })

  it('should set PDF info', () => {
    useNoteStore.getState().setPdf('file://test.pdf', 'test.pdf', '/path/test.pdf')
    expect(useNoteStore.getState().pdfTitle).toBe('test.pdf')
    expect(useNoteStore.getState().currentFilePath).toBe('/path/test.pdf')
  })
})
