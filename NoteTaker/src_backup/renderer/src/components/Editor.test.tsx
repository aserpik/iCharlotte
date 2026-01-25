import { describe, it, expect, vi, beforeEach } from 'vitest'
import { render, screen, waitFor } from '@testing-library/react'
import { act } from 'react'
import Editor from './Editor'
import { useNoteStore } from '../store/useNoteStore'
import React from 'react'

describe('Editor Component', () => {
  beforeEach(() => {
    useNoteStore.setState({
      editorActions: [],
      pdfTitle: 'Test PDF'
    })
    vi.clearAllMocks()
  })

  it('renders placeholder when empty', () => {
    const { container } = render(<Editor />)
    const paragraph = container.querySelector('p.is-editor-empty')
    expect(paragraph?.getAttribute('data-placeholder')).toBe('Start taking notes...')
  })

  it('responds to INSERT_TITLE action', async () => {
    render(<Editor />)
    
    await act(async () => {
        useNoteStore.getState().pushAction({ type: 'INSERT_TITLE', payload: 'New Title' })
    })
    
    await waitFor(() => {
      const h1 = document.querySelector('h1')
      expect(h1?.textContent).toBe('New Title')
    })
  })

  it('responds to INSERT_BULLET action', async () => {
    render(<Editor />)
    
    await act(async () => {
        useNoteStore.getState().pushAction({ 
            type: 'INSERT_BULLET', 
            payload: 'Highlighted Text',
            highlightId: 'h1'
        })
    })
    
    await waitFor(() => {
      expect(screen.getByText(/Highlighted Text/i)).toBeDefined()
      expect(screen.getByText(/\[ref\]/i)).toBeDefined()
    })
  })

  it('clears editor when Clear button is clicked', async () => {
    render(<Editor />)
    
    await act(async () => {
        useNoteStore.getState().pushAction({ type: 'INSERT_TITLE', payload: 'To be cleared' })
    })
    
    await waitFor(() => {
       expect(screen.getByText('To be cleared')).toBeDefined()
    })
    
    const clearBtn = screen.getByText('Clear')
    await act(async () => {
        clearBtn.click()
    })
    
    await waitFor(() => {
       expect(screen.queryByText('To be cleared')).toBeNull()
    })
  })
})
