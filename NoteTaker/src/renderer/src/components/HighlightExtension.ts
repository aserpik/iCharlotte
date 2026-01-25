import { Node, mergeAttributes } from '@tiptap/core'

export interface HighlightAttributes {
  id: string
  label?: string
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    highlightLink: {
      setHighlightLink: (attributes: HighlightAttributes) => ReturnType
    }
  }
}

export const HighlightLink = Node.create<HighlightAttributes>({
  name: 'highlightLink',

  group: 'inline',
  inline: true,
  selectable: true,
  atom: true,

  addAttributes() {
    return {
      id: {
        default: null,
        parseHTML: (element) => element.getAttribute('data-highlight-id'),
        renderHTML: (attributes) => {
          if (!attributes.id) return {}
          return { 'data-highlight-id': attributes.id }
        },
      },
      label: {
        default: '[ref]',
        parseHTML: (element) => element.textContent,
        renderHTML: (attributes) => {
           return {}
        }
      }
    }
  },

  parseHTML() {
    return [
      {
        tag: 'span[data-highlight-id]',
      },
    ]
  },

  renderHTML({ HTMLAttributes, node }) {
    return [
      'span', 
      mergeAttributes(HTMLAttributes, {
        class: 'cursor-pointer bg-blue-100 text-blue-800 text-xs px-1 rounded ml-1 align-super font-mono border border-blue-200 hover:bg-blue-200'
      }), 
      node.attrs.label || '[ref]'
    ]
  },

  addCommands() {
    return {
      setHighlightLink:
        (attributes) =>
        ({ commands }) => {
          return commands.insertContent({
            type: this.name,
            attrs: attributes,
          })
        },
    }
  },
})