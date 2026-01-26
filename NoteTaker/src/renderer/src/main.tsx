import React from 'react'
import ReactDOM from 'react-dom/client'
import './index.css'
import 'react-pdf-highlighter/dist/style/AreaHighlight.css'
import 'react-pdf-highlighter/dist/style/Highlight.css'
import 'react-pdf-highlighter/dist/style/MouseSelection.css'
import 'react-pdf-highlighter/dist/style/pdf_viewer.css'
import 'react-pdf-highlighter/dist/style/PdfHighlighter.css'
import 'react-pdf-highlighter/dist/style/Tip.css'
import App from './App'

// Global error handling for better debugging
window.addEventListener('error', (event) => {
    console.error('GLOBAL ERROR CAUGHT:', {
        message: event.message,
        filename: event.filename,
        lineno: event.lineno,
        colno: event.colno,
        error: event.error?.stack || event.error
    });
});

window.addEventListener('unhandledrejection', (event) => {
    console.error('UNHANDLED REJECTION:', event.reason);
});

ReactDOM.createRoot(document.getElementById('root') as HTMLElement).render(
    <App />
)