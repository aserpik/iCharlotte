import React from 'react'
import ReactDOM from 'react-dom/client'
import './index.css'
import 'react-pdf-highlighter/dist/style.css'
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