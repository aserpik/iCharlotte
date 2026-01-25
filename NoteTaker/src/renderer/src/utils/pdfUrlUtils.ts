export function constructLocalPdfUrl(pdfUrl: string | null): string | null {
  if (!pdfUrl) return null;
  
  let finalPdfUrl = pdfUrl;
  
  if (typeof pdfUrl === 'string' && !pdfUrl.startsWith('blob:') && !pdfUrl.startsWith('http')) {
      // Extract query parameters (like cache busters) before processing path
      let cleanUrl = pdfUrl || "";
      let query = "";
      const queryIdx = cleanUrl.indexOf('?');
      if (queryIdx !== -1) {
          query = cleanUrl.substring(queryIdx);
          cleanUrl = cleanUrl.substring(0, queryIdx);
      }

      const normalizedPath = cleanUrl.replace(/\\/g, '/');
      if (normalizedPath) {
          const segments = normalizedPath.split('/').filter(s => !!s)
          const safePath = segments.map(s => encodeURIComponent(s)).join('/')
          finalPdfUrl = `local-resource:///${safePath}${query}`
      }
  }
  
  return finalPdfUrl;
}
