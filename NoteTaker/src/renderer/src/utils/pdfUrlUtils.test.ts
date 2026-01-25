import { describe, it, expect } from 'vitest';
import { constructLocalPdfUrl } from './pdfUrlUtils';

describe('constructLocalPdfUrl', () => {
  it('should handle standard windows paths', () => {
    const input = 'C:\\Users\\Test\\Doc.pdf';
    const expected = 'local-resource:///C%3A/Users/Test/Doc.pdf';
    expect(constructLocalPdfUrl(input)).toBe(expected);
  });

  it('should handle paths with spaces', () => {
    const input = 'C:\\Users\\Test\\My Doc.pdf';
    const expected = 'local-resource:///C%3A/Users/Test/My%20Doc.pdf';
    expect(constructLocalPdfUrl(input)).toBe(expected);
  });

  it('should preserve query parameters (cache busters)', () => {
    const input = 'C:\\Users\\Test\\Doc.pdf?t=12345';
    // The previous buggy behavior would have encoded ? as %3F and = as %3D
    const expected = 'local-resource:///C%3A/Users/Test/Doc.pdf?t=12345';
    expect(constructLocalPdfUrl(input)).toBe(expected);
  });

  it('should ignore blob URLs', () => {
    const input = 'blob:http://localhost/123';
    expect(constructLocalPdfUrl(input)).toBe(input);
  });

  it('should ignore http URLs', () => {
    const input = 'http://example.com/doc.pdf';
    expect(constructLocalPdfUrl(input)).toBe(input);
  });
});
