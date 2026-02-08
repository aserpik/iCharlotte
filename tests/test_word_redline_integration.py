"""Integration tests for Word redline functionality.

These tests require Word to be installed and may create temporary documents.
They may fail or be skipped if run in non-interactive environments or if
Word COM automation encounters issues.

Run with: pytest tests/test_word_redline_integration.py -v

Note: These tests are primarily for manual verification in development
environments where Word is installed and COM automation is working.
"""

import unittest
import tempfile
import os

try:
    import win32com.client
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

@unittest.skipUnless(HAS_WIN32, "Requires win32com (Word) to be available")
class TestWordRedlineIntegration(unittest.TestCase):

    def setUp(self):
        """Create Word instance and temporary document."""
        self.word = None
        self.doc = None
        try:
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False
            # Add blank document
            self.doc = self.word.Documents.Add()
            # Force early binding by accessing Content
            _ = self.doc.Content
        except Exception as e:
            self.skipTest(f"Failed to initialize Word: {e}")

    def tearDown(self):
        """Close document and Word."""
        try:
            if self.doc:
                self.doc.Close(SaveChanges=False)
            if self.word:
                self.word.Quit()
        except:
            pass

    def test_track_changes_enabled(self):
        """Test that Track Changes can be enabled programmatically."""
        # Ensure document is valid
        self.assertIsNotNone(self.doc)

        # Test Track Changes property
        try:
            initial_state = self.doc.TrackRevisions
            self.doc.TrackRevisions = True
            self.assertTrue(self.doc.TrackRevisions)
            self.doc.TrackRevisions = False
            self.assertFalse(self.doc.TrackRevisions)
        except AttributeError as e:
            self.skipTest(f"TrackRevisions not available: {e}")

    def test_range_reconstruction(self):
        """Test reconstructing a range from start/end coordinates."""
        # Insert some text
        self.doc.Content.Text = "Hello World"

        # Get range
        range_obj = self.doc.Range(0, 5)
        start = range_obj.Start
        end = range_obj.End

        # Reconstruct
        new_range = self.doc.Range(start, end)
        self.assertEqual(new_range.Text, "Hello")

if __name__ == '__main__':
    unittest.main()
