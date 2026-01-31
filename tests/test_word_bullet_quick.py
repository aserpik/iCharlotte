"""Quick test of the updated bullet formatting."""

import sys
sys.path.insert(0, 'C:\\geminiterminal2')

import win32com.client
import pythoncom

def main():
    pythoncom.CoInitialize()

    try:
        word = win32com.client.GetActiveObject("Word.Application")
        print("Connected to existing Word instance")
    except:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True
        print("Created new Word instance")

    # Create new document
    doc = word.Documents.Add()
    selection = word.Selection

    # Test the new approach
    selection.TypeText("Testing updated bullet formatting:")
    selection.TypeParagraph()
    selection.TypeParagraph()

    # Insert bullets using the new method
    test_items = [
        "First bullet point - this is a long text to test wrapping and see if the wrapped lines align properly with the first line of text after the bullet.",
        "Second bullet point - another long text to verify that the indentation is working correctly and wrapped text stays aligned.",
        "Third bullet point - short text.",
    ]

    first_bullet = True
    for i, item_text in enumerate(test_items):
        if first_bullet:
            # Only apply bullet format for the FIRST bullet
            # Subsequent bullets are created by Word's list continuation
            selection.Range.ListFormat.ApplyBulletDefault()

            # Modify the ListTemplate level settings
            try:
                list_fmt = selection.Range.ListFormat
                if list_fmt.ListTemplate and list_fmt.ListLevelNumber > 0:
                    level = list_fmt.ListTemplate.ListLevels(list_fmt.ListLevelNumber)
                    level.NumberPosition = 36  # 0.5 inch - where bullet sits
                    level.TextPosition = 54    # 0.75 inch - where text starts
                    level.TabPosition = 54     # 0.75 inch - tab stop
                    print(f"Set list template: NumberPos=36, TextPos=54, TabPos=54")
            except Exception as e:
                print(f"Error: {e}")

            first_bullet = False

        selection.TypeText(item_text)

        if i < len(test_items) - 1:
            selection.TypeParagraph()  # Word continues the list automatically

    # End bullet list
    selection.TypeParagraph()
    selection.Range.ListFormat.RemoveNumbers()

    print("\nDone! Check the Word document.")
    print("- Bullet should be at 0.5 inch")
    print("- Text should start at 0.75 inch")
    print("- Wrapped text should align with first line of text")
    print("- All 3 bullets should be present")

if __name__ == "__main__":
    main()
