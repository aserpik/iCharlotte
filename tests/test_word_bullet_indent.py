"""
Test script to diagnose Word bullet point indentation issues.
This script will insert bullet points and report the actual indentation values.
"""

import sys
import time
sys.path.insert(0, 'C:\\geminiterminal2')

try:
    import win32com.client
    import pythoncom
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False
    print("ERROR: win32com not available")
    sys.exit(1)


def get_word_app():
    """Get or create Word application."""
    pythoncom.CoInitialize()
    try:
        word = win32com.client.GetActiveObject("Word.Application")
        print("Connected to existing Word instance")
    except:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True
        print("Created new Word instance")
    return word


def report_paragraph_format(selection, label=""):
    """Report the current paragraph formatting."""
    para = selection.ParagraphFormat
    print(f"\n=== Paragraph Format {label} ===")
    print(f"  LeftIndent: {para.LeftIndent} points ({para.LeftIndent/72:.3f} inches)")
    print(f"  FirstLineIndent: {para.FirstLineIndent} points ({para.FirstLineIndent/72:.3f} inches)")
    print(f"  RightIndent: {para.RightIndent} points")

    # List format info
    list_fmt = selection.Range.ListFormat
    print(f"  ListLevelNumber: {list_fmt.ListLevelNumber}")
    print(f"  ListType: {list_fmt.ListType}")  # 0=none, 2=bullet, 3=numbered

    try:
        if list_fmt.ListTemplate:
            lt = list_fmt.ListTemplate
            print(f"  ListTemplate exists")
            # Try to get list level info
            if list_fmt.ListLevelNumber > 0:
                level = lt.ListLevels(list_fmt.ListLevelNumber)
                print(f"    Level TextPosition: {level.TextPosition} points ({level.TextPosition/72:.3f} inches)")
                print(f"    Level NumberPosition: {level.NumberPosition} points ({level.NumberPosition/72:.3f} inches)")
                print(f"    Level TabPosition: {level.TabPosition} points ({level.TabPosition/72:.3f} inches)")
    except Exception as e:
        print(f"  ListTemplate: Error getting details - {e}")


def test_method_1_apply_bullet_then_indent(word, doc):
    """Test: Apply bullet default, then set indentation."""
    print("\n" + "="*60)
    print("METHOD 1: ApplyBulletDefault() THEN set indentation")
    print("="*60)

    selection = word.Selection

    # Apply bullet first
    selection.Range.ListFormat.ApplyBulletDefault()
    print("Applied ApplyBulletDefault()")
    report_paragraph_format(selection, "after ApplyBulletDefault")

    # Now set indentation
    para = selection.ParagraphFormat
    para.LeftIndent = 54  # 0.75 inch
    para.FirstLineIndent = -18  # -0.25 inch
    print("Set LeftIndent=54, FirstLineIndent=-18")
    report_paragraph_format(selection, "after setting indent")

    # Type text
    selection.TypeText("Method 1: Bullet then indent - This is a long bullet point to test text wrapping and see if the wrapped lines align properly with the first line of text.")
    report_paragraph_format(selection, "after typing text")

    selection.TypeParagraph()
    selection.Range.ListFormat.RemoveNumbers()
    selection.TypeParagraph()


def test_method_2_indent_then_bullet(word, doc):
    """Test: Set indentation, then apply bullet."""
    print("\n" + "="*60)
    print("METHOD 2: Set indentation THEN ApplyBulletDefault()")
    print("="*60)

    selection = word.Selection

    # Set indentation first
    para = selection.ParagraphFormat
    para.LeftIndent = 54
    para.FirstLineIndent = -18
    print("Set LeftIndent=54, FirstLineIndent=-18")
    report_paragraph_format(selection, "after setting indent")

    # Apply bullet
    selection.Range.ListFormat.ApplyBulletDefault()
    print("Applied ApplyBulletDefault()")
    report_paragraph_format(selection, "after ApplyBulletDefault")

    # Type text
    selection.TypeText("Method 2: Indent then bullet - This is a long bullet point to test text wrapping and see if the wrapped lines align properly with the first line of text.")
    report_paragraph_format(selection, "after typing text")

    selection.TypeParagraph()
    selection.Range.ListFormat.RemoveNumbers()
    selection.TypeParagraph()


def test_method_3_indent_bullet_indent(word, doc):
    """Test: Set indent, apply bullet, set indent again."""
    print("\n" + "="*60)
    print("METHOD 3: Set indent, ApplyBulletDefault(), set indent again")
    print("="*60)

    selection = word.Selection

    # Set indentation first
    para = selection.ParagraphFormat
    para.LeftIndent = 54
    para.FirstLineIndent = -18
    print("Set LeftIndent=54, FirstLineIndent=-18")

    # Apply bullet
    selection.Range.ListFormat.ApplyBulletDefault()
    print("Applied ApplyBulletDefault()")
    report_paragraph_format(selection, "after ApplyBulletDefault")

    # Set indentation again
    para.LeftIndent = 54
    para.FirstLineIndent = -18
    print("Set LeftIndent=54, FirstLineIndent=-18 again")
    report_paragraph_format(selection, "after re-setting indent")

    # Type text
    selection.TypeText("Method 3: Indent-bullet-indent - This is a long bullet point to test text wrapping and see if the wrapped lines align properly with the first line of text.")
    report_paragraph_format(selection, "after typing text")

    selection.TypeParagraph()
    selection.Range.ListFormat.RemoveNumbers()
    selection.TypeParagraph()


def test_method_4_manual_bullet(word, doc):
    """Test: Manual bullet character with tab (no Word list formatting)."""
    print("\n" + "="*60)
    print("METHOD 4: Manual bullet character (no ApplyBulletDefault)")
    print("="*60)

    selection = word.Selection

    # Set indentation
    para = selection.ParagraphFormat
    para.LeftIndent = 54  # 0.75 inch - where text starts
    para.FirstLineIndent = -18  # -0.25 inch hanging
    print("Set LeftIndent=54, FirstLineIndent=-18")
    report_paragraph_format(selection, "after setting indent")

    # Type bullet + tab + text manually
    selection.TypeText("•\tMethod 4: Manual bullet - This is a long bullet point to test text wrapping and see if the wrapped lines align properly with the first line of text.")
    report_paragraph_format(selection, "after typing text")

    selection.TypeParagraph()
    # Reset for next test
    para.LeftIndent = 0
    para.FirstLineIndent = 0
    selection.TypeParagraph()


def test_method_5_modify_list_template(word, doc):
    """Test: Apply bullet, then modify the list template levels."""
    print("\n" + "="*60)
    print("METHOD 5: ApplyBulletDefault then modify ListTemplate levels")
    print("="*60)

    selection = word.Selection

    # Apply bullet
    selection.Range.ListFormat.ApplyBulletDefault()
    print("Applied ApplyBulletDefault()")
    report_paragraph_format(selection, "after ApplyBulletDefault")

    # Try to modify the list template
    try:
        list_fmt = selection.Range.ListFormat
        if list_fmt.ListTemplate and list_fmt.ListLevelNumber > 0:
            level = list_fmt.ListTemplate.ListLevels(list_fmt.ListLevelNumber)
            print(f"Current level settings:")
            print(f"  NumberPosition: {level.NumberPosition}")
            print(f"  TextPosition: {level.TextPosition}")
            print(f"  TabPosition: {level.TabPosition}")

            # Modify the level
            level.NumberPosition = 36  # 0.5 inch - where bullet sits
            level.TextPosition = 54    # 0.75 inch - where text starts
            level.TabPosition = 54     # 0.75 inch - tab stop
            print("Modified level: NumberPosition=36, TextPosition=54, TabPosition=54")
            report_paragraph_format(selection, "after modifying list level")
    except Exception as e:
        print(f"Error modifying list template: {e}")

    # Type text
    selection.TypeText("Method 5: Modified list template - This is a long bullet point to test text wrapping and see if the wrapped lines align properly with the first line of text.")
    report_paragraph_format(selection, "after typing text")

    selection.TypeParagraph()
    selection.Range.ListFormat.RemoveNumbers()
    selection.TypeParagraph()


def test_method_6_list_indent_method(word, doc):
    """Test: Use ListFormat.ListIndent method."""
    print("\n" + "="*60)
    print("METHOD 6: ApplyBulletDefault then use ListIndent")
    print("="*60)

    selection = word.Selection

    # Apply bullet
    selection.Range.ListFormat.ApplyBulletDefault()
    print("Applied ApplyBulletDefault()")
    report_paragraph_format(selection, "after ApplyBulletDefault")

    # Try ListIndent - this increases indent level
    # Actually we want to set specific values, not increase
    # Let's try ListOutdent first to reset, then ListIndent
    try:
        # This might help adjust the indent
        pass  # ListIndent/ListOutdent change levels, not exact positions
    except Exception as e:
        print(f"Error: {e}")

    # Type text
    selection.TypeText("Method 6: ListIndent test - This is a long bullet point.")

    selection.TypeParagraph()
    selection.Range.ListFormat.RemoveNumbers()
    selection.TypeParagraph()


def test_method_7_apply_list_template_with_level(word, doc):
    """Test: Create custom list template and apply it."""
    print("\n" + "="*60)
    print("METHOD 7: Create and apply custom ListTemplate")
    print("="*60)

    selection = word.Selection

    try:
        # Get the ListGalleries (1=bullet, 2=numbered, 3=outline)
        bullet_gallery = word.ListGalleries(1)  # wdBulletGallery

        # Get or modify a template from the gallery
        # ListTemplates are 1-indexed, there are usually 7
        template = bullet_gallery.ListTemplates(1)

        # Modify level 1
        level = template.ListLevels(1)
        level.NumberPosition = 36  # 0.5 inch - bullet position
        level.TextPosition = 54    # 0.75 inch - text position
        level.TabPosition = 54     # tab stop

        print("Created custom template with NumberPosition=36, TextPosition=54")

        # Apply this template
        selection.Range.ListFormat.ApplyListTemplateWithLevel(
            ListTemplate=template,
            ContinuePreviousList=False,
            ApplyTo=0,  # wdListApplyToWholeList
            DefaultListBehavior=1  # wdWord10ListBehavior
        )

        report_paragraph_format(selection, "after ApplyListTemplateWithLevel")

    except Exception as e:
        print(f"Error creating custom template: {e}")
        import traceback
        traceback.print_exc()

    # Type text
    selection.TypeText("Method 7: Custom template - This is a long bullet point to test text wrapping and see if the wrapped lines align properly with the first line of text.")
    report_paragraph_format(selection, "after typing text")

    selection.TypeParagraph()
    selection.Range.ListFormat.RemoveNumbers()
    selection.TypeParagraph()


def main():
    print("Word Bullet Indentation Test")
    print("="*60)

    word = get_word_app()

    # Create a new document for testing
    doc = word.Documents.Add()
    print("Created new document for testing")

    selection = word.Selection

    # Add header
    selection.TypeText("BULLET INDENTATION TEST RESULTS")
    selection.TypeParagraph()
    selection.TypeText("Target: Bullet at 0.5 inch, text at 0.75 inch (0.25 inch gap)")
    selection.TypeParagraph()
    selection.TypeParagraph()

    # Run all test methods
    test_method_1_apply_bullet_then_indent(word, doc)
    test_method_2_indent_then_bullet(word, doc)
    test_method_3_indent_bullet_indent(word, doc)
    test_method_4_manual_bullet(word, doc)
    test_method_5_modify_list_template(word, doc)
    test_method_7_apply_list_template_with_level(word, doc)

    print("\n" + "="*60)
    print("TEST COMPLETE - Check the Word document to see visual results")
    print("="*60)
    print("\nLook at each bullet point in Word and check:")
    print("1. Does the wrapped text align with the first line?")
    print("2. Is the bullet at 0.5 inch?")
    print("3. Is the text at 0.75 inch?")
    print("\nThe console output above shows what values were actually applied.")


if __name__ == "__main__":
    main()
