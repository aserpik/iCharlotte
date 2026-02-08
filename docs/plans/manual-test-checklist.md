# Word Redline Manual Testing Checklist

## Basic Functionality
- [ ] Open Word document
- [ ] Select text (e.g., a paragraph)
- [ ] Press Win+V to show popup
- [ ] Verify redline checkbox appears
- [ ] Check redline checkbox
- [ ] Enter prompt (e.g., "Make this more professional")
- [ ] Click Process
- [ ] Verify Track Changes appear in Word
- [ ] Accept/reject changes in Word UI

## Edge Cases
- [ ] No text selected -> checkbox disabled with tooltip
- [ ] Very small selection (1 word) -> works
- [ ] Large selection (1000+ words) -> works or shows warning
- [ ] Selection with formatting (bold, italic) -> formatting preserved
- [ ] Selection in table -> works or gracefully falls back

## Persistence
- [ ] Check redline checkbox -> close dialog -> reopen -> still checked
- [ ] Uncheck redline checkbox -> close dialog -> reopen -> still unchecked

## Error Handling
- [ ] Document with Track Changes OFF -> auto-enables
- [ ] adeu fails (simulate by renaming package) -> falls back to replace mode
- [ ] Shows appropriate error message on fallback

## Outlook Context
- [ ] Open Outlook compose window
- [ ] Press Win+V
- [ ] Verify redline checkbox is HIDDEN (not applicable to email)
