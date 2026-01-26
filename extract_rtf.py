import sys
import os
import glob
import re

def strip_rtf(text):
    pattern = re.compile(r"\\([a-z]{1,32})(-?\d{1,10})?[ ]?|\\\'([0-9a-f]{2})|\\([^a-z])|([{}])|[\r\n]+|(.)", re.I)
    destinations = frozenset((
        'ftnid', 'ftncn', 'ftnsep', 'ftnsepc', 'info', 'stylesheet', 'fonttbl',
        'colortbl', 'master', 'brdr', 'par', 'sect', 'pard', 'plain', 'filetbl',
        'revtbl', 'shppict', 'nonshppict', 'picprop', 'obj', 'pict'
    ))
    stack = []
    ignorable = False       # Whether this group (and all inside it) are "ignorable".
    ucskip = 1              # Number of ASCII characters to skip after a unicode character.
    curskip = 0             # Number of ASCII characters left to skip
    out = []                # Output buffer.
    
    for match in pattern.finditer(text):
        word, arg, hex, char, brace, tchar = match.groups()
        if brace:
            if brace == '{':
                # Push state
                stack.append((ucskip, ignorable))
            elif brace == '}':
                # Pop state
                if stack:
                    ucskip, ignorable = stack.pop()
        elif char: # \x (not a letter)
            if curskip > 0:
                curskip -= 1
            elif not ignorable:
                if char == '~': out.append('\xA0')
                elif char == '{': out.append('{')
                elif char == '}': out.append('}')
                elif char == '*': ignorable = True
        elif word: # \foo
            if curskip > 0:
                curskip -= 1
            elif not ignorable:
                if word == 'u':
                    curskip = ucskip
                    out.append(chr(int(arg)))
                elif word == 'uc':
                    ucskip = int(arg)
                elif word in destinations:
                    ignorable = True
                elif word == 'par':
                    out.append('\n')
                elif word == 'tab':
                    out.append('\t')
                elif word in ['line', 'row']:
                    out.append('\n')
        elif hex: # \'xx
            if curskip > 0:
                curskip -= 1
            elif not ignorable:
                out.append(chr(int(hex, 16)))
        elif tchar:
            if curskip > 0:
                curskip -= 1
            elif not ignorable:
                out.append(tchar)
    return "".join(out)

def main():
    target_dir = r"C:\\geminiterminal2\\LLM Resources\\Calendaring\\LLM Scan"
    files = glob.glob(os.path.join(target_dir, "*.doc"))
    
    all_text = ""
    for f_path in files:
        try:
            with open(f_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
                # Check if it looks like RTF
                if content.strip().startswith("{\\rtf"):
                    text = strip_rtf(content)
                    all_text += "\n\n--- FILE: " + os.path.basename(f_path) + " ---\\n\n"
                    all_text += text
                else:
                    all_text += "\n\n--- FILE: " + os.path.basename(f_path) + " (Not RTF) ---\\n\n"
        except Exception as e:
            all_text += "\n\n--- FILE: " + os.path.basename(f_path) + " (Error: " + str(e) + ") ---\\n\n"
            
    print(all_text)

if __name__ == "__main__":
    main()
