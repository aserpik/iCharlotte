"""Run read_docx_temp.py via subprocess, workaround for bash issues."""
import subprocess, sys
result = subprocess.run([sys.executable, 'C:/geminiterminal2/read_docx_temp.py'],
                       capture_output=True, text=True, cwd='C:/geminiterminal2')
if result.returncode == 0:
    # Read and print output
    with open('C:/geminiterminal2/docx_output.txt', 'r', encoding='utf-8') as f:
        print(f.read())
else:
    print(f"ERROR: {result.stderr}")
