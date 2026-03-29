import os
# Create the temp directory that bash is trying to use
path = r'C:\Users\ASERPI~1.DES\AppData\Local\Temp\claude\C--geminiterminal2\tasks'
os.makedirs(path, exist_ok=True)
print(f"Created: {path}")
print(f"Exists: {os.path.exists(path)}")
