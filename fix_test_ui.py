import re

with open("tests/test_ui_limits.py", "r") as f:
    content = f.read()

content = content.replace("with patch('tkinter.messagebox.showwarning') as mock_showwarning:", "with patch('src.dapp.messagebox.showwarning') as mock_showwarning:")

with open("tests/test_ui_limits.py", "w") as f:
    f.write(content)
