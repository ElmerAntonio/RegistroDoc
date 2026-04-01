import re

with open("src/rddata.py", "r") as f:
    rddata_content = f.read()

# Replace bare excepts with except Exception:
rddata_content = re.sub(r'except:\s*pass', 'except Exception: pass', rddata_content)

with open("src/rddata.py", "w") as f:
    f.write(rddata_content)

with open("src/splash.py", "r") as f:
    splash_content = f.read()

# Remove unused imports
splash_content = re.sub(r'import time\n', '', splash_content)

with open("src/splash.py", "w") as f:
    f.write(splash_content)
