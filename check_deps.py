try:
    import openpyxl
    print("openpyxl OK")
except ImportError as e:
    print(f"openpyxl FAIL: {e}")

try:
    import pandas
    print("pandas OK")
except ImportError as e:
    print(f"pandas FAIL: {e}")

try:
    import customtkinter
    print("customtkinter OK")
except ImportError as e:
    print(f"customtkinter FAIL: {e}")

try:
    import matplotlib
    print("matplotlib OK")
except ImportError as e:
    print(f"matplotlib FAIL: {e}")

try:
    import cryptography
    print("cryptography OK")
except ImportError as e:
    print(f"cryptography FAIL: {e}")

try:
    from PIL import Image
    print("Pillow OK")
except ImportError as e:
    print(f"Pillow FAIL: {e}")

try:
    import dotenv
    print("python-dotenv OK")
except ImportError as e:
    print(f"python-dotenv FAIL: {e}")
