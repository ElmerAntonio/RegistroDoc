import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))

import unittest
from unittest.mock import MagicMock, patch

class TestSetupCentering(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # Save original sys.modules
        cls._orig_modules = sys.modules.copy()
        
        # Mock all possible dependencies of setup.py
        cls.mock_ctk = MagicMock()
        sys.modules['customtkinter'] = cls.mock_ctk
        sys.modules['tkinter'] = MagicMock()
        sys.modules['tkinter.messagebox'] = MagicMock()
        sys.modules['pandas'] = MagicMock()
        sys.modules['matplotlib'] = MagicMock()
        sys.modules['rddata'] = MagicMock()
        sys.modules['config'] = MagicMock()
        sys.modules['config'].CONFIG_FILE = "test_config.json"
        sys.modules['config'].BASE_DIR = "."

        class MockCTk:
            def __init__(self, *args, **kwargs): pass
            def title(self, *args, **kwargs): pass
            def geometry(self, *args, **kwargs): pass
            def winfo_screenwidth(self): return 1920
            def winfo_screenheight(self): return 1080
            def crear_interfaz(self): pass

        cls.mock_ctk.CTk = MockCTk
        
        # Add src to path
        src_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '../src'))
        if src_path not in sys.path:
            sys.path.insert(0, src_path)

    @classmethod
    def tearDownClass(cls):
        # Restore sys.modules entirely
        added_keys = set(sys.modules.keys()) - set(cls._orig_modules.keys())
        for k in added_keys:
            del sys.modules[k]
        for k, v in cls._orig_modules.items():
            sys.modules[k] = v

    def test_setup_wizard_centering(self):
        from setup import SetupWizard
        # Patch crear_interfaz so it doesn't do anything
        with patch.object(SetupWizard, 'crear_interfaz'):
            # Re-mock geometry on the class to capture the call in __init__
            mock_ctk_class = sys.modules['customtkinter'].CTk
            with patch.object(mock_ctk_class, 'geometry') as mock_geo:
                wizard = SetupWizard()
                # Verify geometry was called to center the window
                # width=700, height=500, screen_w=1920, screen_h=1080
                # x = (1920-700)//2 = 610
                # y = (1080-500)//2 = 290
                mock_geo.assert_called_with("700x500+610+290")

    def test_logic_directly(self):
        # Simpler test: just the math
        width, height = 700, 500
        screen_w, screen_h = 1920, 1080
        x = (screen_w - width) // 2
        y = (screen_h - height) // 2

        self.assertEqual(x, 610)
        self.assertEqual(y, 290)
        self.assertEqual(f"{width}x{height}+{x}+{y}", "700x500+610+290")

if __name__ == "__main__":
    unittest.main()

