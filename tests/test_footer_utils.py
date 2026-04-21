import unittest
from unittest.mock import patch, MagicMock, mock_open
import os
import sys
import json

# Mock missing dependencies to allow importing footer_utils
sys.modules['docx'] = MagicMock()
sys.modules['docx.shared'] = MagicMock()
sys.modules['docx.oxml'] = MagicMock()
sys.modules['docx.oxml.ns'] = MagicMock()
sys.modules['docx.enum.text'] = MagicMock()
sys.modules['PIL'] = MagicMock()
sys.modules['customtkinter'] = MagicMock()

# Add src to path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src')))

from utils.footer_utils import get_school_logo_path

class TestFooterUtils(unittest.TestCase):

    def setUp(self):
        # We need to mock CONFIG_FILE path because it's imported in the function
        pass

    @patch('os.path.exists')
    @patch('builtins.open', new_callable=mock_open, read_data='{"logo_escuela_path": "custom/path/logo.png"}')
    def test_get_school_logo_path_from_config(self, mock_file, mock_exists):
        # Mocking CONFIG_FILE exists and custom path exists
        mock_exists.side_effect = lambda p: True

        path = get_school_logo_path()
        self.assertEqual(path, "custom/path/logo.png")

    @patch('os.path.exists')
    @patch('builtins.open', new_callable=mock_open, read_data='{"logo_escuela_path": "non/existent/path.png"}')
    def test_get_school_logo_path_config_not_found_fallback_to_meipass(self, mock_file, mock_exists):
        # Mocking CONFIG_FILE exists, but custom path DOES NOT exist.
        # Fallback to _MEIPASS.
        # We also need to mock internal_path exists.

        def exists_side_effect(p):
            if "perfil.json" in p: return True
            if "non/existent/path.png" in p: return False
            if "img/logo.png" in p: return True
            return False

        mock_exists.side_effect = exists_side_effect

        # Mock sys._MEIPASS
        with patch('sys._MEIPASS', '/tmp/bundle', create=True):
            path = get_school_logo_path()
            self.assertEqual(path, os.path.join('/tmp/bundle', "img", "logo.png"))

    @patch('os.path.exists')
    def test_get_school_logo_path_no_config_fallback_to_local(self, mock_exists):
        # CONFIG_FILE does not exist.
        # sys._MEIPASS does not exist.
        # Fallback to os.path.abspath(".") / img / logo.png

        def exists_side_effect(p):
            if "perfil.json" in p: return False
            if "img/logo.png" in p: return True
            return False

        mock_exists.side_effect = exists_side_effect

        # Ensure _MEIPASS is NOT there
        if hasattr(sys, '_MEIPASS'):
            del sys._MEIPASS

        path = get_school_logo_path()
        expected = os.path.join(os.path.abspath("."), "img", "logo.png")
        self.assertEqual(path, expected)

    @patch('os.path.exists')
    def test_get_school_logo_path_nothing_found(self, mock_exists):
        mock_exists.return_value = False

        path = get_school_logo_path()
        self.assertEqual(path, "")

    @patch('os.path.exists')
    @patch('builtins.open', side_effect=Exception("Read error"))
    def test_get_school_logo_path_config_read_error(self, mock_file, mock_exists):
        # CONFIG_FILE exists but reading fails
        def exists_side_effect(p):
            if "perfil.json" in p: return True
            if "img/logo.png" in p: return True
            return False
        mock_exists.side_effect = exists_side_effect

        with patch('sys._MEIPASS', '/tmp/bundle', create=True):
            path = get_school_logo_path()
            self.assertEqual(path, os.path.join('/tmp/bundle', "img", "logo.png"))

if __name__ == '__main__':
    unittest.main()
