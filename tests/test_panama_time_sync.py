import unittest
from unittest.mock import patch, MagicMock
import sys
import os
import datetime

# Add src to path
src_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '../src'))
if src_path not in sys.path:
    sys.path.insert(0, src_path)

from utils.date_helpers import (
    obtener_ahora_panama,
    obtener_hoy_panama,
    es_hora_sincronizada,
    iniciar_sincronizacion_hora_panama
)

class TestPanamaTimeSync(unittest.TestCase):

    @patch('urllib.request.urlopen')
    def test_time_sync_worldtimeapi_success(self, mock_urlopen):
        # Mock successful WorldTimeAPI response
        mock_response = MagicMock()
        mock_response.read.return_value = b'{"datetime": "2026-06-11T12:00:00.000000-05:00"}'
        
        mock_context = MagicMock()
        mock_context.__enter__.return_value = mock_response
        mock_urlopen.return_value = mock_context

        # Call synchronization
        import utils.date_helpers
        utils.date_helpers._TIME_SYNCHRONIZED = False
        utils.date_helpers._PANAMA_TIME_OFFSET_SECONDS = 0.0

        # We call the thread function synchronously for testing
        def dummy_thread(target, daemon=True):
            target()
            return MagicMock()
            
        with patch('threading.Thread', side_effect=dummy_thread):
            iniciar_sincronizacion_hora_panama()

        self.assertTrue(es_hora_sincronizada())
        now_panama = obtener_ahora_panama()
        self.assertEqual(now_panama.year, 2026)
        self.assertEqual(now_panama.month, 6)
        self.assertEqual(now_panama.day, 11)

    @patch('urllib.request.urlopen')
    def test_time_sync_google_fallback(self, mock_urlopen):
        # First call to WorldTimeAPI fails, second call to Google succeeds
        mock_response_google = MagicMock()
        mock_response_google.headers = {"Date": "Thu, 11 Jun 2026 17:00:00 GMT"}
        
        def urlopen_side_effect(*args, **kwargs):
            if "worldtimeapi.org" in args[0].full_url:
                raise Exception("Network Timeout")
            
            mock_context = MagicMock()
            mock_context.__enter__.return_value = mock_response_google
            return mock_context

        mock_urlopen.side_effect = urlopen_side_effect

        import utils.date_helpers
        utils.date_helpers._TIME_SYNCHRONIZED = False
        utils.date_helpers._PANAMA_TIME_OFFSET_SECONDS = 0.0

        def dummy_thread(target, daemon=True):
            target()
            return MagicMock()

        with patch('threading.Thread', side_effect=dummy_thread):
            iniciar_sincronizacion_hora_panama()

        self.assertTrue(es_hora_sincronizada())
        now_panama = obtener_ahora_panama()
        self.assertEqual(now_panama.year, 2026)
        self.assertEqual(now_panama.month, 6)
        self.assertEqual(now_panama.day, 11)

if __name__ == '__main__':
    unittest.main()
