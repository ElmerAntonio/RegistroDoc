import unittest
import os
import tempfile
import hashlib
from unittest.mock import patch, MagicMock
import sys
import datetime

# Pre-import setup
os.environ["REGISTRODOC_MASTER_SALT"] = "TEST_SALT_INTEGRITY"

# Add src to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

# Mock dependencies before importing rdsecurity
sys.modules['dotenv'] = MagicMock()
sys.modules['cryptography'] = MagicMock()
sys.modules['cryptography.hazmat'] = MagicMock()
sys.modules['cryptography.hazmat.primitives'] = MagicMock()
sys.modules['cryptography.hazmat.primitives.ciphers'] = MagicMock()
sys.modules['cryptography.hazmat.primitives.ciphers.aead'] = MagicMock()
sys.modules['cryptography.hazmat.primitives.kdf'] = MagicMock()
sys.modules['cryptography.hazmat.primitives.kdf.pbkdf2'] = MagicMock()
sys.modules['cryptography.hazmat.backends'] = MagicMock()

import rdsecurity
from rdsecurity import (
    calcular_hash_excel,
    guardar_hash_excel,
    verificar_integridad_excel,
    actualizar_hash_excel
)

class TestRdSecurityIntegrity(unittest.TestCase):

    def setUp(self):
        # Create a temporary Excel-like file
        fd, self.temp_excel_path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        self.content = b"fake excel data content"
        with open(self.temp_excel_path, "wb") as f:
            f.write(self.content)

        # Create temporary hash and audit file paths
        fd, self.temp_hash_file = tempfile.mkstemp(suffix=".bin")
        os.close(fd)
        fd, self.temp_audit_file = tempfile.mkstemp(suffix=".bin")
        os.close(fd)

    def tearDown(self):
        if os.path.exists(self.temp_excel_path):
            os.remove(self.temp_excel_path)
        if os.path.exists(self.temp_hash_file):
            os.remove(self.temp_hash_file)
        if os.path.exists(self.temp_audit_file):
            os.remove(self.temp_audit_file)

    def test_calcular_hash_excel_valid(self):
        expected_hash = hashlib.sha3_256(self.content).hexdigest()
        self.assertEqual(calcular_hash_excel(self.temp_excel_path), expected_hash)

    def test_calcular_hash_excel_missing(self):
        self.assertEqual(calcular_hash_excel("non_existent_file.xlsx"), "")

    @patch("rdsecurity.guardar_cifrado")
    def test_guardar_hash_excel(self, mock_guardar):
        with patch("rdsecurity.EXCEL_HASH_FILE", self.temp_hash_file), \
             patch("rdsecurity.AUDIT_FILE", self.temp_audit_file), \
             patch("rdsecurity.registrar_auditoria") as mock_audit:

            guardar_hash_excel(self.temp_excel_path)

            # Verify it tried to save the hash
            mock_guardar.assert_called_once()
            args, _ = mock_guardar.call_args
            self.assertEqual(args[0], self.temp_hash_file)
            datos = args[1]
            self.assertEqual(datos["hash"], calcular_hash_excel(self.temp_excel_path))

            # Verify audit
            mock_audit.assert_called_once()

    @patch("rdsecurity.guardar_cifrado")
    @patch("rdsecurity.cargar_cifrado")
    def test_verificar_integridad_excel_workflow(self, mock_cargar, mock_guardar):
        with patch("rdsecurity.EXCEL_HASH_FILE", self.temp_hash_file), \
             patch("rdsecurity.AUDIT_FILE", self.temp_audit_file), \
             patch("rdsecurity.registrar_auditoria"):

            # 1. Initial call (no hash file exists)
            with patch("os.path.exists", side_effect=lambda p: p != self.temp_hash_file if p == self.temp_hash_file else os.path.exists(p)):
                integro, msg = verificar_integridad_excel(self.temp_excel_path)
                self.assertTrue(integro)
                self.assertIn("inicial", msg)
                mock_guardar.assert_called()

            mock_guardar.reset_mock()

            # 2. Subsequent call (no changes)
            current_hash = calcular_hash_excel(self.temp_excel_path)
            mock_cargar.return_value = {"hash": current_hash}
            with patch("os.path.exists", return_value=True):
                integro, msg = verificar_integridad_excel(self.temp_excel_path)
                self.assertTrue(integro)
                self.assertIn("verificada", msg)

            # 3. Modify file (conceptually by changing the mocked loaded hash)
            mock_cargar.return_value = {"hash": "old_hash"}
            with patch("os.path.exists", return_value=True):
                integro, msg = verificar_integridad_excel(self.temp_excel_path)
                self.assertFalse(integro)
                self.assertIn("modificado", msg)

    @patch("rdsecurity.guardar_cifrado")
    @patch("rdsecurity.cargar_cifrado")
    def test_verificar_integridad_hash_file_corrupt(self, mock_cargar, mock_guardar):
        with patch("rdsecurity.EXCEL_HASH_FILE", self.temp_hash_file), \
             patch("rdsecurity.AUDIT_FILE", self.temp_audit_file), \
             patch("rdsecurity.registrar_auditoria"):

            # Simulate corruption: cargar_cifrado returns {} when it fails to decrypt
            mock_cargar.return_value = {}

            with patch("os.path.exists", return_value=True):
                integro, msg = verificar_integridad_excel(self.temp_excel_path)
                self.assertTrue(integro)
                self.assertIn("regenerado", msg)
                mock_guardar.assert_called()

if __name__ == "__main__":
    unittest.main()
