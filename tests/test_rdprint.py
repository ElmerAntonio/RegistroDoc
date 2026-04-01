import pytest
import os
import sys
from unittest.mock import patch, MagicMock

# To fix the unresolvable import in testing environment, since rdprint is in src
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

# Note: We import functions directly to mock internal dependencies effectively.
# rdprint uses local imports of win32com to prevent loading errors on non-Windows.
from rdprint import imprimir_hoja_directo, abrir_para_imprimir


@pytest.fixture
def mock_platform_windows():
    with patch('platform.system', return_value="Windows"):
        yield


@pytest.fixture
def mock_ruta_excel(tmp_path):
    # Create a dummy file to bypass the os.path.isfile check
    dummy_file = tmp_path / "Registro_2026.xlsx"
    dummy_file.touch()
    with patch('rdprint._ruta_excel', return_value=str(dummy_file)):
        yield str(dummy_file)


class TestRDPrint:
    def test_imprimir_hoja_directo_success(self, mock_platform_windows, mock_ruta_excel):
        # This test verifies that win32com is called correctly with sanitized inputs
        # We mock win32com to prevent any actual interaction with Excel/COM
        mock_win32com = MagicMock()
        mock_excel_app = MagicMock()
        mock_workbook = MagicMock()
        mock_sheet = MagicMock()

        mock_win32com.client.DispatchEx.return_value = mock_excel_app
        mock_excel_app.Workbooks.Open.return_value = mock_workbook
        mock_workbook.Sheets.return_value = mock_sheet

        # We need to inject the mock into sys.modules so the local import in rdprint works
        with patch.dict('sys.modules', {'win32com': mock_win32com, 'win32com.client': mock_win32com.client}):
            exito, mensaje = imprimir_hoja_directo("Portada")

            # Verify success
            assert exito is True
            assert "enviada a la impresora" in mensaje

            # Verify COM interactions
            mock_win32com.client.DispatchEx.assert_called_once_with("Excel.Application")
            mock_excel_app.Workbooks.Open.assert_called_once_with(mock_ruta_excel, ReadOnly=True)
            mock_workbook.Sheets.assert_called_once_with("Portada")
            mock_sheet.PrintOut.assert_called_once()

            # Verify proper cleanup
            mock_workbook.Close.assert_called_once_with(SaveChanges=False)
            mock_excel_app.Quit.assert_called_once()


    def test_imprimir_hoja_directo_file_not_found(self):
        with patch('rdprint._ruta_excel', return_value="dummy_missing_file.xlsx"):
            exito, mensaje = imprimir_hoja_directo("Portada")
            assert exito is False
            assert "No se encontró el archivo" in mensaje


    def test_imprimir_hoja_directo_non_windows(self, mock_ruta_excel):
        with patch('platform.system', return_value="Linux"):
            with patch('rdprint.abrir_para_imprimir', return_value=(True, "Linux Fallback")):
                exito, mensaje = imprimir_hoja_directo("Portada")
                assert exito is True
                assert mensaje == "Linux Fallback"


    def test_imprimir_hoja_directo_com_error(self, mock_platform_windows, mock_ruta_excel):
        # Test fallback behavior when COM throws an exception
        mock_win32com = MagicMock()
        mock_win32com.client.DispatchEx.side_effect = Exception("COM Error")

        with patch.dict('sys.modules', {'win32com': mock_win32com, 'win32com.client': mock_win32com.client}):
            with patch('rdprint.abrir_para_imprimir', return_value=(True, "Fallback Abierto")):
                exito, mensaje = imprimir_hoja_directo("Portada")
                assert exito is True
                assert mensaje == "Fallback Abierto"
