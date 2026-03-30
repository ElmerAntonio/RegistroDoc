import os
import sys
import pytest
from unittest.mock import patch, MagicMock

# Mock cryptography for rdsecurity if needed during test imports
sys.modules['cryptography'] = MagicMock()
sys.modules['cryptography.hazmat'] = MagicMock()
sys.modules['cryptography.hazmat.primitives'] = MagicMock()

# Mock tkinter and winreg before importing rdprint
sys.modules['winreg'] = MagicMock()
sys.modules['tkinter'] = MagicMock()
sys.modules['tkinter.ttk'] = MagicMock()
sys.modules['tkinter.messagebox'] = MagicMock()

from src.rdprint import imprimir_hoja_directo, abrir_para_imprimir

@pytest.fixture
def mock_platform_windows():
    with patch('src.rdprint.platform.system', return_value="Windows"):
        yield

@pytest.fixture
def mock_ruta_excel(tmp_path):
    # Create a dummy excel file in a temp directory
    fake_excel = tmp_path / "Registro_2026.xlsx"
    fake_excel.touch()
    with patch('src.rdprint._ruta_excel', return_value=str(fake_excel)):
        yield str(fake_excel)

def test_imprimir_hoja_directo_success(mock_platform_windows, mock_ruta_excel):
    # This test verifies that win32com is called correctly with sanitized inputs
    # We mock win32com to prevent any actual interaction with Excel/COM
    mock_win32com = MagicMock()
    mock_excel_app = MagicMock()
    mock_workbook = MagicMock()
    mock_sheet = MagicMock()

    mock_win32com.client.Dispatch.return_value = mock_excel_app
    mock_excel_app.Workbooks.Open.return_value = mock_workbook
    mock_workbook.Sheets.return_value = mock_sheet

    # We need to inject the mock into sys.modules so the local import in rdprint works
    with patch.dict('sys.modules', {'win32com': mock_win32com, 'win32com.client': mock_win32com.client}):
        exito, mensaje = imprimir_hoja_directo("Portada")

        # Verify success
        assert exito is True
        assert "enviada a la impresora" in mensaje

        # Verify COM interactions
        mock_win32com.client.Dispatch.assert_called_once_with("Excel.Application")
        assert mock_excel_app.Visible is False
        assert mock_excel_app.DisplayAlerts is False

        # Verify open was called with absolute path
        abs_expected_path = os.path.abspath(mock_ruta_excel)
        mock_excel_app.Workbooks.Open.assert_called_once_with(abs_expected_path)

        # Verify specific sheet was printed
        mock_workbook.Sheets.assert_called_once_with("Portada")
        mock_sheet.PrintOut.assert_called_once()

        # Verify proper cleanup
        mock_workbook.Close.assert_called_once_with(SaveChanges=False)
        mock_excel_app.Quit.assert_called_once()

def test_imprimir_hoja_directo_file_not_found():
    with patch('src.rdprint._ruta_excel', return_value="invalid_path.xlsx"):
        exito, mensaje = imprimir_hoja_directo("Portada")
        assert exito is False
        assert "No se encontró el archivo Excel" in mensaje

def test_imprimir_hoja_directo_non_windows(mock_ruta_excel):
    with patch('src.rdprint.platform.system', return_value="Linux"):
        with patch('src.rdprint.abrir_para_imprimir', return_value=(True, "Abierto en Linux")):
            exito, mensaje = imprimir_hoja_directo("Portada")
            assert exito is True
            assert mensaje == "Abierto en Linux"

def test_imprimir_hoja_directo_com_error(mock_platform_windows, mock_ruta_excel):
    # Test fallback behavior when COM throws an exception
    mock_win32com = MagicMock()
    mock_win32com.client.Dispatch.side_effect = Exception("COM Error")

    with patch.dict('sys.modules', {'win32com': mock_win32com, 'win32com.client': mock_win32com.client}):
        with patch('src.rdprint.abrir_para_imprimir', return_value=(True, "Fallback Abierto")):
            exito, mensaje = imprimir_hoja_directo("Portada")
            assert exito is True
            assert mensaje == "Fallback Abierto"
