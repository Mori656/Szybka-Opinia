import unittest
from unittest.mock import MagicMock, patch
from Classes.device import Device

# Testy jednostkowe dla klasy Device
class TestDeviceUnit(unittest.TestCase):

    @patch("tkinter.Frame")
    def setUp(self, _):
        self.device = Device(frame=MagicMock(), lp=1)

    def test_getDecision_liquidation(self):
        self.device.decision = MagicMock()
        self.device.decision.get.return_value = "Likwidacja"
        self.assertFalse(self.device.getDecision())

    def test_getDecision_valuation(self):
        self.device.decision = MagicMock()
        self.device.decision.get.return_value = "WycenaL"
        self.assertTrue(self.device.getDecision())

    @patch("tkinter.messagebox.askyesno", return_value=False)
    def test_validateRequiredFields_missing(self, mock_msg):
        widget = MagicMock()
        widget.get.return_value = ""
        widget.__getitem__.return_value = "normal"

        self.device.fields = {
            "Numer ewidencyjny": widget
        }

        result = self.device.validateRequiredFields()

        self.assertFalse(result)
        mock_msg.assert_called_once()

    def test_updateDeviceInfo_updates_info(self):
        entry = MagicMock()
        entry.get.return_value = "TEST"
        entry.cget.return_value = "normal"

        self.device.fields = {
            "Producent": entry
        }

        self.device.formula_area = MagicMock()
        self.device.formula_area.get.return_value = "FORMULA"

        self.device.state_area = MagicMock()
        self.device.state_area.cget.return_value = "disabled"

        self.device.change_area = MagicMock()
        self.device.change_area.cget.return_value = "disabled"

        self.device.change_cooling = MagicMock(get=MagicMock(return_value=False))
        self.device.change_batery = MagicMock(get=MagicMock(return_value=False))

        self.device.updateDeviceInfo()

        self.assertEqual(self.device.info["Producent"], "TEST")
        self.assertEqual(self.device.info["Formula"], "FORMULA")

# Testy integracyjne dla klasy Device
class TestDeviceIntegration(unittest.TestCase):

    @patch("tkinter.Frame")
    @patch("openpyxl.load_workbook")
    def test_findDataByKey_fills_info(self, mock_wb, _):

        # Fake excel
        sheet = MagicMock()
        sheet.max_row = 2
        sheet.max_column = 2
        sheet.__getitem__.return_value = [
            MagicMock(value="Numer ewidencyjny"),
            MagicMock(value="Producent")
        ]

        def cell(row, column):
            values = {
                (1, 1): "Numer ewidencyjny",
                (1, 2): "Producent",
                (2, 1): "123",
                (2, 2): "Dell"
            }
            m = MagicMock()
            m.value = values.get((row, column))
            return m

        sheet.cell.side_effect = cell
        mock_wb.return_value.active = sheet 

        device = Device(frame=MagicMock(), lp=1)
        device.paths = ["Log1.xlsx"]
        device.field_id = MagicMock()
        device.field_id.get.return_value = "123"

        device.info = {
            "Numer ewidencyjny": "123",
            "Producent": ""
        }

        found = device.findDataByKey("Numer ewidencyjny")

        self.assertIn("Producent", found)
        self.assertEqual(device.info["Producent"], "Dell")

    @patch("tkinter.messagebox.showwarning")
    @patch("tkinter.Frame")
    def test_findData_shows_warning_on_missing_fields(self, _, mock_warning):
        device = Device(frame=MagicMock(), lp=1)
        device.combo = MagicMock()
        device.combo.get.return_value = "ID"
        device.field_id = MagicMock()
        device.field_id.get.return_value = "999"

        device.findDataByKey = MagicMock(return_value=set())

        device.findData()

        mock_warning.assert_called_once()

