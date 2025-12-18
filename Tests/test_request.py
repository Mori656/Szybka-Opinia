import unittest
from unittest.mock import MagicMock, patch
import tkinter as tk
import tkinter.messagebox as messagebox

from Classes.request import Request

# Mockowanie messageboxów, aby nie wywoływały okien podczas testów
messagebox.showerror = lambda *a, **k: None
messagebox.showinfo = lambda *a, **k: None
messagebox.showwarning = lambda *a, **k: None
messagebox.askyesno = lambda *a, **k: True


# Urządzenie-mock do testów
def mock_device(valid=True, decision=True):
    device = MagicMock()
    device.field_id.get.return_value = "123" if valid else ""
    device.validateRequiredFields.return_value = valid
    device.getDecision.return_value = decision
    device.pricing_file_name = "pricing.xlsx"
    device.info = {"Numer wyceny": "999"}
    device.updateDeviceInfo = MagicMock()
    device.prepare_frame = MagicMock()
    return device


## Testy jednostkowe dla klasy Request

class TestRequestUnit(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.root = tk.Tk()
        cls.root.withdraw()

    @classmethod
    def tearDownClass(cls):
        cls.root.destroy()

    def setUp(self):
        self.main_frame = tk.Frame(self.root)
        self.request = Request(self.root, self.main_frame)
        self.request.document_info = {"Numer opinii": "321"}

    @patch("tkinter.messagebox.showerror")
    def test_create_opinion_fails_on_missing_device_id(self, mock_error):
        self.request.devices = [mock_device(valid=False)]

        result = self.request.createOpinion()

        self.assertFalse(result)
        mock_error.assert_called_once()

    def test_create_opinion_fails_on_invalid_required_fields(self):
        device = mock_device()
        device.validateRequiredFields.return_value = False
        self.request.devices = [device]

        result = self.request.createOpinion()

        self.assertFalse(result)
        device.updateDeviceInfo.assert_not_called()

    @patch("Classes.request.Opinion")
    @patch("tkinter.messagebox.showinfo")
    def test_create_opinion_success(self, mock_info, mock_opinion):
        self.request.devices = [mock_device()]

        self.request.createOpinion()

        mock_opinion.assert_called_once()
        mock_info.assert_called_once()

    @patch("Classes.request.Opinion", side_effect=Exception("error"))
    @patch("tkinter.messagebox.showerror")
    def test_create_opinion_exception_sets_opinion_none(self, mock_error, _):
        self.request.devices = [mock_device()]

        self.request.createOpinion()

        self.assertIsNone(self.request.opinion)
        self.assertGreaterEqual(mock_error.call_count, 1)

    @patch("Classes.request.Pricing")
    def test_create_pricing_called_when_decision_true(self, mock_pricing):
        self.request.devices = [mock_device(decision=True)]

        self.request.createPricing()

        mock_pricing.assert_called_once()

    @patch("Classes.request.Pricing")
    def test_create_pricing_not_called_when_decision_false(self, mock_pricing):
        self.request.devices = [mock_device(decision=False)]

        self.request.createPricing()

        mock_pricing.assert_not_called()

    @patch("Classes.request.Pricing", side_effect=Exception("error"))
    @patch("tkinter.messagebox.showerror")
    def test_create_pricing_exception_is_handled(self, mock_error, _):
        self.request.devices = [mock_device(decision=True)]

        self.request.createPricing()

        mock_error.assert_called_once()


## Testy integracyjne dla klasy Request

class TestRequestIntegrationValidation(unittest.TestCase):

    def setUp(self):
        self.request = Request(MagicMock(), MagicMock())
        self.request.document_info = {"Numer opinii": "1"}

    @patch("tkinter.messagebox.showerror")
    def test_createOpinion_stops_on_first_invalid_device(self, mock_error):
        self.request.devices = [
            mock_device(valid=False),
            mock_device(valid=True)
        ]

        result = self.request.createOpinion()

        self.assertFalse(result)
        mock_error.assert_called_once()


class TestRequestIntegrationOpinion(unittest.TestCase):

    def setUp(self):
        self.request = Request(MagicMock(), MagicMock())
        self.request.document_info = {"Numer opinii": "5"}

    @patch("Classes.request.Opinion")
    def test_updateDeviceInfo_called_for_all_devices(self, mock_opinion):
        d1 = mock_device()
        d2 = mock_device()
        self.request.devices = [d1, d2]

        self.request.createOpinion()

        d1.updateDeviceInfo.assert_called_once()
        d2.updateDeviceInfo.assert_called_once()
        mock_opinion.assert_called_once()


class TestRequestIntegrationPricing(unittest.TestCase):

    @patch("Classes.request.Pricing")
    @patch("Classes.request.Opinion")
    def test_pricing_created_only_for_devices_with_decision(self, _, mock_pricing):
        request = Request(MagicMock(), MagicMock())
        request.document_info = {"Numer opinii": "10"}

        d1 = mock_device(decision=True)
        d2 = mock_device(decision=False)

        request.devices = [d1, d2]

        request.createOpinion()

        self.assertEqual(mock_pricing.call_count, 1)


class TestRequestIntegrationPricingErrors(unittest.TestCase):

    @patch("tkinter.messagebox.showerror")
    @patch("Classes.request.Pricing", side_effect=[Exception("boom"), None])
    @patch("Classes.request.Opinion")
    def test_pricing_exception_does_not_stop_other_pricing(self, _, __, mock_error):
        request = Request(MagicMock(), MagicMock())
        request.document_info = {"Numer opinii": "99"}

        d1 = mock_device(decision=True)
        d2 = mock_device(decision=True)

        request.devices = [d1, d2]

        request.createOpinion()

        self.assertEqual(mock_error.call_count, 1)
