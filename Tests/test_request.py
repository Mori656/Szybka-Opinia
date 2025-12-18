import unittest
from unittest.mock import MagicMock, patch

from Classes.request import Request

class TestRequestIntegrationValidation(unittest.TestCase):

    def setUp(self):
        self.window = MagicMock()
        self.main_frame = MagicMock()
        self.request = Request(self.window, self.main_frame)
        self.request.document_info = {"Numer opinii": "1"}

    @patch("tkinter.messagebox.showerror")
    def test_createOpinion_stops_on_missing_device_id(self, mock_error):
        device1 = MagicMock()
        device1.field_id.get.return_value = ""

        device2 = MagicMock()
        device2.field_id.get.return_value = "123"

        self.request.devices = [device1, device2]

        result = self.request.createOpinion()

        mock_error.assert_called_once()
        self.assertFalse(result)

    @patch("tkinter.messagebox.showerror")
    def test_createOpinion_stops_on_validateRequiredFields_false(self, _):
        device = MagicMock()
        device.field_id.get.return_value = "123"
        device.validateRequiredFields.return_value = False

        self.request.devices = [device]

        result = self.request.createOpinion()

        self.assertFalse(result)
        device.updateDeviceInfo.assert_not_called()

class TestRequestIntegrationOpinion(unittest.TestCase):

    def setUp(self):
        self.window = MagicMock()
        self.main_frame = MagicMock()
        self.request = Request(self.window, self.main_frame)
        self.request.document_info = {"Numer opinii": "5"}

    @patch("Classes.request.Opinion")
    def test_updateDeviceInfo_called_before_opinion_creation(self, mock_opinion):
        device = MagicMock()
        device.field_id.get.return_value = "321"
        device.validateRequiredFields.return_value = True

        self.request.devices = [device]

        self.request.createOpinion()

        device.updateDeviceInfo.assert_called_once()
        mock_opinion.assert_called_once()

class TestRequestIntegrationPricing(unittest.TestCase):

    def setUp(self):
        self.window = MagicMock()
        self.main_frame = MagicMock()
        self.request = Request(self.window, self.main_frame)
        self.request.document_info = {"Numer opinii": "9"}

    @patch("Classes.request.Pricing")
    @patch("Classes.request.Opinion")
    def test_pricing_created_only_after_opinion(self, mock_opinion, mock_pricing):
        device = MagicMock()
        device.field_id.get.return_value = "123"
        device.validateRequiredFields.return_value = True
        device.getDecision.return_value = True
        device.pricing_file_name = "pricing"
        device.info = {"Numer wyceny": "77"}

        self.request.devices = [device]

        self.request.createOpinion()

        mock_opinion.assert_called_once()
        mock_pricing.assert_called_once()

class TestRequestIntegrationMultipleDevices(unittest.TestCase):

    @patch("Classes.request.Pricing")
    @patch("Classes.request.Opinion")
    def test_multiple_devices_partial_pricing(self, mock_opinion, mock_pricing):
        request = Request(MagicMock(), MagicMock())
        request.document_info = {"Numer opinii": "10"}

        device1 = MagicMock()
        device1.field_id.get.return_value = "1"
        device1.validateRequiredFields.return_value = True
        device1.getDecision.return_value = True
        device1.pricing_file_name = "p1"
        device1.info = {"Numer wyceny": "1"}

        device2 = MagicMock()
        device2.field_id.get.return_value = "2"
        device2.validateRequiredFields.return_value = True
        device2.getDecision.return_value = False

        request.devices = [device1, device2]

        request.createOpinion()

        self.assertEqual(mock_pricing.call_count, 1)

class TestRequestIntegrationPricingErrors(unittest.TestCase):

    @patch("tkinter.messagebox.showerror")
    @patch("Classes.request.Pricing", side_effect=[Exception("boom"), None])
    @patch("Classes.request.Opinion")
    def test_pricing_exception_does_not_break_flow(self, mock_opinion, _, mock_error):
        request = Request(MagicMock(), MagicMock())
        request.document_info = {"Numer opinii": "99"}

        d1 = MagicMock()
        d1.field_id.get.return_value = "1"
        d1.validateRequiredFields.return_value = True
        d1.getDecision.return_value = True
        d1.pricing_file_name = "p1"
        d1.info = {"Numer wyceny": "1"}

        d2 = MagicMock()
        d2.field_id.get.return_value = "2"
        d2.validateRequiredFields.return_value = True
        d2.getDecision.return_value = True
        d2.pricing_file_name = "p2"
        d2.info = {"Numer wyceny": "2"}

        request.devices = [d1, d2]

        request.createOpinion()

        self.assertEqual(mock_error.call_count, 1)
