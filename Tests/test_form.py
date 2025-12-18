import unittest
from Classes.form import Form
from unittest.mock import MagicMock, patch

class DummyDevice:
    def __init__(self):
        self.field_id = MagicMock()
        self.field_id.get.return_value = "123"

    def findData(self):
        pass


## Testy jednostkowe dla klasy Form
class formTestCase(unittest.TestCase):
    def setUp(self):
        self.form = Form()

    def test_format_reporter_name_standard(self):
        self.assertEqual(
            self.form.format_reporter_name("Jan Kowalski"),
            "J.Kowalski"
        )

    def test_format_reporter_name_many_parts(self):
        self.assertEqual(
            self.form.format_reporter_name("Jan Adam Kowalski"),
            "J.Kowalski"
        )

    def test_format_reporter_name_single_word(self):
        self.assertEqual(
            self.form.format_reporter_name("Admin"),
            "Admin"
        )

    def test_format_author_name(self):
        self.assertEqual(
            self.form.format_author_name("Jan Kowalski"),
            "JK"
        )

    def test_format_author_name_single(self):
        self.assertEqual(
            self.form.format_author_name("Admin"),
            "Admin"
        )

## Testy integracyjne dla klasy Form

class formIntegrationTestCase(unittest.TestCase):
    @patch("openpyxl.load_workbook")
    def test_get_departments_mocked(self, mock_wb):
        mock_sheet = MagicMock()
        mock_sheet.max_row = 2
        mock_sheet.__getitem__.return_value = [
            MagicMock(value="Kod działu")
        ]
        mock_wb.return_value.active = mock_sheet

        form = Form()
        result = form.get_departments()
        self.assertIsInstance(result, list)


if __name__ == "__main__":
    unittest.main()

