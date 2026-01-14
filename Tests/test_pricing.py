import unittest
import os
import shutil

from Classes.pricing import Pricing

class DummyDevice:
    def __init__(self):
        self.info = {
            "Numer ewidencyjny": "123",
            "Numer wyceny": "WYC-1",
            "MPK": "IT",
            "Numer środka trwałego": "ST-1",
            "Nr inwentarzowy": "INV-1",
            "Producent": "Dell",
            "Model": "Latitude",
            "Typ": "Laptop",
            "Numer seryjny": "SN123",
            "Data zakupu": "01.01.2020",
            "Model procesora": "i5",
            "Grafika": "Intel",
            "Pamięć RAM (GB)": "16",
            "Pamięć HDD (GB)": "512",
            "System operacyjny": "Windows",
            "Stan": "Sprawny",
            "Wymiana": "",

            # ceny / linki
            **{f"Price{i}": "1000" for i in range(1, 11)},
            **{f"Link{i}": "http://example.com" for i in range(1, 11)},

            "Cena_cooling": "",
            "Link_cooling": "",
            "Cena_battery": "",
            "Link_battery": "",
        }


class TestPricingIntegration(unittest.TestCase):

    def setUp(self):
        self.info = {
            "Nazwa pliku": "TEST_OPINIA",
            "Numer opinii": "001"
        }

        os.makedirs(
            f"./Wygenerowane_opinie/{self.info['Nazwa pliku']}",
            exist_ok=True
        )
    def tearDown(self):
        if os.path.exists("./Wygenerowane_opinie/TEST_OPINIA"):
            shutil.rmtree("./Wygenerowane_opinie/TEST_OPINIA")

    def test_pricing_creates_correct_directory(self):
        device = DummyDevice()

        Pricing(self.info, device, "wycena_test")

        expected_dir = (
            "./Wygenerowane_opinie/"
            f"{self.info['Nazwa pliku']}/"
            f"Wycena_{device.info['Numer ewidencyjny']}"
        )

        self.assertTrue(
            os.path.isdir(expected_dir),
            "Folder wyceny nie został utworzony"
        )

    def test_pricing_copies_cennik_pdf(self):
        device = DummyDevice()

        Pricing(self.info, device, "wycena_test")

        expected_dir = (
            "./Wygenerowane_opinie/"
            f"{self.info['Nazwa pliku']}/"
            f"Wycena_{device.info['Numer ewidencyjny']}"
        )

        cennik_path = os.path.join(expected_dir, "Cennik.pdf")

        self.assertTrue(
            os.path.exists(cennik_path),
            "Cennik.pdf nie został skopiowany do folderu wyceny"
        )

    def test_pricing_creates_xlsx_file(self):
        device = DummyDevice()

        Pricing(self.info, device, "wycena_test")

        expected_xlsx = (
            "./Wygenerowane_opinie/"
            f"{self.info['Nazwa pliku']}/"
            f"Wycena_{device.info['Numer ewidencyjny']}/"
            "wycena_test.xlsx"
        )

        self.assertTrue(
            os.path.exists(expected_xlsx),
            "Plik XLSX z wyceną nie został utworzony"
        )
