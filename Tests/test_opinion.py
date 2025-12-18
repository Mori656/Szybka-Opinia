import unittest
import tempfile
import os
from Classes.opinion import Opinion

## Testy integracyjne dla klasy Opinion

# DummyDevice - klasa symulująca urządzenie do testów
class DummyDevice:
    def __init__(self):
        self.info = {
            "Typ": "Laptop",
            "Model": "Dell",
            "Nr inwentarzowy": "123",
            "Numer ewidencyjny": "456",
            "Numer środka trwałego": "789",
            "Data zakupu": "01.01.2020",
            "Model procesora": "i5",
            "Grafika": "Intel",
            "Pamięć RAM (GB)": "16",
            "Pamięć HDD (GB)": "512",
            "System operacyjny": "Windows",
            "Formula": "Sprzęt sprawny"
        }

class TestOpinion(unittest.TestCase):

    def test_creates_docx_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            info = {
                "Nazwa pliku": "TEST",
                "Numer opinii": "001",
                "Dla kogo": "IT",
                "Zglaszajacy": "Jan Kowalski",
                "Autor": "Adam Nowak"
            }

            opinion = Opinion(info, [DummyDevice()], tmp)

            expected_path = os.path.join(tmp, "TEST", "Opinia 001.docx")

            self.assertTrue(os.path.exists(expected_path))