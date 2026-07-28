import importlib.util
import sys
import unittest
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = ROOT / "mercadopago_processor.py"

sys.argv = ["mercadopago_processor.py", "dummy.csv"]

spec = importlib.util.spec_from_file_location("mercadopago_processor", MODULE_PATH)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


class MercadoPagoDuplicateTests(unittest.TestCase):
    def test_same_date_and_description_is_duplicate_even_with_different_amount(self):
        existing = {("2024-01-01", "juan perez")}

        self.assertTrue(module.is_duplicate_existing(existing, "2024-01-01", "Juan Perez", 1500.00))
        self.assertFalse(module.is_duplicate_existing(existing, "2024-01-02", "Juan Perez", 1500.00))
        self.assertFalse(module.is_duplicate_existing(existing, "2024-01-01", "Maria Perez", 1500.00))

    def test_normalize_sort_value_handles_date_and_datetime(self):
        self.assertEqual(module.normalize_sort_value(date(2026, 7, 1)), datetime(2026, 7, 1, 0, 0))
        self.assertEqual(module.normalize_sort_value(datetime(2026, 7, 1, 12, 30)), datetime(2026, 7, 1, 12, 30))


if __name__ == "__main__":
    unittest.main()
