import unittest

from services.driver_check import (
    PrinterDriverMissingError,
    check_printer_driver,
    ensure_printer_driver,
    missing_driver_message,
)


class TestDriverCheck(unittest.TestCase):
    def test_detects_installed_sato_printer(self):
        status = check_printer_driver(
            "SATO WS408",
            printer_names_provider=lambda: ["Microsoft Print to PDF", "SATO WS408"],
        )

        self.assertTrue(status.installed)
        self.assertEqual(status.message, "")

    def test_missing_driver_message_guides_user_to_it(self):
        message = missing_driver_message("SATO WS408")

        self.assertIn("Controlador SATO WS408 no detectado", message)
        self.assertIn("contactá a TI", message)
        self.assertIn("departamento de informática", message)

    def test_ensure_printer_driver_raises_when_missing(self):
        with self.assertRaises(PrinterDriverMissingError):
            ensure_printer_driver(
                "SATO WS408",
                printer_names_provider=lambda: ["Microsoft Print to PDF"],
            )


if __name__ == "__main__":
    unittest.main()
