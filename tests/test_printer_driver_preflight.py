import subprocess
import unittest
from importlib import import_module

preflight = import_module("deploy.printer_driver_preflight")


class TestPrinterDriverPreflight(unittest.TestCase):
    def test_build_get_printer_command_targets_sato(self):
        command = preflight.build_get_printer_command("SATO WS408")
        self.assertIn("powershell", command[0].lower())
        self.assertIn("SATO WS408", command[-1])

    def test_installer_printer_check_uses_runner_return_code(self):
        def runner(*_args, **_kwargs):
            return subprocess.CompletedProcess(_args, 0)

        self.assertTrue(preflight.is_printer_installed_for_installer(runner=runner))

    def test_guidance_mentions_it_and_detected_installers(self):
        message = preflight.format_available_driver_artifacts(["drivers/SATO_WS4.exe"])

        self.assertIn("departamento de informática", message)
        self.assertIn("48 mm x 23 mm", message)
        self.assertIn("203 DPI", message)
        self.assertIn("drivers/SATO_WS4.exe", message)


if __name__ == "__main__":
    unittest.main()
