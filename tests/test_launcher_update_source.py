import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import deploy.launcher_generadoretiquetassap as launcher


class TestLauncherUpdateSource(unittest.TestCase):
    def test_reads_update_root_from_installed_feed(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            install_root = Path(tmpdir) / "install"
            latest_root = Path(tmpdir) / "latest"
            install_root.mkdir()
            latest_root.mkdir()

            feed_file = install_root / launcher.UPDATE_FEED_FILE_NAME
            feed_file.write_text(
                json.dumps({"update_root": str(latest_root)}), encoding="utf-8"
            )

            with patch.object(launcher, "INSTALL_ROOT", install_root), patch.object(
                launcher, "LATEST_RELEASE_DIR", Path("fallback-unused")
            ):
                self.assertEqual(launcher._read_update_root(), latest_root)

    def test_falls_back_to_bundled_latest_when_feed_missing(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            install_root = Path(tmpdir) / "install"
            fallback_root = Path(tmpdir) / "bundled-latest"
            install_root.mkdir()

            with patch.object(launcher, "INSTALL_ROOT", install_root), patch.object(
                launcher, "LATEST_RELEASE_DIR", fallback_root
            ):
                self.assertEqual(launcher._read_update_root(), fallback_root)


if __name__ == "__main__":
    unittest.main()
