from pathlib import Path
import unittest


PROJECT_DIR = Path(__file__).resolve().parents[1]
LAUNCHER_PATH = PROJECT_DIR / "启动采购自动化.bat"


class OneClickLauncherTests(unittest.TestCase):
    def test_launcher_bootstraps_a_local_environment_and_starts_the_workbench(self):
        content = LAUNCHER_PATH.read_text(encoding="utf-8")

        self.assertIn('cd /d "%~dp0"', content)
        self.assertIn('-m venv .venv', content)
        self.assertIn('-m pip install -r requirements.txt', content)
        self.assertIn('-m playwright install msedge', content)
        self.assertIn('"装潢透视表工具.py"', content)
        self.assertIn('fc /b "requirements.txt" ".venv\\.requirements-installed.txt"', content)
        self.assertNotIn('D:\\anaconda', content)


if __name__ == "__main__":
    unittest.main()
