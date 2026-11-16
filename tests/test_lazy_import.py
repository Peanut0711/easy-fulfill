"""pandas 지연 로딩: 첫 화면 경로에서 불러오지 않고 처음 쓸 때·미리 읽기로 로드되는지 검증한다."""

import subprocess
import sys
import tempfile
import textwrap
import unittest
from pathlib import Path

from shared.lazy_import import deferred_module, preload_modules_in_background


ROOT = Path(__file__).resolve().parents[1]


class DeferredModuleTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        Path(self.directory.name, "lazy_probe_module.py").write_text(
            "LOADS = __import__('builtins').__dict__.setdefault('_lazy_probe_loads', [])\n"
            "LOADS.append(1)\nVALUE = 42\n",
            encoding="utf-8",
        )
        sys.path.insert(0, self.directory.name)

    def tearDown(self):
        sys.path.remove(self.directory.name)
        sys.modules.pop("lazy_probe_module", None)
        import builtins
        builtins.__dict__.pop("_lazy_probe_loads", None)
        self.directory.cleanup()

    def test_module_is_imported_on_first_attribute_access_only_once(self):
        proxy = deferred_module("lazy_probe_module")
        self.assertNotIn("lazy_probe_module", sys.modules)
        self.assertIn("deferred", repr(proxy))
        self.assertEqual(proxy.VALUE, 42)
        self.assertEqual(proxy.VALUE, 42)
        self.assertIs(proxy._load(), sys.modules["lazy_probe_module"])
        import builtins
        self.assertEqual(builtins._lazy_probe_loads, [1])

    def test_background_preload_imports_module(self):
        preload_modules_in_background("lazy_probe_module").join(timeout=10)
        self.assertIn("lazy_probe_module", sys.modules)

    def test_failed_preload_does_not_raise(self):
        preload_modules_in_background("module_that_does_not_exist_for_preload").join(timeout=10)


class StartupImportTests(unittest.TestCase):
    def test_startup_modules_do_not_import_pandas_or_numpy(self):
        code = textwrap.dedent("""
            import sys
            import orders.bulk, orders.invoice, orders.matching, orders.quick
            import sheets.db_sheet_sync, tracking.controller, tracking.repository, tracking.workers
            print(sorted(name for name in ("pandas", "numpy") if name in sys.modules))
        """)
        result = subprocess.run(
            [sys.executable, "-c", code], cwd=ROOT, capture_output=True, text=True, check=True,
        )
        self.assertEqual(result.stdout.strip(), "[]")

    def test_easy_fulfill_entry_module_defers_pandas(self):
        code = textwrap.dedent("""
            import importlib.util, os, sys
            os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
            spec = importlib.util.spec_from_file_location("ef_startup_probe", "easy-fulfill.py")
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            print("pandas" in sys.modules, "numpy" in sys.modules)
            print(module.pd.DataFrame({"a": [1]}).shape)
        """)
        result = subprocess.run(
            [sys.executable, "-c", code], cwd=ROOT, capture_output=True, text=True, check=True,
        )
        self.assertEqual(result.stdout.split("\n")[:2], ["False False", "(1, 1)"])


if __name__ == "__main__":
    unittest.main()
