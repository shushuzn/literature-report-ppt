"""核心函数的回归测试。

运行： python -m unittest discover -s tests
"""
import os
import subprocess
import sys
import tempfile
import unittest

import yaml

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "tools"))

from fact_check import _number, numbers_in  # noqa: E402
from pptkit.compliance import check  # noqa: E402
from pptkit.engine import Deck  # noqa: E402
from pptkit.flow import page_texts  # noqa: E402
from pptkit.metrics import block_h  # noqa: E402
from pptkit.pages import page_min_pt, page_style, PAGES  # noqa: E402
from pptkit.style import Style  # noqa: E402
from pptkit.template import resolve_pages  # noqa: E402


class TestFactNumbers(unittest.TestCase):
    def test_no_substring_match(self):
        nums = numbers_in("比电容 10.7 F/g，循环 300 次")
        self.assertIn(_number("10.7"), nums)
        self.assertNotIn(_number("0.7"), nums)
        self.assertNotIn(_number("30"), nums)

    def test_equivalent_spellings(self):
        nums = numbers_in("孔径 0.70 nm，比表面积 1,700 m²/g")
        self.assertIn(_number("0.7"), nums)
        self.assertIn(_number("1700"), nums)


class TestMetrics(unittest.TestCase):
    def test_block_h_has_line_height(self):
        # 以前这里引用了不存在的 LINE_H，调用即 NameError。
        h = block_h("中文" * 10, 20, 3.0)
        self.assertGreater(h, 0.3)


class TestPageOverrides(unittest.TestCase):
    def test_page_floor_follows_page_body(self):
        self.assertEqual(page_min_pt({"body_pt": 22}, 16), 22)
        # 页级覆盖允许低于全篇下限（文档承诺的行为）。
        self.assertEqual(page_min_pt({"body_pt": 16}, 20), 16)

    def test_explicit_page_min_wins(self):
        self.assertEqual(page_min_pt({"body_pt": 16, "min_pt": 18}, 20), 18)

    def test_page_style_applies_both(self):
        st = page_style(Style(body_pt=20, min_pt=20), {"body_pt": 22})
        self.assertEqual((st.body_pt, st.min_pt), (22, 22))


class TestStyleBooleans(unittest.TestCase):
    def test_string_false_is_false(self):
        self.assertFalse(Style.from_config({"auto_pt": "false"}).auto_pt)

    def test_invalid_bool_raises(self):
        with self.assertRaises(ValueError):
            Style.from_config({"auto_pt": "maybe"})


class TestSaveAndCompliance(unittest.TestCase):
    def test_save_bare_filename(self):
        with tempfile.TemporaryDirectory() as d:
            old = os.getcwd()
            os.chdir(d)
            try:
                Deck().save("out.pptx")
                self.assertTrue(os.path.exists("out.pptx"))
            finally:
                os.chdir(old)

    def test_latin_only_run_needs_no_ea_font(self):
        st = Style()
        deck = Deck(style=st)
        deck.new_slide()
        deck.text(1.0, 1.0, 5.0, 1.0, [("Hello 2026", 20, False, st.ink)])
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "one.pptx")
            deck.save(path)
            findings, _ = check(path, style=st)
        self.assertFalse([f for f in findings if "中文字体" in f["item"]])


class TestTextFlowAndTemplate(unittest.TestCase):
    def test_compare_page_is_scanned(self):
        page = {"type": "compare",
                "left": {"label": "旧范式", "points": ["一条足够长的对照要点" * 3]},
                "right": {"label": "新范式", "points": ["另一条对照要点" * 3]}}
        rows = page_texts("", page, Style())
        self.assertTrue(any(tag.startswith("对照") for tag, _, _ in rows))

    def test_legacy_config_resolves(self):
        pages = resolve_pages({"cover": {"title_cn": "旧版封面"}})
        self.assertEqual(pages[0]["type"], "cover")

    def test_all_registered_pages_have_functions(self):
        self.assertEqual(len(PAGES), 10)
        for name, fn in PAGES.items():
            self.assertTrue(callable(fn), name)


class TestExportNotes(unittest.TestCase):
    def test_default_is_txt_and_uses_configured_speed(self):
        with tempfile.TemporaryDirectory() as d:
            cfg = {
                "out": os.path.join(d, "deck.pptx"),
                "compliance": {"limit_min": 5, "time_buffer": 0.4,
                               "chars_per_min": 200},
                "deck": [{"type": "cover", "title_cn": "封面",
                          "notes": "讲" * 100}],
            }
            cfg_path = os.path.join(d, "content.yaml")
            with open(cfg_path, "w", encoding="utf-8") as f:
                yaml.safe_dump(cfg, f, allow_unicode=True)
            r = subprocess.run(
                [sys.executable, os.path.join(ROOT, "tools", "export_notes.py"),
                 cfg_path],
                capture_output=True, text=True, encoding="utf-8")
            self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
            out = os.path.join(d, "讲稿.txt")
            self.assertTrue(os.path.exists(out))
            with open(out, encoding="utf-8") as f:
                text = f.read()
            self.assertIn("约 100 字 / 30 秒", text)


if __name__ == "__main__":
    unittest.main()
