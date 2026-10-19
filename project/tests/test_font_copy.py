"""Tests for source-font copying at write time and font unification."""

import unittest

from openpyxl import Workbook
from openpyxl.styles import Font

from engine.font import apply_uniform_font
from engine.writer import copy_source_font


class CopySourceFontTests(unittest.TestCase):
    def setUp(self):
        self.source_book = Workbook()
        self.source_sheet = self.source_book.active
        self.target_book = Workbook()
        self.target_sheet = self.target_book.active

    def test_full_font_is_copied_like_manual_paste(self):
        self.source_sheet["A1"].font = Font(
            name="Calibri", size=14, bold=True, color="FFFF0000"
        )
        self.target_sheet["A1"].font = Font(name="宋体", size=11)

        copied = copy_source_font(
            self.source_sheet["A1"], self.target_sheet["A1"]
        )

        self.assertTrue(copied)
        target_font = self.target_sheet["A1"].font
        self.assertEqual(target_font.name, "Calibri")
        self.assertEqual(target_font.size, 14)
        self.assertTrue(target_font.bold)
        self.assertEqual(target_font.color.rgb, "FFFF0000")

    def test_default_style_source_leaves_target_font_untouched(self):
        self.target_sheet["A1"].font = Font(name="宋体", size=11)

        copied = copy_source_font(
            self.source_sheet["A1"], self.target_sheet["A1"]
        )

        self.assertFalse(copied)
        target_font = self.target_sheet["A1"].font
        self.assertEqual(target_font.name, "宋体")
        self.assertEqual(target_font.size, 11)
        self.assertIsNone(target_font.color)

    def test_copied_font_is_independent_of_source_workbook(self):
        self.source_sheet["A1"].font = Font(
            name="Calibri", size=14, color="FFFF0000"
        )

        copy_source_font(self.source_sheet["A1"], self.target_sheet["A1"])
        # 之后修改来源字体，不应影响已复制的目标单元格
        self.source_sheet["A1"].font = Font(name="Arial", size=9)

        target_font = self.target_sheet["A1"].font
        self.assertEqual(target_font.name, "Calibri")
        self.assertEqual(target_font.size, 14)
        self.assertEqual(target_font.color.rgb, "FFFF0000")


class ApplyUniformFontColorPreservationTests(unittest.TestCase):
    def _make_config(self):
        return {
            "font": {
                "enabled": True,
                "name": "宋体",
                "size": 11,
                "bold": False,
            },
            "reports": {
                "report7": {
                    "sheet_name": "报表7",
                    "data_start_col": "A",
                    "data_end_col": "C",
                    "sub_tables": [
                        {
                            "name": "子表",
                            "data_start_row": 4,
                            "data_end_row": 5,
                            "total_row": 6,
                        }
                    ],
                }
            },
        }

    def test_uniform_font_keeps_red_cells_red_and_normalizes_name_size(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "报表7"
        sheet["A4"] = "园区A"
        sheet["B4"] = "运营方A"
        sheet["C4"] = 100
        # 红色标注：最新调整数据
        for coordinate in ("A4", "B4", "C4"):
            sheet[coordinate].font = Font(
                name="Calibri", size=14, bold=True, color="FFFF0000"
            )
        # 普通黑色单元格
        sheet["A5"] = "园区B"
        sheet["A5"].font = Font(name="Calibri", size=14)

        modified = apply_uniform_font(workbook, self._make_config())

        self.assertGreater(modified, 0)
        for coordinate in ("A4", "B4", "C4"):
            font = sheet[coordinate].font
            self.assertEqual(font.name, "宋体")
            self.assertEqual(font.size, 11)
            self.assertFalse(font.bold)
            self.assertEqual(font.color.rgb, "FFFF0000")
        # 无颜色单元格统一后仍为默认（黑色）
        self.assertEqual(sheet["A5"].font.name, "宋体")
        self.assertEqual(sheet["A5"].font.size, 11)
        self.assertIsNone(sheet["A5"].font.color)

    def test_uniform_font_applies_report_override_and_keeps_color(self):
        config = self._make_config()
        config["font"]["report_overrides"] = {7: {"name": "微软雅黑", "size": 12}}
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "报表7"
        sheet["B4"].font = Font(name="Calibri", size=14, color="FFFF0000")

        apply_uniform_font(workbook, config)

        font = sheet["B4"].font
        self.assertEqual(font.name, "微软雅黑")
        self.assertEqual(font.size, 12)
        self.assertEqual(font.color.rgb, "FFFF0000")


if __name__ == "__main__":
    unittest.main()
