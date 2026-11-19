import unittest

import format_plan


class FormatPlanTests(unittest.TestCase):
    def test_parser_and_validator_without_uno(self):
        response = '''```json
        {"paragraph_2": {"font_size": 12, "bold": true},
         "tables": {"table_1": {"header_background": "#EEEEEE"}},
         "unsupported": {"replace_text": "changed"}}
        ```'''
        warnings = []
        plan = format_plan.validate_format_request(
            format_plan.parse_model_json(response), warnings
        )
        self.assertEqual(plan, {
            "paragraph_2": {"font_size": 12, "bold": True},
            "tables": {"table_1": {"header_background": "#EEEEEE"}},
        })
        self.assertEqual(len(warnings), 1)
        self.assertEqual(format_plan.count_format_operations(plan), 3)
        self.assertFalse(format_plan.changes_document_content(plan))

    def test_reports_conflicting_paragraph_commands(self):
        warnings = []
        format_plan.validate_format_request({"paragraph_1": {
            "align_left": True, "align_center": True,
            "title": True, "body": True,
        }}, warnings)
        self.assertEqual(len(warnings), 2)
        self.assertIn("alignments", warnings[0])
        self.assertIn("styles", warnings[1])

    def test_table_caption_is_content_change(self):
        self.assertTrue(format_plan.changes_document_content({
            "tables": {"table_1": {"caption": "Results"}}
        }))
        self.assertTrue(format_plan.changes_document_content({
            "tables": {"table_1": {"title": "Results"}}
        }))
        self.assertFalse(format_plan.changes_document_content({
            "paragraph_1": {"title": True}
        }))

    def test_rejects_layout_dependent_page_lines(self):
        warnings = []
        plan = format_plan.validate_format_request({
            "page_1": {"line_2": {"bold": True}}
        }, warnings)
        self.assertEqual(plan, {})
        self.assertIn("unsupported formatting scope: page_1", warnings[0].lower())

    def test_count_excludes_target_metadata_and_includes_structure_style(self):
        self.assertEqual(format_plan.count_format_operations({
            "paragraph_text": {"text": "unique target", "format": {"bold": True}},
            "structure": {"infer": True, "heading_1": True, "body": {"font_size": 12}},
            "tables": {"table_1": {"header": True, "cells": {
                "row_2_col_1": {"italic": True}
            }}},
        }), 6)


if __name__ == "__main__":
    unittest.main()
