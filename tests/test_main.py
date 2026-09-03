import json
import unittest
from unittest.mock import patch

import main


class FakeResponse:
    def __init__(self, payload):
        self.payload = payload

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False

    def read(self):
        return json.dumps(self.payload).encode("utf-8")


class MainLogicTests(unittest.TestCase):
    def test_validates_stable_locations_and_rejects_unknown_fields(self):
        request = {
            "paragraph_2": {"first_line_indent": 2},
            "paragraph_text": {
                "text": "摘要",
                "format": {"font_size": 16},
            },
            "tables": {
                "table_1": {
                    "header": True,
                    "cells": {"row_2_col_3": {"align_center": True}},
                }
            },
            "unknown_scope": {"bold": True},
        }
        self.assertEqual(
            main.validate_format_request(request),
            {
                "paragraph_2": {"first_line_indent": 2},
                "paragraph_text": {
                    "text": "摘要",
                    "format": {"font_size": 16},
                },
                "tables": {
                    "table_1": {
                        "header": True,
                        "cells": {"row_2_col_3": {"align_center": True}},
                    }
                },
            },
        )

    def test_column_name_conversion(self):
        self.assertEqual(main.Format._column_name(1), "A")
        self.assertEqual(main.Format._column_name(26), "Z")
        self.assertEqual(main.Format._column_name(27), "AA")

    def test_first_line_indent_conversion(self):
        class Cursor:
            CharHeight = 12
            ParaFirstLineIndent = 0

        cursor = Cursor()
        fmt = object.__new__(main.Format)
        fmt.set_first_line_indent(cursor, 2)
        self.assertEqual(cursor.ParaFirstLineIndent, 846)

    def test_openai_compatible_request(self):
        payload = {"choices": [{"message": {"content": '{"all_pages": {"bold": true}}'}}]}

        def fake_urlopen(request, timeout):
            self.assertEqual(request.full_url, "https://example.test/v1/chat/completions")
            self.assertEqual(request.get_header("Authorization"), "Bearer test-key")
            body = json.loads(request.data.decode("utf-8"))
            self.assertEqual(body["model"], "test-model")
            self.assertFalse(body["stream"])
            self.assertEqual(timeout, 60)
            return FakeResponse(payload)

        with patch.object(main, "urlopen", fake_urlopen):
            result = main.MainJob.call_model(
                "bold the document", "test-key", "test-model", "https://example.test/v1"
            )
        self.assertEqual(result, {"all_pages": {"bold": True}})

    def test_network_failure_returns_empty_request(self):
        with patch.object(main, "urlopen", side_effect=main.URLError("offline")):
            result = main.MainJob.call_model(
                "test", "test-key", "test-model", "https://example.test/v1"
            )
        self.assertEqual(result, {})

    def test_invalid_model_json_returns_no_request(self):
        def fake_urlopen(request, timeout):
            return FakeResponse({"choices": [{"message": {"content": "not json"}}]})

        with patch.object(main, "urlopen", fake_urlopen):
            result = main.MainJob.call_model(
                "test", "test-key", "test-model", "https://example.test/v1"
            )
        self.assertIsNone(result)

    def test_cancelled_response_is_ignored(self):
        job = object.__new__(main.MainJob)
        job._request_generation = 2
        job._request_in_progress = True
        job.cancel_format_request()
        self.assertFalse(job._request_in_progress)
        self.assertEqual(job._request_generation, 3)


if __name__ == "__main__":
    unittest.main()
