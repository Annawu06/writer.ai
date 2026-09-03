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

    def test_format_preview_uses_user_friendly_summary(self):
        summary = main.summarize_format_request({
            "page_1": {"line_1": {"highlight": True}},
        })
        self.assertEqual(summary, ["第 1 页第 1 行：高亮（黄色）"])

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

    def test_format_removal_properties(self):
        class Cursor:
            CharWeight = main.BOLD
            CharPosture = main.ITALIC
            CharUnderline = 1
            CharUnderlineHasColor = True

        cursor = Cursor()
        formatter = object.__new__(main.Format)
        formatter.set_bold(cursor, False)
        formatter.set_italic(cursor, False)
        formatter.set_underline(cursor, False)
        self.assertEqual(cursor.CharWeight, main.NORMAL)
        self.assertEqual(cursor.CharPosture, main.NONE)
        self.assertEqual(cursor.CharUnderline, 0)
        self.assertFalse(cursor.CharUnderlineHasColor)

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
        errors = []
        with patch.object(main, "urlopen", side_effect=main.URLError("offline")):
            result = main.MainJob.call_model(
                "test", "test-key", "test-model", "https://example.test/v1", error_report=errors
            )
        self.assertEqual(result, {})
        self.assertTrue(errors)

    def test_invalid_model_json_returns_no_request(self):
        def fake_urlopen(request, timeout):
            return FakeResponse({"choices": [{"message": {"content": "not json"}}]})

        errors = []
        with patch.object(main, "urlopen", fake_urlopen):
            result = main.MainJob.call_model(
                "test", "test-key", "test-model", "https://example.test/v1", error_report=errors
            )
        self.assertIsNone(result)
        self.assertTrue(errors)

    def test_model_validation_report_lists_ignored_properties(self):
        payload = {
            "choices": [{
                "message": {
                    "content": '{"all_pages": {"bold": true, "not_a_format": true}}'
                }
            }]
        }
        report = []

        with patch.object(main, "urlopen", lambda request, timeout: FakeResponse(payload)):
            result = main.MainJob.call_model(
                "test", "test-key", "test-model", "https://example.test/v1", report
            )

        self.assertEqual(result, {"all_pages": {"bold": True}})
        self.assertEqual(len(report), 1)
        self.assertIn("not_a_format", report[0])

    def test_cancelled_response_is_ignored(self):
        job = object.__new__(main.MainJob)
        job._request_generation = 2
        job._request_in_progress = True
        job.cancel_format_request()
        self.assertFalse(job._request_in_progress)
        self.assertEqual(job._request_generation, 3)

    def test_status_indicator_lifecycle(self):
        class Indicator:
            def __init__(self):
                self.started = None
                self.ended = False

            def start(self, text, value):
                self.started = (text, value)

            def end(self):
                self.ended = True

        class Frame:
            def __init__(self, indicator):
                self.indicator = indicator

            def createStatusIndicator(self):
                return self.indicator

        class Controller:
            def __init__(self, frame):
                self.frame = frame

            def getFrame(self):
                return self.frame

        class Document:
            def __init__(self, controller):
                self.controller = controller

            def getCurrentController(self):
                return self.controller

        indicator = Indicator()
        job = object.__new__(main.MainJob)
        job._status_indicator = None
        job._start_status_indicator(Document(Controller(Frame(indicator))))
        self.assertEqual(indicator.started, ("writer.ai: 正在分析", 100))
        job._end_status_indicator()
        self.assertTrue(indicator.ended)

    def test_empty_settings_fields_preserve_saved_values(self):
        class Model:
            def __init__(self, text="", selected_items=()):
                self.Text = text
                self.SelectedItems = selected_items

        class Control:
            def __init__(self, model):
                self.model = model

            def getModel(self):
                return self.model

        job = object.__new__(main.MainJob)
        job._settings_initial_values = {
            "api_key": "saved-key",
            "model": "saved-model",
            "endpoint": "https://saved.example/v1",
        }
        controls = {
            "api_key": Control(Model()),
            "model": Control(Model()),
            "endpoint": Control(Model()),
        }
        self.assertEqual(
            job._read_dialog_config(controls),
            {
                "api_key": "saved-key",
                "model": "saved-model",
                "endpoint": "https://saved.example/v1",
            },
        )

    def test_api_key_reads_password_container_with_handler(self):
        class Record:
            UserList = [type("User", (), {"Passwords": ["stored-key"]})()]

        class Container:
            def findForName(self, url, user, handler):
                self.handler = handler
                return Record()

        class ServiceManager:
            def __init__(self, container, handler):
                self.container = container
                self.handler = handler

            def createInstanceWithContext(self, service, context):
                return self.container if service.endswith("PasswordContainer") else self.handler

        container = Container()
        handler = object()
        job = object.__new__(main.MainJob)
        job.ctx = object()
        job.sm = ServiceManager(container, handler)
        self.assertEqual(job.get_api_key(), "stored-key")
        self.assertIn(container.handler, (None, handler))


if __name__ == "__main__":
    unittest.main()
