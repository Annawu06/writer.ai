import json
import unittest
from unittest.mock import ANY, patch

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
    def test_extracts_fenced_json_and_repairs_trailing_comma(self):
        content = 'Done:\n```json\n{"all_pages":{"insert_text":"a,}","bold":true,},}\n```'
        self.assertEqual(main.parse_model_json(content), {
            "all_pages": {"insert_text": "a,}", "bold": True}
        })

    def test_rejects_invalid_numeric_and_color_values(self):
        report = []
        self.assertEqual(main.validate_format_request({
            "all_pages": {"font_size": 0, "font_color": "invalid", "bold": True},
            "tables": {"table_1": {"column_widths": [float("nan"), 50]}},
        }, report), {"all_pages": {"bold": True}})
        self.assertEqual(len(report), 3)

    def test_table_caption_title_and_invalid_merge_cells(self):
        self.assertEqual(main.validate_format_request({
            "tables": {"table_1": {
                "title": "Results", "merge_cells": [{"start": 4, "end": "B1"}],
            }}
        }), {"tables": {"table_1": {"title": "Results"}}})

    def test_invalid_nested_structure_is_reported(self):
        warnings = []
        self.assertEqual(main.validate_format_request({
            "paragraph_text": {"format": {"bold": True}}
        }, warnings), {})
        self.assertTrue(warnings)

    def test_warned_model_plan_is_not_applied(self):
        class Guard:
            def __init__(self):
                self.closed = False

            def close(self):
                self.closed = True

        class Message:
            def execute(self):
                pass

        job = object.__new__(main.MainJob)
        job._request_generation = 1
        job._request_in_progress = True
        job._request_callback = None
        job._async_callback = None
        job._request_timeout_timer = None
        job._request_guard = Guard()
        job._status_indicator = None
        job._message_box = lambda *args: Message()
        response = json.dumps({
            "request": {"all_pages": {"bold": True}},
            "warnings": ["Unsupported property"], "errors": [],
        })
        with patch.object(main, "execute_undoable_format") as execute:
            job._finish_format_request(object(), response, 1)
        execute.assert_not_called()
        self.assertTrue(job._request_guard is None)

    def test_document_scope_does_not_switch_to_selection_for_text_values(self):
        class Formatter:
            def get_document_cursor(self):
                return "document"

            def get_selection_cursor(self):
                raise AssertionError("Selection must not be used")

        plan = {"all_pages": {"insert_text": "selection"}}
        with patch.object(main, "apply_styles") as apply:
            main.execute_format_request(plan, Formatter())
        apply.assert_called_once_with(ANY, "document", plan["all_pages"])

    def test_content_changes_require_confirmation(self):
        self.assertFalse(main.changes_document_content({"all_pages": {"bold": True}}))
        self.assertTrue(main.changes_document_content({"paragraphs": {
            "contains": {"text": "old", "format": {"replace_text": "new"}}
        }}))
        self.assertTrue(main.changes_document_content({"tables": {
            "table_1": {"merge_cells": [{"start": "A1", "end": "B1"}]}
        }}))

    def test_undoable_format_rolls_back_partial_failure(self):
        class UndoManager:
            def __init__(self):
                self.actions = ["previous edit"]
                self.undone = False

            def getAllUndoActionTitles(self):
                return self.actions

            def enterUndoContext(self, title):
                self.title = title

            def leaveUndoContext(self):
                if self.changed:
                    self.actions.insert(0, self.title)

            def undo(self):
                self.undone = True
                self.actions.pop(0)

        manager = UndoManager()
        manager.changed = True
        with patch.object(main, "execute_format_request", side_effect=RuntimeError("UNO failed")):
            with self.assertRaisesRegex(RuntimeError, "UNO failed"):
                main.execute_undoable_format({"all_pages": {"bold": True}}, object(), manager)
        self.assertTrue(manager.undone)
        self.assertEqual(manager.actions, ["previous edit"])

        manager.changed = False
        manager.undone = False
        with patch.object(main, "execute_format_request", side_effect=RuntimeError("UNO failed")):
            with self.assertRaises(RuntimeError):
                main.execute_undoable_format({"all_pages": {"bold": True}}, object(), manager)
        self.assertFalse(manager.undone)

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
            self.assertEqual(body["max_completion_tokens"], 512)
            self.assertEqual(timeout, main.MainJob.REQUEST_TIMEOUT)
            return FakeResponse(payload)

        with patch.object(main, "urlopen", fake_urlopen):
            result = main.MainJob.call_model(
                "bold the document", "test-key", "test-model", "https://example.test/v1"
            )
        self.assertEqual(result, {"all_pages": {"bold": True}})

    def test_dashscope_qwen_requests_json_mode(self):
        payload = {"choices": [{"message": {"content": '{"all_pages":{"bold":true}}'}}]}

        def fake_urlopen(request, timeout):
            body = json.loads(request.data.decode("utf-8"))
            self.assertEqual(body["response_format"], {"type": "json_object"})
            return FakeResponse(payload)

        with patch.object(main, "urlopen", fake_urlopen):
            result = main.MainJob.call_model(
                "bold all", "test-key", "qwen-plus",
                "https://dashscope.aliyuncs.com/compatible-mode/v1",
            )
        self.assertEqual(result, {"all_pages": {"bold": True}})

    def test_local_provider_can_run_without_api_key(self):
        payload = {"choices": [{"message": {"content": '{"all_pages":{"bold":true}}'}}]}

        def fake_urlopen(request, timeout):
            self.assertEqual(request.full_url, "http://127.0.0.1:11434/v1/chat/completions")
            self.assertIsNone(request.get_header("Authorization"))
            return FakeResponse(payload)

        with patch.dict(main.os.environ, {"WRITER_AI_API_KEY": ""}), patch.object(main, "urlopen", fake_urlopen):
            result = main.MainJob.call_model("bold all", "saved-cloud-key", "local-model", "http://127.0.0.1:11434/v1")
            keyless_result = main.MainJob.call_model("bold all", "", "local-model", "http://127.0.0.1:11434/v1")
        self.assertEqual(result, {"all_pages": {"bold": True}})
        self.assertEqual(keyless_result, result)

    def test_remote_http_endpoint_is_rejected(self):
        errors = []
        with patch.object(main, "urlopen") as request:
            result = main.MainJob.call_model("bold all", "key", "model", "http://example.test/v1", error_report=errors)
        self.assertEqual(result, {})
        self.assertTrue(errors)
        request.assert_not_called()

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
        self.assertEqual(indicator.started, ("Writer.AI: analysing", 100))
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
