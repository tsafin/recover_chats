from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from pipeline.model import PipelineConfig
from pipeline.artifact_graph import discover_python_artifacts
from pipeline.state import PipelineState
from regenerate_python_office import regenerate_python
from restore_export import chat_view_url, render_chat_html, render_chat_viewer, split_chat_page
from check_restoration import Links, validate_restoration


class DependencySchedulerTest(unittest.TestCase):
    def test_ast_discovers_literal_inputs_and_outputs(self) -> None:
        spec = discover_python_artifacts(
            "from pathlib import Path\n"
            "payload = Path('input.png').read_bytes()\n"
            "Path('output.html').write_text(payload.hex())\n"
        )
        self.assertEqual(spec.inputs, {"input.png"})
        self.assertEqual(spec.outputs, {"output.html"})

    def test_media_consumer_is_retried_after_producer(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "conversations"
            conversation = root / "synthetic_conversation"
            sources = conversation / "files"
            sources.mkdir(parents=True)
            # Alphabetical order deliberately places the consumer first.
            (sources / "a_build_html.py").write_text(
                "import base64\n"
                "from pathlib import Path\n"
                "payload = base64.b64encode(Path('input.png').read_bytes()).decode()\n"
                "Path('output.html').write_text('<img src=\"data:image/png;base64,' + payload + '\">')\n",
                encoding="utf-8",
            )
            (sources / "z_make_image.py").write_text(
                "from PIL import Image\n"
                "Image.new('RGB', (2, 2), 'red').save('input.png')\n",
                encoding="utf-8",
            )

            report, summary = regenerate_python(root, max_passes=4)

            output = conversation / "regenerated" / "output.html"
            self.assertTrue(output.is_file())
            self.assertIn("data:image/png;base64,", output.read_text(encoding="utf-8"))
            consumer = next(item for item in report if item["source"].endswith("a_build_html.py"))
            self.assertEqual(consumer["status"], "ok")
            self.assertEqual(consumer["attempts"], 2)
            self.assertEqual(summary["failed"], 0)


class PipelineStateTest(unittest.TestCase):
    def test_state_round_trip_and_config_normalization(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            config = PipelineConfig(root, root / "output").normalized()
            state = PipelineState.load(config.output)
            state.record_stage("restore", {"status": "ok"})
            restored = PipelineState.load(config.output)
            self.assertEqual(restored.stages["restore"]["status"], "ok")
            self.assertTrue(config.source.is_absolute())
            self.assertTrue(config.output.is_absolute())


class ChatAttachmentRenderingTest(unittest.TestCase):
    def test_viewer_links_resolve_chat_data_and_relative_artifacts(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            folder = "Диалог & данные"
            chat = root / "conversations" / folder
            chat.mkdir(parents=True)
            (chat / "report.pdf").touch()
            (root / "viewer.html").write_text(render_chat_viewer("ru"), encoding="utf-8")
            index = root / "INDEX.html"
            index.write_text(f'<a href="{chat_view_url(folder, "ru")}">Chat</a>', encoding="utf-8")
            _, markup, _ = split_chat_page(render_chat_html({}, [], [], [], chat))
            self.assertNotIn("function startChat", markup)
            payload = {"pages": {"ru": '<a href="report.pdf">Report</a>'}}
            data = chat / "chat-data.js"
            data.write_text("window.restoredChatData = " + json.dumps(payload) + ";", encoding="utf-8")
            report = validate_restoration(root)
            self.assertEqual(report["local_links"], 2)
            self.assertEqual(report["broken_links"], [])
            data.unlink()
            report = validate_restoration(root)
            self.assertEqual(report["broken_links"][0]["reason"], "missing or invalid chat data")

    def test_long_chat_keeps_message_links_and_tool_logs_in_payload(self) -> None:
        messages = [{"uuid": str(i), "sender": "assistant", "content": [
            {"type": "text", "text": "Message"},
            {"type": "tool_use", "name": "Read", "input": {"file_path": "report.md"}},
        ]} for i in range(41)]
        with tempfile.TemporaryDirectory() as temporary:
            rendered = render_chat_html(
                {"chat_messages": messages},
                [{"original_path": "report.md", "local_path": "files/report.md"}],
                [], [], Path(temporary), "en",
            )
        self.assertNotIn('<article class="message', rendered)
        parser = Links()
        parser.feed(rendered)
        payload, = parser.message_payloads
        data = json.loads(payload)
        self.assertEqual(len(data), 41)
        markup, height = data[-1]
        self.assertIn('id="message-41"', markup)
        self.assertIn('<details class="tool-log">', markup)
        self.assertIn('<details class="tool-step">', markup)
        self.assertGreater(height, 0)
        parser.feed(markup)
        self.assertIn(("files/report.md", False), parser.hrefs)

    def test_recovered_attachment_covers_one_uuid_only_document_block(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            conversation = {
                "name": "Attachment duplicate",
                "uuid": "conversation-id",
                "created_at": "2026-01-01T00:00:00Z",
                "chat_messages": [{
                    "uuid": "message-id",
                    "sender": "human",
                    "created_at": "2026-01-01T00:00:00Z",
                    "content": [
                        {"type": "document", "file_uuid": "covered-document"},
                        {"type": "document", "file_uuid": "actually-missing-document"},
                    ],
                }],
            }
            attachments = [{
                "message_uuid": "message-id",
                "name": "",
                "file_type": "txt",
                "local_path": "attachments/input.extracted.txt",
                "bytes": 7,
                "lines": 1,
                "content": "payload",
            }]

            rendered = render_chat_html(
                conversation, [], [], attachments, Path(temporary), "en"
            )

            self.assertIn('class="input-attachment preview-button"', rendered)
            self.assertNotIn("covered-document", rendered)
            self.assertIn("actually-missing-document", rendered)
            self.assertEqual(rendered.count("binary data missing from export"), 1)


if __name__ == "__main__":
    unittest.main()
