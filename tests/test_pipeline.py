from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from pipeline.model import PipelineConfig
from pipeline.artifact_graph import discover_python_artifacts
from pipeline.state import PipelineState
from regenerate_python_office import regenerate_python


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


if __name__ == "__main__":
    unittest.main()
