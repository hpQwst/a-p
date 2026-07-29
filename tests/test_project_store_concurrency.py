import json
import os
import tempfile
import unittest
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from unittest.mock import MagicMock, patch

from ppt_automator import project_store


class AtomicWriteTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tempdir = tempfile.TemporaryDirectory()
        self._previous_root = os.environ.get("AUTO_PPT_DATA_ROOT")
        os.environ["AUTO_PPT_DATA_ROOT"] = self._tempdir.name

    def tearDown(self) -> None:
        if self._previous_root is None:
            os.environ.pop("AUTO_PPT_DATA_ROOT", None)
        else:
            os.environ["AUTO_PPT_DATA_ROOT"] = self._previous_root
        self._tempdir.cleanup()

    def test_atomic_write_leaves_no_partial_file_and_no_temp_leftovers(self) -> None:
        path = Path(self._tempdir.name) / "nested" / "data.json"
        project_store._atomic_write_bytes(path, b'{"a": 1}')
        project_store._atomic_write_bytes(path, b'{"a": 2}')

        self.assertEqual(json.loads(path.read_text(encoding="utf-8")), {"a": 2})
        leftovers = [item.name for item in path.parent.iterdir() if item.name.endswith(".tmp")]
        self.assertEqual(leftovers, [])

    def test_write_still_lands_when_the_os_refuses_to_replace(self) -> None:
        """No Windows o antivirus segura arquivos grandes recem-escritos e o
        os.replace falha com PermissionError. A gravacao nao pode se perder."""
        path = Path(self._tempdir.name) / "travado.json"
        path.write_bytes(b'{"antes": true}')
        original_replace = os.replace

        def always_denied(src, dst):
            raise PermissionError(5, "Acesso negado")

        os.replace = always_denied
        try:
            project_store.LOCK_TIMEOUT_SECONDS  # garante import intacto
            project_store._atomic_write_bytes(path, b'{"depois": true}')
        finally:
            os.replace = original_replace

        self.assertEqual(json.loads(path.read_text(encoding="utf-8")), {"depois": True})
        leftovers = [item.name for item in path.parent.iterdir() if item.name.endswith(".tmp")]
        self.assertEqual(leftovers, [])

    def test_concurrent_mapping_template_saves_do_not_lose_entries(self) -> None:
        project = project_store.create_project("squad1", "Projeto concorrencia")

        def save(index: int):
            return project_store.save_mapping_template(
                project,
                name="Template compartilhado",
                slug="template-compartilhado",
                entries={f"S001_T{index:03d}_CHART": f"fonte_{index}.xlsx"},
                metadata={f"chave_{index}": str(index)},
            )

        with ThreadPoolExecutor(max_workers=8) as pool:
            list(pool.map(save, range(8)))

        template = project_store.load_mapping_template("squad1", "template-compartilhado")
        self.assertIsNotNone(template)
        # O ultimo a gravar define as entries, mas o arquivo tem que continuar
        # valido e os metadados acumulados de todos (a fusao nao pode se perder).
        self.assertEqual(len(template["entries"]), 1)
        self.assertEqual(len(template["metadata"]), 8)

    def test_concurrent_memory_corrections_keep_every_append(self) -> None:
        project = project_store.create_project("squad2", "Projeto correcoes")

        def append(index: int):
            project_store.append_memory_correction(project, {"target": f"T{index:03d}"})

        with ThreadPoolExecutor(max_workers=8) as pool:
            list(pool.map(append, range(12)))

        corrections = project_store.load_memory_corrections(project)
        self.assertEqual(len(corrections), 12)
        self.assertEqual(
            sorted(item["target"] for item in corrections),
            sorted(f"T{index:03d}" for index in range(12)),
        )


class S3RetentionTaggingTests(unittest.TestCase):
    def setUp(self) -> None:
        self.project = project_store.ProjectRef(
            project_id="squad4|projeto",
            squad="squad4",
            slug="projeto",
            name="Projeto",
            description="",
            created_at="2026-07-29T00:00:00Z",
            updated_at="2026-07-29T00:00:00Z",
            backend="s3",
        )

    def test_run_artifacts_receive_retention_tag(self) -> None:
        client = MagicMock()
        with (
            patch.dict(
                os.environ,
                {
                    "AUTO_PPT_STORAGE_BACKEND": "s3",
                    "AUTO_PPT_S3_BUCKET": "squad4e5-state",
                    "AUTO_PPT_S3_PREFIX": "auto-ppt",
                },
                clear=False,
            ),
            patch.object(project_store, "_s3_client", return_value=client),
        ):
            project_store.save_project_bytes(
                self.project,
                ["runs", "run-1"],
                "resultado.pptx",
                b"ppt",
            )

        self.assertEqual(client.put_object.call_args.kwargs["Tagging"], "retention=run")

    def test_checkpoint_is_not_marked_for_expiration(self) -> None:
        client = MagicMock()
        with (
            patch.dict(
                os.environ,
                {
                    "AUTO_PPT_STORAGE_BACKEND": "s3",
                    "AUTO_PPT_S3_BUCKET": "squad4e5-state",
                    "AUTO_PPT_S3_PREFIX": "auto-ppt",
                },
                clear=False,
            ),
            patch.object(project_store, "_s3_client", return_value=client),
        ):
            project_store.save_project_bytes(
                self.project,
                ["checkpoint"],
                "input.pptx",
                b"ppt",
            )

        self.assertNotIn("Tagging", client.put_object.call_args.kwargs)


if __name__ == "__main__":
    unittest.main()
