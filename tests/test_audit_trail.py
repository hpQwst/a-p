import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from ppt_automator import project_store
from web import audit, auth


class FakeAuth:
    def __init__(self, enabled: bool, user: str = "") -> None:
        self._enabled = enabled
        self._user = user

    def auth_enabled(self) -> bool:
        return self._enabled

    def current_user(self, cookies) -> str:
        return self._user


class ActorTests(unittest.TestCase):
    def test_microsoft_login_gives_the_email(self) -> None:
        self.assertEqual(audit.actor_from({}, FakeAuth(True, "pessoa@qwst.co")), "pessoa@qwst.co")

    def test_shared_password_is_recorded_as_unidentified(self) -> None:
        actor = audit.actor_from({}, FakeAuth(True, ""))
        self.assertEqual(actor, audit.SHARED_PASSWORD_ACTOR)
        self.assertFalse(audit.is_identified(actor))

    def test_open_app_is_recorded_as_unidentified(self) -> None:
        actor = audit.actor_from({}, FakeAuth(False))
        self.assertEqual(actor, audit.ANONYMOUS_ACTOR)
        self.assertFalse(audit.is_identified(actor))

    def test_only_a_real_email_counts_as_identified(self) -> None:
        self.assertTrue(audit.is_identified("pessoa@qwst.co"))
        self.assertFalse(audit.is_identified(""))


class RecordTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tempdir = tempfile.TemporaryDirectory()
        self._previous = os.environ.get("AUTO_PPT_DATA_ROOT")
        os.environ["AUTO_PPT_DATA_ROOT"] = self._tempdir.name
        self.project = project_store.create_project("squad4", "Projeto auditoria")
        self.job_dir = Path(self._tempdir.name) / "job"
        self.job_dir.mkdir()
        (self.job_dir / "metadata.json").write_text(
            json.dumps(
                {
                    "job_id": "abc123",
                    "project": {"squad": self.project.squad, "slug": self.project.slug},
                }
            ),
            encoding="utf-8",
        )

    def tearDown(self) -> None:
        if self._previous is None:
            os.environ.pop("AUTO_PPT_DATA_ROOT", None)
        else:
            os.environ["AUTO_PPT_DATA_ROOT"] = self._previous
        self._tempdir.cleanup()

    def test_action_lands_in_the_project_history(self) -> None:
        audit.record(self.job_dir, "pessoa@qwst.co", "trocou_planilha_do_grafico", {"target": "S001_T001_CHART"})
        entries = project_store.load_memory_corrections(self.project)
        self.assertEqual(len(entries), 1)
        self.assertEqual(entries[0]["actor"], "pessoa@qwst.co")
        self.assertTrue(entries[0]["identified"])
        self.assertEqual(entries[0]["action"], "trocou_planilha_do_grafico")
        self.assertEqual(entries[0]["target"], "S001_T001_CHART")
        self.assertEqual(entries[0]["job_id"], "abc123")

    def test_every_action_is_kept_in_order(self) -> None:
        audit.record(self.job_dir, "a@qwst.co", "aprovou_grafico", {"target": "T1"})
        audit.record(self.job_dir, "b@qwst.co", "pulou_grafico", {"target": "T2"})
        entries = project_store.load_memory_corrections(self.project)
        self.assertEqual([item["actor"] for item in entries], ["a@qwst.co", "b@qwst.co"])

    def test_recording_never_breaks_the_user_action(self) -> None:
        """Auditoria e importante, mas nao mais que o trabalho: se o registro
        falhar, a operacao do usuario segue."""
        with patch("web.audit.append_memory_correction", side_effect=RuntimeError("s3 fora do ar")):
            audit.record(self.job_dir, "pessoa@qwst.co", "aprovou_slide", {"slide": 3})
        self.assertEqual(project_store.load_memory_corrections(self.project), [])

    def test_unknown_project_is_ignored_quietly(self) -> None:
        (self.job_dir / "metadata.json").write_text(
            json.dumps({"job_id": "x", "project": {"squad": "squad4", "slug": "nao-existe"}}),
            encoding="utf-8",
        )
        audit.record(self.job_dir, "pessoa@qwst.co", "aprovou_slide", {"slide": 1})
        self.assertEqual(project_store.load_memory_corrections(self.project), [])

    def test_actor_is_remembered_for_background_generation(self) -> None:
        audit.remember_actor(self.job_dir, "pessoa@qwst.co")
        self.assertEqual(audit.remembered_actor(self.job_dir), "pessoa@qwst.co")

    def test_remembering_keeps_the_rest_of_the_metadata(self) -> None:
        audit.remember_actor(self.job_dir, "pessoa@qwst.co")
        metadata = json.loads((self.job_dir / "metadata.json").read_text(encoding="utf-8"))
        self.assertEqual(metadata["job_id"], "abc123")
        self.assertEqual(metadata["project"]["slug"], self.project.slug)


if __name__ == "__main__":
    unittest.main()
