from __future__ import annotations

from io import BytesIO
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZIP_DEFLATED, ZipFile

import openpyxl
from fastapi.testclient import TestClient

from ppt_automator import project_store
from ppt_automator.engine import generate_updated_pptx
from ppt_automator.ppt_discovery import PptTarget
from ppt_automator.table_normalizer import build_transform_plans, normalize_to_target
from ppt_automator.xlsx_parser import parse_xlsx_table
from web import auth, main


class UserIsolationWebTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.env = patch.dict(
            os.environ,
            {
                "AUTO_PPT_STORAGE_BACKEND": "local",
                "AUTO_PPT_DATA_ROOT": self.tmp.name,
                "AUTO_PPT_TEAM_PASSWORD": "teste-local",
                "AUTO_PPT_TEAM_PASSWORD_ENABLED": "1",
                "AUTO_PPT_SESSION_SECRET": "segredo-de-teste",
                "AUTO_PPT_BOOTSTRAP_ADMINS": "hugo.rocha@qwst.co",
            },
            clear=False,
        )
        self.env.start()
        self.addCleanup(self.env.stop)
        self.original_runtime = main.RUNTIME_ROOT
        main.RUNTIME_ROOT = Path(self.tmp.name) / "jobs"
        self.addCleanup(setattr, main, "RUNTIME_ROOT", self.original_runtime)
        main.RUNTIME_ROOT.mkdir(parents=True, exist_ok=True)
        project_store.ensure_store()

    def client_for(self, email: str) -> TestClient:
        client = TestClient(main.app)
        client.cookies.set(auth.SESSION_COOKIE, auth.issue_session_token(email))
        return client

    def test_first_login_requires_one_time_squad_choice(self) -> None:
        email = "pessoa@qwst.co"
        project_store.ensure_user(email)
        client = self.client_for(email)

        response = client.get("/", follow_redirects=False)
        self.assertEqual(response.status_code, 303)
        self.assertEqual(response.headers["location"], "/choose-squad")
        self.assertIn("Qual é o seu squad?", client.get("/choose-squad").text)

        selected = client.post("/choose-squad", data={"squad": "squad3"}, follow_redirects=False)
        self.assertEqual(selected.status_code, 303)
        self.assertEqual(project_store.load_user(email).squad, "squad3")

        client.post("/choose-squad", data={"squad": "squad4"}, follow_redirects=False)
        self.assertEqual(project_store.load_user(email).squad, "squad3")

    def test_common_user_only_sees_and_opens_own_squad(self) -> None:
        squad1_project = project_store.create_project("squad1", "Projeto visivel")
        squad2_project = project_store.create_project("squad2", "Projeto secreto")
        user = project_store.ensure_user("squad1@qwst.co")
        project_store.update_user(user.email, squad="squad1")
        client = self.client_for(user.email)

        home = client.get("/")
        self.assertEqual(home.status_code, 200)
        self.assertIn("Projeto visivel", home.text)
        self.assertNotIn("Projeto secreto", home.text)

        forbidden = client.get(
            f"/projects/{squad2_project.squad}/{squad2_project.slug}/preview",
            follow_redirects=False,
        )
        self.assertEqual(forbidden.status_code, 403)
        self.assertIn("outro squad", forbidden.text)
        self.assertNotIn(squad1_project.slug, forbidden.text)

    def test_job_url_is_also_isolated_by_squad(self) -> None:
        user = project_store.ensure_user("squad1-job@qwst.co")
        project_store.update_user(user.email, squad="squad1")
        job_id = "a" * 32
        job_dir = main.RUNTIME_ROOT / job_id
        job_dir.mkdir()
        (job_dir / "metadata.json").write_text(
            json.dumps({"job_id": job_id, "project": {"squad": "squad2", "slug": "x"}}),
            encoding="utf-8",
        )

        response = self.client_for(user.email).get(
            f"/jobs/{job_id}/generation-status",
            headers={"Accept": "application/json"},
        )
        self.assertEqual(response.status_code, 403)
        self.assertEqual(response.json()["error"], "Este job pertence a outro squad.")

    def test_existing_job_without_readable_squad_is_denied(self) -> None:
        user = project_store.ensure_user("squad1-unknown-job@qwst.co")
        project_store.update_user(user.email, squad="squad1")
        for job_id, metadata in [
            ("c" * 32, None),
            ("d" * 32, "{json quebrado"),
            ("e" * 32, json.dumps({"project": {"squad": "squad9"}})),
            ("1" * 32, json.dumps({"project": "squad1"})),
        ]:
            with self.subTest(job_id=job_id):
                job_dir = main.RUNTIME_ROOT / job_id
                job_dir.mkdir()
                if metadata is not None:
                    (job_dir / "metadata.json").write_text(metadata, encoding="utf-8")
                response = self.client_for(user.email).get(
                    f"/jobs/{job_id}/generation-status",
                    headers={"Accept": "application/json"},
                )
                self.assertEqual(response.status_code, 403)
                self.assertEqual(
                    response.json()["error"],
                    "Nao foi possivel validar o squad deste job.",
                )

    def test_unknown_job_path_is_left_for_the_route_to_report_not_found(self) -> None:
        user = project_store.ensure_user("squad1-missing-job@qwst.co")
        project_store.update_user(user.email, squad="squad1")
        response = self.client_for(user.email).get(
            f"/jobs/{'f' * 32}/generation-status",
            headers={"Accept": "application/json"},
        )
        self.assertNotEqual(response.status_code, 403)

    def test_admin_selects_a_view_and_can_manage_users(self) -> None:
        project_store.create_project("squad1", "Projeto um")
        project_store.create_project("squad2", "Projeto dois")
        admin = project_store.ensure_user("hugo.rocha@qwst.co")
        target = project_store.ensure_user("lider@qwst.co")
        client = self.client_for(admin.email)

        home = client.get("/?squad=squad2")
        self.assertIn("Projeto dois", home.text)
        self.assertNotIn("Projeto um", home.text)
        self.assertEqual(client.get("/admin/users").status_code, 200)

        changed = client.post(
            "/admin/users/update",
            data={"email": target.email, "squad": "squad2", "role": "admin", "active": "1"},
            follow_redirects=False,
        )
        self.assertEqual(changed.status_code, 303)
        updated = project_store.load_user(target.email)
        self.assertEqual((updated.squad, updated.role, updated.active), ("squad2", "admin", True))
        self.assertTrue(list((Path(self.tmp.name) / "admin_audit").glob("*.json")))

    def test_disabled_user_sees_clear_block_page(self) -> None:
        user = project_store.ensure_user("bloqueado@qwst.co")
        project_store.update_user(user.email, squad="squad4", active=False)
        response = self.client_for(user.email).get("/")
        self.assertEqual(response.status_code, 403)
        self.assertIn("Seu acesso está bloqueado", response.text)

    def test_checkpoint_inputs_are_uploaded_once_and_manual_save_works(self) -> None:
        project = project_store.create_project("squad1", "Checkpoint economico")
        user = project_store.ensure_user("checkpoint@qwst.co")
        project_store.update_user(user.email, squad="squad1")
        job_id = "b" * 32
        job_dir = main.RUNTIME_ROOT / job_id
        job_dir.mkdir()
        (job_dir / "input.pptx").write_bytes(b"ppt-grande")
        (job_dir / "datasources.zip").write_bytes(b"dados")
        override_dir = job_dir / "overrides" / "slide-1-chart-1"
        override_dir.mkdir(parents=True)
        (override_dir / "ajuste.xlsx").write_bytes(b"xlsx-manual")
        (job_dir / "metadata.json").write_text(
            json.dumps(
                {
                    "job_id": job_id,
                    "project": {"squad": project.squad, "slug": project.slug, "name": project.name},
                    "files": {"pptx": "deck.pptx", "datasources": "dados.zip", "mapping": ""},
                    "slides": {"numbers": []},
                }
            ),
            encoding="utf-8",
        )

        main._save_project_checkpoint(job_dir, include_inputs=True, reason="preview_criado")
        stored_input = Path(self.tmp.name) / "squads" / "squad1" / "projects" / project.slug / "checkpoint" / "input.pptx"
        stored_override = (
            Path(self.tmp.name)
            / "squads"
            / "squad1"
            / "projects"
            / project.slug
            / "checkpoint"
            / "overrides"
            / "slide-1-chart-1"
            / "ajuste.xlsx"
        )
        first_mtime = stored_input.stat().st_mtime_ns
        first_override_mtime = stored_override.stat().st_mtime_ns
        response = self.client_for(user.email).post(
            f"/jobs/{job_id}/save",
            headers={"Accept": "application/json"},
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()["status"], "saved")
        self.assertEqual(stored_input.stat().st_mtime_ns, first_mtime)
        self.assertEqual(stored_override.stat().st_mtime_ns, first_override_mtime)
        checkpoint = project_store.load_project_json(project, ["checkpoint"], "checkpoint.json")
        self.assertEqual(checkpoint["save_reason"], "salvamento_manual")
        self.assertGreaterEqual(checkpoint["save_count"], 2)
        self.assertEqual(len(checkpoint["manual_overrides"]["slide-1-chart-1"]["sha256"]), 64)

    def test_generated_ppt_becomes_stale_after_manual_override_changes(self) -> None:
        job_dir = main.RUNTIME_ROOT / ("9" * 32)
        job_dir.mkdir()
        (job_dir / "input.pptx").write_bytes(b"ppt")
        (job_dir / "datasources.zip").write_bytes(b"zip")
        (job_dir / "metadata.json").write_text(
            json.dumps({"slides": {"numbers": [1]}, "project": {"squad": "squad1"}}),
            encoding="utf-8",
        )
        signature = main._generation_input_signature(job_dir)
        (job_dir / "generated.pptx").write_bytes(b"output")
        (job_dir / "generated.json").write_text(
            json.dumps({"file_name": "output.pptx", "input_signature": signature}),
            encoding="utf-8",
        )

        self.assertTrue(main._generated_is_current(job_dir))

        override_dir = job_dir / "overrides" / "target"
        override_dir.mkdir(parents=True)
        (override_dir / "manual.xlsx").write_bytes(b"new data")

        self.assertFalse(main._generated_is_current(job_dir))

    def test_preview_uses_before_after_tables_and_download_cleanup(self) -> None:
        template = (main.APP_ROOT / "templates" / "preview.html").read_text(encoding="utf-8")
        javascript = (main.APP_ROOT / "static" / "app.js").read_text(encoding="utf-8")
        stylesheet = (main.APP_ROOT / "static" / "style.css").read_text(encoding="utf-8")

        self.assertIn("Antes — estrutura atual no PowerPoint", template)
        self.assertIn("Depois — dados que serão gravados", template)
        self.assertNotIn("data-chart-canvas", template)
        self.assertIn("finishAsyncDownload(asyncDownload, state.download_url)", javascript)
        self.assertNotIn("window.location.assign(state.download_url)", javascript)
        self.assertIn("--accent-foreground", stylesheet)
        self.assertNotIn(':root[data-theme="dark"] button {', stylesheet)


class ProgressCallbackTests(unittest.TestCase):
    def test_preview_and_generation_use_the_same_completion_scale(self) -> None:
        self.assertEqual(main._object_progress_percent(1, 1, "matching"), 50)
        self.assertEqual(main._object_progress_percent(1, 1, "targets"), 50)
        self.assertEqual(main._object_progress_percent(1, 1, "complete"), 100)

    def test_preview_matching_reports_each_object(self) -> None:
        source = parse_xlsx_table(_workbook_bytes(), file_name="dados.xlsx")
        targets = [
            PptTarget(
                slide_index=0,
                slide_number=1,
                slide_path="ppt/slides/slide1.xml",
                shape_name=f"Tabela {index}",
                shape_id=str(index),
                object_type="table",
                left_in=0,
                top_in=0,
                width_in=1,
                height_in=1,
                table_cells=[["Indicador", "Valor"], ["Total", "0"]],
            )
            for index in (1, 2)
        ]
        events: list[dict] = []

        build_transform_plans(targets, [source], progress_callback=events.append)

        self.assertEqual([event["completed"] for event in events], [1, 2])
        self.assertTrue(all(event["total"] == 2 for event in events))
        self.assertEqual([event["target_id"] for event in events], [target.target_id for target in targets])

    def test_table_generation_reports_objects_packaging_and_completion(self) -> None:
        source = parse_xlsx_table(_workbook_bytes(), file_name="dados.xlsx")
        target = PptTarget(
            slide_index=0,
            slide_number=1,
            slide_path="ppt/slides/slide1.xml",
            shape_name="Tabela 1",
            shape_id="1",
            object_type="table",
            left_in=0,
            top_in=0,
            width_in=1,
            height_in=1,
            table_cells=[["Indicador", "Valor"], ["Total", "0"]],
        )
        plan = normalize_to_target(target, source)
        events: list[dict] = []

        output = generate_updated_pptx(
            _minimal_table_pptx(),
            [plan],
            targets=[target],
            progress_callback=events.append,
        )

        self.assertTrue(output)
        self.assertEqual([event["phase"] for event in events], ["targets", "targets", "packaging", "complete"])
        self.assertEqual(events[1]["completed"], 1)
        self.assertEqual(events[1]["total"], 1)


class DisabledPasswordTests(unittest.TestCase):
    def test_password_can_remain_stored_but_be_disabled(self) -> None:
        with patch.dict(
            os.environ,
            {"AUTO_PPT_TEAM_PASSWORD": "nao-usar", "AUTO_PPT_TEAM_PASSWORD_ENABLED": "0"},
            clear=False,
        ):
            self.assertFalse(auth.team_password_enabled())
            self.assertFalse(auth.password_matches("nao-usar"))


def _workbook_bytes() -> bytes:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["Indicador", "Valor"])
    sheet.append(["Total", 42])
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def _minimal_table_pptx() -> bytes:
    slide = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree><p:graphicFrame>
    <p:nvGraphicFramePr><p:cNvPr id="1" name="Tabela 1"/></p:nvGraphicFramePr>
    <a:graphic><a:graphicData><a:tbl>
      <a:tr h="1">
        <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Indicador</a:t></a:r></a:p></a:txBody></a:tc>
        <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Valor</a:t></a:r></a:p></a:txBody></a:tc>
      </a:tr>
      <a:tr h="1">
        <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Total</a:t></a:r></a:p></a:txBody></a:tc>
        <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>0</a:t></a:r></a:p></a:txBody></a:tc>
      </a:tr>
    </a:tbl></a:graphicData></a:graphic>
  </p:graphicFrame></p:spTree></p:cSld>
</p:sld>"""
    output = BytesIO()
    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        archive.writestr("ppt/slides/slide1.xml", slide)
    return output.getvalue()


if __name__ == "__main__":
    unittest.main()
