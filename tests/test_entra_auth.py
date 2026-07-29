import os
import time
import unittest
from unittest.mock import patch

from fastapi.testclient import TestClient

from web import auth, entra

TENANT = "6d620aff-4c64-4458-bac3-2e502b255ee1"
REDIRECT = "https://app.example.com/auth/callback"

ENTRA_ENV = {
    "ENTRA_TENANT_ID": TENANT,
    "ENTRA_CLIENT_ID": "client-abc",
    "ENTRA_CLIENT_SECRET": "segredo",
    "ENTRA_REDIRECT_URI": REDIRECT,
    "AUTO_PPT_SESSION_SECRET": "chave-de-sessao-para-teste",
}


class EntraEnvTestCase(unittest.TestCase):
    extra_env: dict[str, str] = {}

    def setUp(self) -> None:
        self._patcher = patch.dict(os.environ, {**ENTRA_ENV, **self.extra_env}, clear=False)
        self._patcher.start()
        for name in ("AUTO_PPT_TEAM_PASSWORD",):
            if name not in {**ENTRA_ENV, **self.extra_env}:
                os.environ.pop(name, None)

    def tearDown(self) -> None:
        self._patcher.stop()

    def client(self) -> TestClient:
        from web.main import app

        return TestClient(app, follow_redirects=False)


class ConfigTests(EntraEnvTestCase):
    def test_authority_is_derived_from_tenant(self) -> None:
        self.assertEqual(auth.entra_authority(), f"https://login.microsoftonline.com/{TENANT}")

    def test_placeholder_authority_is_ignored(self) -> None:
        with patch.dict(os.environ, {"ENTRA_AUTHORITY": "https://login.microsoftonline.com/ENTRA_TENANT_ID"}):
            self.assertEqual(auth.entra_authority(), f"https://login.microsoftonline.com/{TENANT}")

    def test_malformed_redirect_uri_is_reported(self) -> None:
        with patch.dict(os.environ, {"ENTRA_REDIRECT_URI": "https://https//app.example.com/auth/callback"}):
            problems = " ".join(auth.config_problems())
        self.assertIn("malformado", problems)

    def test_partial_configuration_is_reported(self) -> None:
        with patch.dict(os.environ, {"ENTRA_CLIENT_SECRET": ""}):
            problems = " ".join(auth.config_problems())
        self.assertIn("ENTRA_CLIENT_SECRET", problems)


class RealAuthorizationUrlTests(EntraEnvTestCase):
    """Sem mock: a MSAL de verdade monta a URL. Pega erros que so aparecem na
    biblioteca, como escopo reservado."""

    def test_url_has_everything_microsoft_needs(self) -> None:
        from urllib.parse import parse_qs, urlparse

        url = entra.authorization_url("estado123", "nonce123")
        parsed = urlparse(url)
        query = parse_qs(parsed.query)

        self.assertEqual(parsed.netloc, "login.microsoftonline.com")
        self.assertTrue(parsed.path.startswith(f"/{TENANT}/"), "authority precisa ser do nosso tenant")
        self.assertEqual(query["client_id"][0], "client-abc")
        self.assertEqual(query["redirect_uri"][0], REDIRECT)
        self.assertEqual(query["response_type"][0], "code")
        self.assertEqual(query["state"][0], "estado123")
        self.assertEqual(query["nonce"][0], "nonce123")
        scopes = query["scope"][0].split()
        for expected in ("openid", "profile", "email"):
            self.assertIn(expected, scopes)

    def test_reserved_scopes_are_not_passed_to_msal(self) -> None:
        self.assertNotIn("openid", auth.SCOPES)
        self.assertNotIn("profile", auth.SCOPES)


class HandshakeTests(EntraEnvTestCase):
    def test_handshake_survives_roundtrip(self) -> None:
        token = auth.issue_handshake_token("st", "no", "/jobs/x")
        self.assertEqual(auth.read_handshake_token(token), ("st", "no", "/jobs/x"))

    def test_tampered_handshake_is_rejected(self) -> None:
        token = auth.issue_handshake_token("st", "no", "/")
        expires, payload, signature = token.split(".")
        forged = f"{expires}.{payload}x.{signature}"
        self.assertIsNone(auth.read_handshake_token(forged))

    def test_expired_handshake_is_rejected(self) -> None:
        token = auth.issue_handshake_token("st", "no", "/", now=time.time() - auth.HANDSHAKE_TTL_SECONDS - 5)
        self.assertIsNone(auth.read_handshake_token(token))


class SessionSubjectTests(EntraEnvTestCase):
    def test_session_carries_the_email(self) -> None:
        token = auth.issue_session_token("pessoa@qwst.co")
        self.assertTrue(auth.session_token_valid(token))
        self.assertEqual(auth.session_subject(token), "pessoa@qwst.co")

    def test_session_signed_with_another_secret_is_rejected(self) -> None:
        token = auth.issue_session_token("pessoa@qwst.co")
        with patch.dict(os.environ, {"AUTO_PPT_SESSION_SECRET": "outra-chave"}):
            self.assertFalse(auth.session_token_valid(token))


class TenantIsolationTests(EntraEnvTestCase):
    def _claims(self, tid: str) -> dict:
        return {"tid": tid, "preferred_username": "pessoa@qwst.co"}

    def test_user_from_another_tenant_is_refused(self) -> None:
        with patch("web.entra._client") as client:
            client.return_value.acquire_token_by_authorization_code.return_value = {
                "id_token_claims": self._claims("outro-tenant-9999")
            }
            with self.assertRaises(entra.EntraError) as caught:
                entra.exchange_code("code", "nonce")
        self.assertIn("organização autorizada", str(caught.exception))

    def test_user_from_our_tenant_is_accepted(self) -> None:
        with patch("web.entra._client") as client:
            client.return_value.acquire_token_by_authorization_code.return_value = {
                "id_token_claims": self._claims(TENANT)
            }
            self.assertEqual(entra.exchange_code("code", "nonce"), "pessoa@qwst.co")


class CallbackRouteTests(EntraEnvTestCase):
    def test_login_page_offers_microsoft(self) -> None:
        page = self.client().get("/login")
        self.assertEqual(page.status_code, 200)
        self.assertIn("/auth/login", page.text)

    def test_auth_login_redirects_to_microsoft(self) -> None:
        with patch("web.entra.authorization_url", return_value="https://login.microsoftonline.com/authorize?x=1"):
            response = self.client().get("/auth/login?next=%2F")
        self.assertEqual(response.status_code, 303)
        self.assertIn("login.microsoftonline.com", response.headers["location"])
        self.assertIn(auth.HANDSHAKE_COOKIE, response.cookies)

    def test_callback_without_handshake_cookie_is_refused(self) -> None:
        response = self.client().get("/auth/callback?code=abc&state=xyz")
        self.assertEqual(response.status_code, 400)

    def test_callback_with_mismatched_state_is_refused(self) -> None:
        client = self.client()
        client.cookies.set(auth.HANDSHAKE_COOKIE, auth.issue_handshake_token("estado-certo", "nonce", "/"))
        response = client.get("/auth/callback?code=abc&state=estado-errado")
        self.assertEqual(response.status_code, 400)
        self.assertNotIn(auth.SESSION_COOKIE, response.cookies)

    def test_callback_with_valid_state_starts_session(self) -> None:
        client = self.client()
        client.cookies.set(auth.HANDSHAKE_COOKIE, auth.issue_handshake_token("estado", "nonce", "/"))
        with patch("web.entra.exchange_code", return_value="pessoa@qwst.co"):
            response = client.get("/auth/callback?code=abc&state=estado")
        self.assertEqual(response.status_code, 303)
        self.assertIn(auth.SESSION_COOKIE, response.cookies)

    def test_protected_page_still_blocked_without_session(self) -> None:
        response = self.client().get("/")
        self.assertEqual(response.status_code, 303)
        self.assertIn("/login", response.headers["location"])


class FallbackPasswordTests(EntraEnvTestCase):
    extra_env = {"AUTO_PPT_TEAM_PASSWORD": "senha-reserva"}

    def test_password_still_works_alongside_microsoft(self) -> None:
        client = self.client()
        response = client.post("/login", data={"password": "senha-reserva", "next": "/"})
        self.assertEqual(response.status_code, 303)
        self.assertIn(auth.SESSION_COOKIE, response.cookies)

    def test_login_page_shows_both_options(self) -> None:
        page = self.client().get("/login")
        self.assertIn("/auth/login", page.text)
        self.assertIn("senha da equipe", page.text.lower())


if __name__ == "__main__":
    unittest.main()
