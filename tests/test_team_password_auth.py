import os
import time
import unittest

from fastapi.testclient import TestClient

from web import auth


class TeamPasswordAuthTests(unittest.TestCase):
    def setUp(self) -> None:
        self._previous = os.environ.get("AUTO_PPT_TEAM_PASSWORD")
        os.environ["AUTO_PPT_TEAM_PASSWORD"] = "senha-de-teste"

    def tearDown(self) -> None:
        if self._previous is None:
            os.environ.pop("AUTO_PPT_TEAM_PASSWORD", None)
        else:
            os.environ["AUTO_PPT_TEAM_PASSWORD"] = self._previous

    def _client(self) -> TestClient:
        from web.main import app

        return TestClient(app, follow_redirects=False)

    def test_token_survives_roundtrip_but_expired_token_is_rejected(self) -> None:
        self.assertTrue(auth.session_token_valid(auth.issue_session_token()))
        already_expired = auth.session_token_for(int(time.time()) - 10)
        self.assertFalse(auth.session_token_valid(already_expired))

    def test_tampered_token_is_rejected(self) -> None:
        """Estende a validade mantendo assinatura e formato: o token continua
        com as tres partes, entao so a assinatura pode reprova-lo."""
        expires_raw, payload, signature = auth.issue_session_token().split(".")
        forged = f"{int(expires_raw) + 86400}.{payload}.{signature}"
        self.assertEqual(len(forged.split(".")), 3)
        self.assertFalse(auth.session_token_valid(forged))

    def test_token_from_another_password_is_rejected(self) -> None:
        token = auth.issue_session_token()
        os.environ["AUTO_PPT_TEAM_PASSWORD"] = "outra-senha"
        self.assertFalse(auth.session_token_valid(token))

    def test_home_redirects_to_login_without_session(self) -> None:
        response = self._client().get("/")
        self.assertEqual(response.status_code, 303)
        self.assertIn("/login", response.headers["location"])

    def test_health_stays_public_for_the_load_balancer(self) -> None:
        self.assertEqual(self._client().get("/health/live").status_code, 200)

    def test_wrong_password_does_not_create_session(self) -> None:
        response = self._client().post("/login", data={"password": "errada", "next": "/"})
        self.assertEqual(response.status_code, 401)
        self.assertNotIn(auth.SESSION_COOKIE, response.cookies)

    def test_correct_password_opens_the_app(self) -> None:
        client = self._client()
        response = client.post("/login", data={"password": "senha-de-teste", "next": "/"})
        self.assertEqual(response.status_code, 303)
        self.assertIn(auth.SESSION_COOKIE, response.cookies)

        home = client.get("/")
        self.assertEqual(home.status_code, 200)

    def test_next_cannot_redirect_to_an_external_site(self) -> None:
        client = self._client()
        response = client.post(
            "/login",
            data={"password": "senha-de-teste", "next": "https://site-malicioso.example/x"},
        )
        self.assertEqual(response.headers["location"], "/")

    def test_app_stays_open_when_no_password_is_configured(self) -> None:
        os.environ.pop("AUTO_PPT_TEAM_PASSWORD", None)
        self.assertFalse(auth.auth_enabled())
        self.assertEqual(self._client().get("/").status_code, 200)


if __name__ == "__main__":
    unittest.main()
