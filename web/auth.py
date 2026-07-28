"""Login por senha unica da equipe.

Protege o app quando ele fica exposto numa URL publica (App Runner). Nao ha
cadastro de usuarios: existe uma senha compartilhada, definida em
`AUTO_PPT_TEAM_PASSWORD`. Quem acerta recebe um cookie de sessao assinado.

Decisoes:
- A senha NUNCA vai para o cookie. O cookie guarda so a validade e uma
  assinatura HMAC dela.
- A chave de assinatura e derivada da propria senha, entao varias instancias do
  container aceitam o mesmo cookie sem precisar compartilhar outro segredo.
- Comparacoes usam `hmac.compare_digest` (tempo constante).
- Sem senha configurada o app fica aberto, e o /health avisa. Isso mantem o
  desenvolvimento local sem atrito, mas nunca deve ser o caso em producao.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import os
import time

SESSION_COOKIE = "qwst_session"
SESSION_TTL_SECONDS = 12 * 60 * 60
_PUBLIC_PREFIXES = ("/static/", "/health", "/login", "/logout", "/favicon.ico")


def team_password() -> str:
    return os.getenv("AUTO_PPT_TEAM_PASSWORD", "").strip()


def auth_enabled() -> bool:
    return bool(team_password())


def _session_key() -> bytes:
    """Chave de assinatura derivada da senha compartilhada."""
    return hashlib.sha256(("qwst-auto-ppt-session:" + team_password()).encode("utf-8")).digest()


def password_matches(candidate: str) -> bool:
    expected = team_password()
    if not expected:
        return False
    return hmac.compare_digest(candidate.strip().encode("utf-8"), expected.encode("utf-8"))


def issue_session_token(now: float | None = None) -> str:
    expires_at = int((now if now is not None else time.time()) + SESSION_TTL_SECONDS)
    payload = str(expires_at).encode("utf-8")
    signature = hmac.new(_session_key(), payload, hashlib.sha256).digest()
    return f"{expires_at}.{base64.urlsafe_b64encode(signature).decode('ascii').rstrip('=')}"


def session_token_valid(token: str, now: float | None = None) -> bool:
    if not token or "." not in token:
        return False
    expires_raw, _, signature_raw = token.partition(".")
    if not expires_raw.isdigit():
        return False
    expected = issue_session_token_for_expiry(int(expires_raw))
    if not hmac.compare_digest(token, expected):
        return False
    return int(expires_raw) > (now if now is not None else time.time())


def issue_session_token_for_expiry(expires_at: int) -> str:
    payload = str(expires_at).encode("utf-8")
    signature = hmac.new(_session_key(), payload, hashlib.sha256).digest()
    return f"{expires_at}.{base64.urlsafe_b64encode(signature).decode('ascii').rstrip('=')}"


def path_is_public(path: str) -> bool:
    return path.startswith(_PUBLIC_PREFIXES)


def request_is_authenticated(cookies: dict) -> bool:
    if not auth_enabled():
        return True
    return session_token_valid(cookies.get(SESSION_COOKIE, ""))
