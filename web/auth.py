"""Autenticacao do app.

Dois modos, que convivem:

- **Microsoft Entra (OIDC)**: modo principal quando `ENTRA_CLIENT_ID` e companhia
  estao definidos. Fluxo Authorization Code, aplicativo single-tenant.
- **Senha unica da equipe**: reserva, em `AUTO_PPT_TEAM_PASSWORD`. Continua
  existindo de proposito, para ninguem ficar sem acesso se o Entra estiver fora
  do ar ou mal configurado.

Sem nenhum dos dois configurados o app fica aberto, o que so faz sentido em
desenvolvimento local.

Em ambos os casos a sessao e um cookie assinado: guarda validade e, quando ha
login Microsoft, o e-mail de quem entrou. Nunca guarda senha nem token da
Microsoft.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import os
import time

SESSION_COOKIE = "qwst_session"
HANDSHAKE_COOKIE = "qwst_oidc"
SESSION_TTL_SECONDS = 12 * 60 * 60
HANDSHAKE_TTL_SECONDS = 10 * 60
# A MSAL adiciona openid, profile e offline_access sozinha e levanta ValueError
# se eles vierem na lista ("You cannot use any scope value that is reserved").
# Por isso passamos apenas o que nao e reservado.
SCOPES = ["email"]
_PUBLIC_PREFIXES = ("/static/", "/health", "/login", "/logout", "/auth/", "/favicon.ico")


# --------------------------------------------------------------------------
# configuracao
# --------------------------------------------------------------------------

def team_password() -> str:
    return os.getenv("AUTO_PPT_TEAM_PASSWORD", "").strip()


def entra_tenant_id() -> str:
    return os.getenv("ENTRA_TENANT_ID", "").strip()


def entra_client_id() -> str:
    return os.getenv("ENTRA_CLIENT_ID", "").strip()


def entra_client_secret() -> str:
    return os.getenv("ENTRA_CLIENT_SECRET", "").strip()


def entra_redirect_uri() -> str:
    return os.getenv("ENTRA_REDIRECT_URI", "").strip()


def entra_authority() -> str:
    """Derivada do tenant. Aceita ENTRA_AUTHORITY so como escape, mas o padrao
    evita a duplicidade de configuracao (e o erro de deixar o placeholder)."""
    explicit = os.getenv("ENTRA_AUTHORITY", "").strip()
    if explicit and "ENTRA_TENANT_ID" not in explicit:
        return explicit.rstrip("/")
    return f"https://login.microsoftonline.com/{entra_tenant_id()}"


def entra_enabled() -> bool:
    return bool(entra_tenant_id() and entra_client_id() and entra_client_secret() and entra_redirect_uri())


def team_password_enabled() -> bool:
    return bool(team_password())


def auth_enabled() -> bool:
    return entra_enabled() or team_password_enabled()


def config_problems() -> list[str]:
    """Erros de configuracao que so apareceriam no meio do login. Melhor
    mostrar cedo do que deixar o usuario bater num erro da Microsoft."""
    problems: list[str] = []
    partial = [
        name
        for name, value in (
            ("ENTRA_TENANT_ID", entra_tenant_id()),
            ("ENTRA_CLIENT_ID", entra_client_id()),
            ("ENTRA_CLIENT_SECRET", entra_client_secret()),
            ("ENTRA_REDIRECT_URI", entra_redirect_uri()),
        )
        if not value
    ]
    if partial and len(partial) < 4:
        problems.append("Configuração da Microsoft incompleta: falta " + ", ".join(partial) + ".")
    redirect = entra_redirect_uri()
    if redirect and not redirect.startswith("https://"):
        problems.append(f"ENTRA_REDIRECT_URI deve começar com https:// (valor atual: {redirect}).")
    # Pega o erro classico de colar o endereco duas vezes: "https://https//host".
    # Nao basta contar "://", porque a segunda ocorrencia costuma vir sem os
    # dois-pontos.
    host_and_path = redirect.split("://", 1)[1] if "://" in redirect else redirect
    if redirect.count("://") > 1 or host_and_path.lower().startswith("http"):
        problems.append(f"ENTRA_REDIRECT_URI está malformado: {redirect}")
    return problems


# --------------------------------------------------------------------------
# sessao
# --------------------------------------------------------------------------

def _session_key() -> bytes:
    """Chave de assinatura dos cookies.

    Prioriza AUTO_PPT_SESSION_SECRET. Sem ele, cai na senha da equipe, para o
    modo senha continuar funcionando sem configuracao extra. Derivar (em vez de
    guardar a chave) mantem varias instancias validando o mesmo cookie."""
    secret = os.getenv("AUTO_PPT_SESSION_SECRET", "").strip() or team_password()
    return hashlib.sha256(("qwst-auto-ppt-session:" + secret).encode("utf-8")).digest()


def _sign(payload: str) -> str:
    digest = hmac.new(_session_key(), payload.encode("utf-8"), hashlib.sha256).digest()
    return base64.urlsafe_b64encode(digest).decode("ascii").rstrip("=")


def _encode(value: str) -> str:
    return base64.urlsafe_b64encode(value.encode("utf-8")).decode("ascii").rstrip("=")


def _decode(value: str) -> str:
    padding = "=" * (-len(value) % 4)
    return base64.urlsafe_b64decode(value + padding).decode("utf-8")


def issue_session_token(subject: str = "", now: float | None = None) -> str:
    expires_at = int((now if now is not None else time.time()) + SESSION_TTL_SECONDS)
    return session_token_for(expires_at, subject)


def session_token_for(expires_at: int, subject: str = "") -> str:
    encoded = _encode(subject or "")
    payload = f"{expires_at}:{encoded}"
    return f"{expires_at}.{encoded}.{_sign(payload)}"


def session_token_valid(token: str, now: float | None = None) -> bool:
    parts = (token or "").split(".")
    if len(parts) != 3:
        return False
    expires_raw, encoded, _signature = parts
    if not expires_raw.isdigit():
        return False
    if not hmac.compare_digest(token, session_token_for(int(expires_raw), _safe_decode(encoded))):
        return False
    return int(expires_raw) > (now if now is not None else time.time())


def session_subject(token: str) -> str:
    parts = (token or "").split(".")
    if len(parts) != 3:
        return ""
    return _safe_decode(parts[1])


def _safe_decode(value: str) -> str:
    try:
        return _decode(value)
    except Exception:
        return ""


def password_matches(candidate: str) -> bool:
    expected = team_password()
    if not expected:
        return False
    return hmac.compare_digest(candidate.strip().encode("utf-8"), expected.encode("utf-8"))


def path_is_public(path: str) -> bool:
    return path.startswith(_PUBLIC_PREFIXES)


def request_is_authenticated(cookies: dict) -> bool:
    if not auth_enabled():
        return True
    return session_token_valid(cookies.get(SESSION_COOKIE, ""))


def current_user(cookies: dict) -> str:
    token = cookies.get(SESSION_COOKIE, "")
    return session_subject(token) if session_token_valid(token) else ""


# --------------------------------------------------------------------------
# handshake OIDC (state + nonce)
# --------------------------------------------------------------------------

def issue_handshake_token(state: str, nonce: str, destination: str, now: float | None = None) -> str:
    """State e nonce viajam num cookie assinado, nao em memoria do processo:
    assim o login sobrevive a reinicio do container e a mais de uma instancia."""
    expires_at = int((now if now is not None else time.time()) + HANDSHAKE_TTL_SECONDS)
    encoded = _encode(f"{state}|{nonce}|{destination}")
    payload = f"{expires_at}:{encoded}"
    return f"{expires_at}.{encoded}.{_sign(payload)}"


def read_handshake_token(token: str, now: float | None = None) -> tuple[str, str, str] | None:
    parts = (token or "").split(".")
    if len(parts) != 3 or not parts[0].isdigit():
        return None
    expires_at = int(parts[0])
    payload = f"{expires_at}:{parts[1]}"
    if not hmac.compare_digest(parts[2], _sign(payload)):
        return None
    if expires_at <= (now if now is not None else time.time()):
        return None
    decoded = _safe_decode(parts[1])
    pieces = decoded.split("|", 2)
    if len(pieces) != 3:
        return None
    return pieces[0], pieces[1], pieces[2]
