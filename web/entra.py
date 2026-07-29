"""Fluxo OIDC com o Microsoft Entra (Authorization Code, single-tenant).

A MSAL cuida da parte criptografica: baixa as chaves do tenant, valida
assinatura, emissor, audiencia e nonce do id_token. Aqui ficam apenas as
decisoes que sao do produto:

- so aceitar usuarios do NOSSO tenant (o `tid` do token tem que bater);
- nunca guardar o token da Microsoft, so o e-mail, no nosso cookie de sessao;
- state e nonce viajam em cookie assinado, nao em memoria do processo.
"""

from __future__ import annotations

from typing import Any

from . import auth


class EntraError(RuntimeError):
    """Falha no login Microsoft, com texto que pode ser mostrado ao usuario."""


def _client():
    try:
        import msal
    except ImportError as exc:  # pragma: no cover - dependencia declarada em requirements
        raise EntraError("Biblioteca msal nao instalada no servidor.") from exc

    return msal.ConfidentialClientApplication(
        client_id=auth.entra_client_id(),
        client_credential=auth.entra_client_secret(),
        authority=auth.entra_authority(),
    )


def authorization_url(state: str, nonce: str) -> str:
    return _client().get_authorization_request_url(
        scopes=auth.SCOPES,
        state=state,
        nonce=nonce,
        redirect_uri=auth.entra_redirect_uri(),
        prompt="select_account",
    )


def exchange_code(code: str, nonce: str) -> str:
    """Troca o code pelo id_token e devolve o e-mail de quem entrou."""
    result = _client().acquire_token_by_authorization_code(
        code,
        scopes=auth.SCOPES,
        redirect_uri=auth.entra_redirect_uri(),
        nonce=nonce,
    )
    if "error" in result:
        description = str(result.get("error_description") or result.get("error") or "")
        raise EntraError(_friendly_error(description))

    claims: dict[str, Any] = result.get("id_token_claims") or {}
    if not claims:
        raise EntraError("A Microsoft nao devolveu os dados do usuario.")

    # Aplicativo single-tenant: recusa quem vier de outro diretorio, mesmo que a
    # Microsoft tenha autenticado com sucesso.
    tenant = str(claims.get("tid") or "")
    if tenant != auth.entra_tenant_id():
        raise EntraError("Esta conta nao pertence à organização autorizada.")

    email = _email_from(claims)
    if not email:
        raise EntraError("Nao consegui identificar o e-mail desta conta.")
    return email


def _email_from(claims: dict[str, Any]) -> str:
    for key in ("preferred_username", "email", "upn"):
        value = str(claims.get(key) or "").strip()
        if value:
            return value
    return ""


def _friendly_error(description: str) -> str:
    text = description.lower()
    if "redirect_uri" in text:
        return (
            "O endereço de retorno não confere com o cadastrado no Entra. "
            "Confira ENTRA_REDIRECT_URI e o registro do aplicativo."
        )
    if "invalid_client" in text or "secret" in text:
        return "O segredo do aplicativo (ENTRA_CLIENT_SECRET) está inválido ou expirado."
    return "A Microsoft recusou o login: " + (description[:200] or "motivo não informado")
