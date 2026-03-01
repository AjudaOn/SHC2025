import json

from flask_login import current_user
from sqlalchemy import text

from models import db


def log_reserva_etapa(reserva_id: int, etapa: str, acao: str, details=None, actor_name: str | None = None) -> None:
    """
    Best-effort audit log: nunca deve quebrar o fluxo principal.
    """
    if not reserva_id or not etapa or not acao:
        return

    user_id = None
    username = None
    is_authenticated = False
    try:
        if actor_name:
            username = actor_name
        elif current_user and getattr(current_user, "is_authenticated", False):
            is_authenticated = True
            user_id = getattr(current_user, "id", None)
            username = getattr(current_user, "username", None)
    except Exception:
        user_id = None
        username = None
        is_authenticated = False

    if not username and not is_authenticated and user_id is None:
        username = "Internet"

    details_json = None
    if details is not None:
        try:
            details_json = json.dumps(details, ensure_ascii=False, default=str)
        except Exception:
            details_json = None

    try:
        with db.engine.begin() as conn:
            conn.execute(
                text(
                    """
                    INSERT INTO reserva_etapa_audit (
                        reserva_id, etapa, acao,
                        user_id, username,
                        created_at, details
                    )
                    VALUES (
                        :reserva_id, :etapa, :acao,
                        :user_id, :username,
                        CURRENT_TIMESTAMP, :details
                    )
                    """
                ),
                {
                    "reserva_id": int(reserva_id),
                    "etapa": str(etapa),
                    "acao": str(acao),
                    "user_id": int(user_id) if user_id is not None else None,
                    "username": str(username) if username else None,
                    "details": details_json,
                },
            )
    except Exception:
        return
