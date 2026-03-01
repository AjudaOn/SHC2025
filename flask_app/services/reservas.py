from datetime import date

from models import db, BaseDados
from constants import (
    STATUS_RESERVA_APROVADA,
    STATUS_RESERVA_EXPIRADA,
    STATUS_RESERVA_PENDENTE,
)


def expire_reservas(today=None, include_pendente=True):
    """Marca reservas aprovadas/pendentes como expiradas quando a data de entrada passou."""
    if today is None:
        today = date.today()

    statuses = [STATUS_RESERVA_APROVADA]
    if include_pendente:
        statuses.append(STATUS_RESERVA_PENDENTE)

    updated = BaseDados.query.filter(
        BaseDados.status_reserva.in_(statuses),
        BaseDados.entrada.isnot(None),
        BaseDados.entrada < today,
    ).update({BaseDados.status_reserva: STATUS_RESERVA_EXPIRADA}, synchronize_session=False)

    if updated:
        db.session.commit()

    return updated
