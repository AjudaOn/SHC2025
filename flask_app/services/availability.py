from models import BaseDados
from constants import (
    get_uh_choices,
    STATUS_RESERVA_APROVADA,
    STATUS_RESERVA_CHECKIN,
    STATUS_RESERVA_PAGO,
)

def get_available_uhs(mhex, entrada, saida, exclude_id=None):
    """
    Returns a list of available UHs (tuples of (value, label)) for the given criteria.
    Excludes reservations with 'Recusada' status.
    """
    if not mhex or not entrada or not saida:
        return get_uh_choices(mhex)

    # Find overlapping reservations
    # Overlap logic: (StartA < EndB) and (EndA > StartB)
    # Only active statuses should block availability.
    query = BaseDados.query.filter(
        BaseDados.mhex == mhex,
        BaseDados.entrada < saida,
        BaseDados.saida > entrada,
        BaseDados.status_reserva.in_([
            STATUS_RESERVA_APROVADA,
            STATUS_RESERVA_CHECKIN,
            STATUS_RESERVA_PAGO,
        ])
    )

    if exclude_id:
        query = query.filter(BaseDados.id != exclude_id)

    # Normaliza para string porque o SQLite pode devolver números em colunas textuais
    occupied_uhs = {str(r.uh) for r in query.all() if r.uh is not None and str(r.uh).strip() != ''}

    # Filter UHs do hotel selecionado
    uh_choices = get_uh_choices(mhex)
    available = [
        (uh_code, uh_label) 
        for uh_code, uh_label in uh_choices 
        if uh_code not in occupied_uhs
    ]

    return available
