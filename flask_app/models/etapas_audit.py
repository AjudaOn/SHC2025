from datetime import datetime

from . import db


class ReservaEtapaAudit(db.Model):
    __tablename__ = "reserva_etapa_audit"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)

    reserva_id = db.Column(db.Integer, nullable=False, index=True)
    etapa = db.Column(db.String(32), nullable=False, index=True)
    acao = db.Column(db.String(16), nullable=False)

    user_id = db.Column(db.Integer, nullable=True, index=True)
    username = db.Column(db.String(150), nullable=True)

    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, index=True)
    details = db.Column(db.Text, nullable=True)

    __table_args__ = (
        db.Index(
            "ix_reserva_etapa_audit_reserva_created_at",
            "reserva_id",
            "created_at",
        ),
    )

