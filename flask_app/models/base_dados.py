from . import db
from sqlalchemy.orm import validates


class BaseDados(db.Model):
    __tablename__ = 'base_dados'

    id = db.Column('ID', db.Integer, primary_key=True)
    entrada = db.Column('ENTRADA', db.Date)
    saida = db.Column('SAÍDA', db.Date)
    nome = db.Column('NOME', db.String(100))
    diarias = db.Column('DIARIAS', db.Integer, default=0)
    graduacao = db.Column('GRADUAÇÃO', db.String(8))
    quartel = db.Column('QUARTEL', db.String(120))
    funcao = db.Column('FUNCAO', db.String(120))
    telefone = db.Column('TELEFONE', db.String(15))
    qtde_quartos = db.Column('QTDE_QUARTOS', db.Integer)
    qtde_hosp = db.Column('QTDE_HOSP', db.Integer)
    especial = db.Column('ESPECIAL', db.String(3))
    qtde_acomp = db.Column('QTDE_ACOMP', db.Integer)
    email = db.Column('EMAIL', db.String(100))
    cpf = db.Column('CPF', db.String(15))
    status = db.Column('STATUS', db.String(25))
    tipo = db.Column('TIPO', db.String(25))
    sexo = db.Column('SEXO', db.String(1))
    cidade = db.Column('CIDADE', db.String(100))
    uf = db.Column('UF', db.String(2))
    status_reserva = db.Column('STATUS_RESERVA', db.String(15))

    # Acompanhantes
    nome_acomp1 = db.Column('NOME_ACOMP1', db.String(100))
    vinculo_acomp1 = db.Column('VINCULO_ACOMP1', db.String(22))
    idade_acomp1 = db.Column('IDADE_ACOMP1', db.Integer)
    sexo_acomp1 = db.Column('SEXO_ACOMP1', db.String(1))

    nome_acomp2 = db.Column('NOME_ACOMP2', db.String(100))
    vinculo_acomp2 = db.Column('VINCULO_ACOMP2', db.String(22))
    idade_acomp2 = db.Column('IDADE_ACOMP2', db.Integer)
    sexo_acomp2 = db.Column('SEXO_ACOMP2', db.String(1))

    nome_acomp3 = db.Column('NOME_ACOMP3', db.String(100))
    vinculo_acomp3 = db.Column('VINCULO_ACOMP3', db.String(22))
    idade_acomp3 = db.Column('IDADE_ACOMP3', db.Integer)
    sexo_acomp3 = db.Column('SEXO_ACOMP3', db.String(1))

    nome_acomp4 = db.Column('NOME_ACOMP4', db.String(100))
    vinculo_acomp4 = db.Column('VINCULO_ACOMP4', db.String(22))
    idade_acomp4 = db.Column('IDADE_ACOMP4', db.Integer)
    sexo_acomp4 = db.Column('SEXO_ACOMP4', db.String(1))

    nome_acomp5 = db.Column('NOME_ACOMP5', db.String(100))
    vinculo_acomp5 = db.Column('VINCULO_ACOMP5', db.String(22))
    idade_acomp5 = db.Column('IDADE_ACOMP5', db.Integer)
    sexo_acomp5 = db.Column('SEXO_ACOMP5', db.String(1))

    mhex = db.Column('MHEx', db.String(6))
    uh = db.Column('UH', db.String(2))
    forma_pagamento = db.Column('FORMA_PAGAMENTO', db.String(2))

    # Valores financeiros
    valor_hosp = db.Column('VALOR_HOSP', db.Numeric(10, 2), default=0)
    valor_acomp1 = db.Column('VALOR_ACOMP1', db.Numeric(10, 2), default=0)
    valor_acomp2 = db.Column('VALOR_ACOMP2', db.Numeric(10, 2), default=0)
    valor_acomp3 = db.Column('VALOR_ACOMP3', db.Numeric(10, 2), default=0)
    valor_acomp4 = db.Column('VALOR_ACOMP4', db.Numeric(10, 2), default=0)
    valor_acomp5 = db.Column('VALOR_ACOMP5', db.Numeric(10, 2), default=0)

    valor_dia = db.Column('VALOR_DIA', db.Numeric(10, 2), default=0)
    valor_ajuste = db.Column('VALOR_AJUSTE', db.Numeric(10, 2), default=0, nullable=True)
    subtotal = db.Column('SUBTOTAL', db.Numeric(10, 2), default=0)
    valor_total = db.Column('VALOR_TOTAL', db.Numeric(10, 2), default=0)

    # Consumação
    qtde_agua = db.Column('QTDE_AGUA', db.Integer)
    qtde_refri = db.Column('QTDE_REFRI', db.Integer)
    qtde_cerveja = db.Column('QTDE_CERVEJA', db.Integer)
    qtde_pet = db.Column('QTDE_PET', db.Integer)

    total_agua = db.Column('TOTAL_AGUA', db.Numeric(10, 2), default=0)
    total_refri = db.Column('TOTAL_REFRI', db.Numeric(10, 2), default=0)
    total_cerveja = db.Column('TOTAL_CERVEJA', db.Numeric(10, 2), default=0)
    total_pet = db.Column('TOTAL_PET', db.Numeric(10, 2), default=0)
    total_consumacao = db.Column('TOTAL_CONSUMACAO', db.Numeric(10, 2), default=0)

    nome_pagante = db.Column('NOME_PAGANTE', db.String(100))
    cpf_pagante = db.Column('CPF_PAGANTE', db.String(100))

    motivo_viagem = db.Column('MOTIVO_VIAGEM', db.String(20))
    desc_saude = db.Column('DESC_SAUDE', db.Numeric(10, 2), default=0)
    observacao = db.Column('OBSERVACAO', db.String(100))

    def __repr__(self):
        return f'<BaseDados {self.id}: {self.nome}>'

    @staticmethod
    def _cpf_digits(value):
        if value is None:
            return None
        digits = ''.join(ch for ch in str(value) if ch.isdigit())
        return digits or None

    @validates('cpf', 'cpf_pagante')
    def _normalize_cpf(self, key, value):
        return self._cpf_digits(value)
