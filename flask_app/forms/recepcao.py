from flask_wtf import FlaskForm
from wtforms import StringField, SelectField, IntegerField, DecimalField, BooleanField, SubmitField, DateField, RadioField
from wtforms.validators import DataRequired, Optional
from constants import (
    STATUS_CHOICES,
    GRADUACAO_CHOICES,
    TIPO_CHOICES,
    UF_CHOICES,
    ESPECIAL_CHOICES,
    SEXO_CHOICES,
    VINCULO_CHOICES,
    MHEx_CHOICES,
    UH_CHOICES,
)

class CheckinForm(FlaskForm):
    entrada = DateField('Entrada', validators=[DataRequired()])
    saida = DateField('Saída', validators=[DataRequired()])
    diarias = IntegerField('Diárias', validators=[Optional()])
    qtde_hosp = IntegerField('Qtde de Hóspedes', validators=[DataRequired()])
    qtde_quartos = IntegerField('Qtde de Quartos', validators=[Optional()])
    especial = SelectField('Necessidades Especiais', choices=ESPECIAL_CHOICES, validators=[Optional()])

    nome = StringField('Nome', validators=[DataRequired()])
    cpf = StringField('CPF', validators=[DataRequired()])
    telefone = StringField('Telefone', validators=[Optional()])
    graduacao = SelectField('Graduação', choices=GRADUACAO_CHOICES, validators=[Optional()])
    status = SelectField('Status', choices=STATUS_CHOICES, validators=[Optional()])
    tipo = SelectField('Tipo de Hospedagem', choices=TIPO_CHOICES, validators=[Optional()])

    cidade = StringField('Cidade', validators=[Optional()])
    uf = SelectField('UF', choices=UF_CHOICES, validators=[Optional()])
    mhex = SelectField('MHEx', choices=MHEx_CHOICES, validators=[Optional()])
    uh = SelectField('UH', choices=[('', '---')] + UH_CHOICES, coerce=str, validators=[Optional()])

    # Acompanhantes
    nome_acomp1 = StringField('Nome Acomp 1', validators=[Optional()])
    sexo_acomp1 = SelectField('Sexo Acomp 1', choices=[('', '---')] + SEXO_CHOICES, validators=[Optional()])
    vinculo_acomp1 = SelectField('Vínculo Acomp 1', choices=[('', '---')] + VINCULO_CHOICES, validators=[Optional()])
    idade_acomp1 = IntegerField('Idade Acomp 1', validators=[Optional()])

    nome_acomp2 = StringField('Nome Acomp 2', validators=[Optional()])
    sexo_acomp2 = SelectField('Sexo Acomp 2', choices=[('', '---')] + SEXO_CHOICES, validators=[Optional()])
    vinculo_acomp2 = SelectField('Vínculo Acomp 2', choices=[('', '---')] + VINCULO_CHOICES, validators=[Optional()])
    idade_acomp2 = IntegerField('Idade Acomp 2', validators=[Optional()])

    nome_acomp3 = StringField('Nome Acomp 3', validators=[Optional()])
    sexo_acomp3 = SelectField('Sexo Acomp 3', choices=[('', '---')] + SEXO_CHOICES, validators=[Optional()])
    vinculo_acomp3 = SelectField('Vínculo Acomp 3', choices=[('', '---')] + VINCULO_CHOICES, validators=[Optional()])
    idade_acomp3 = IntegerField('Idade Acomp 3', validators=[Optional()])

    nome_acomp4 = StringField('Nome Acomp 4', validators=[Optional()])
    sexo_acomp4 = SelectField('Sexo Acomp 4', choices=[('', '---')] + SEXO_CHOICES, validators=[Optional()])
    vinculo_acomp4 = SelectField('Vínculo Acomp 4', choices=[('', '---')] + VINCULO_CHOICES, validators=[Optional()])
    idade_acomp4 = IntegerField('Idade Acomp 4', validators=[Optional()])

    nome_acomp5 = StringField('Nome Acomp 5', validators=[Optional()])
    sexo_acomp5 = SelectField('Sexo Acomp 5', choices=[('', '---')] + SEXO_CHOICES, validators=[Optional()])
    vinculo_acomp5 = SelectField('Vínculo Acomp 5', choices=[('', '---')] + VINCULO_CHOICES, validators=[Optional()])
    idade_acomp5 = IntegerField('Idade Acomp 5', validators=[Optional()])

    submit = SubmitField('Salvar')

class CheckoutForm(FlaskForm):
    status_reserva = SelectField('Status da Reserva', choices=[
        ('Checkin', 'Check-in'),
        ('Pago', 'Pago')
    ], validators=[DataRequired()])
    forma_pagamento = SelectField('Forma de Pagamento', choices=[
        ('', 'Selecione'),
        ('DINHEIRO', 'Dinheiro'),
        ('PIX', 'PIX')
    ], validators=[DataRequired()])
    qtde_agua = IntegerField('Qtd Água', validators=[Optional()])
    qtde_refri = IntegerField('Qtd Refri', validators=[Optional()])
    qtde_cerveja = IntegerField('Qtd Cerveja', validators=[Optional()])
    qtde_pet = IntegerField('Qtd Pet', validators=[Optional()])
    pagante_tipo = RadioField(
        'Pagante',
        choices=[('principal', 'Hóspede principal'), ('outro', 'Outro pagante')],
        validators=[DataRequired()]
    )
    nome_pagante = StringField('Nome completo', validators=[Optional()])
    cpf_pagante = StringField('CPF', validators=[Optional()])
    submit = SubmitField('Salvar')

class ConsumacaoForm(FlaskForm):
    qtde_agua = IntegerField('Água', default=0)
    qtde_refri = IntegerField('Refrigerante', default=0)
    qtde_cerveja = IntegerField('Cerveja', default=0)
    qtde_pet = IntegerField('Pets', default=0)
    submit = SubmitField('Salvar')


class SaidaForm(FlaskForm):
    saida = DateField('Saída', validators=[DataRequired()], render_kw={'type': 'date', 'id': 'saida'})
    submit = SubmitField('Salvar')
