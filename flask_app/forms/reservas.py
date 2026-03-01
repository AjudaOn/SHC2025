from flask_wtf import FlaskForm
from wtforms import StringField, DateField, IntegerField, SelectField, DecimalField, TextAreaField, HiddenField, BooleanField
from wtforms.validators import DataRequired, Email, Optional, ValidationError
from datetime import date
from constants import *
from models import BaseDados

class ReservasForm(FlaskForm):
    _GRADUACOES_COM_QUARTEL_FUNCAO = {
        GRADUACAO_GEN,
        GRADUACAO_CEL,
        GRADUACAO_TC,
        GRADUACAO_MAJ,
    }

    id = HiddenField('ID')
    
    # Datas
    entrada = DateField('Entrada', validators=[DataRequired()], render_kw={'type': 'date', 'class': 'form-control', 'id': 'entrada'})
    saida = DateField('SaÃ­da', validators=[DataRequired()], render_kw={'type': 'date', 'class': 'form-control', 'id': 'saida'})
    diarias = IntegerField('DiÃ¡rias', render_kw={'class': 'form-control', 'readonly': True})
    
    # Dados Pessoais
    nome = StringField('Nome', validators=[DataRequired()], render_kw={'class': 'form-control', 'id': 'nome'})
    email = StringField('Email', validators=[DataRequired(), Email()], render_kw={'class': 'form-control', 'id': 'email', 'type': 'email'})
    cpf = StringField('CPF', validators=[DataRequired()], render_kw={'class': 'form-control', 'id': 'cpf'})
    telefone = StringField('Telefone', validators=[DataRequired()], render_kw={'class': 'form-control', 'id': 'telefone'})
    sexo = SelectField('Sexo', choices=[('', '---')] + SEXO_CHOICES, validators=[DataRequired()], render_kw={'class': 'form-control', 'id': 'inputSexo'})
    
    # Status e Graduação
    status = SelectField('Status', choices=[('', '---')] + STATUS_CHOICES, validators=[DataRequired()], render_kw={'class': 'form-control', 'id': 'inputStatus'})
    graduacao = SelectField('Graduação', choices=[('', '---')] + GRADUACAO_CHOICES, validators=[DataRequired()], render_kw={'class': 'form-control', 'id': 'inputGraduacao'})
    quartel = StringField('Quartel', validators=[Optional()], render_kw={'class': 'form-control', 'id': 'inputQuartel'})
    funcao = StringField('Funcao', validators=[Optional()], render_kw={'class': 'form-control', 'id': 'inputFuncao'})
    tipo = SelectField('Tipo', choices=[('', '---')] + TIPO_CHOICES, validators=[DataRequired()], render_kw={'class': 'form-control', 'id': 'inputTipo'})
    
    # LocalizaÃ§Ã£o
    cidade = StringField('Cidade', validators=[DataRequired()], render_kw={'class': 'form-control', 'id': 'inputCidade'})
    uf = SelectField('UF', choices=[('', '---')] + UF_CHOICES, validators=[DataRequired()], render_kw={'class': 'form-control', 'id': 'inputUF'})
    
    # Acompanhantes
    # Acomp 1
    nome_acomp1 = StringField('Nome Acomp 1', render_kw={'class': 'form-control', 'id': 'inputNome_acomp1'})
    sexo_acomp1 = SelectField('Sexo Acomp 1', choices=[('', '---')] + SEXO_CHOICES, validators=[Optional()], render_kw={'class': 'form-control', 'id': 'inputSexo_acomp1'})
    vinculo_acomp1 = SelectField('VÃ­nculo Acomp 1', choices=[('', '---')] + VINCULO_CHOICES, validators=[Optional()], render_kw={'class': 'form-control', 'id': 'inputVinculo_acomp1'})
    idade_acomp1 = IntegerField('Idade Acomp 1', validators=[Optional()], render_kw={'class': 'form-control', 'id': 'inputIdade_acomp1'})
    
    # Acomp 2
    nome_acomp2 = StringField('Nome Acomp 2', render_kw={'class': 'form-control', 'id': 'inputNome_acomp2'})
    sexo_acomp2 = SelectField('Sexo Acomp 2', choices=[('', '---')] + SEXO_CHOICES, validators=[Optional()], render_kw={'class': 'form-control', 'id': 'inputSexo_acomp2'})
    vinculo_acomp2 = SelectField('VÃ­nculo Acomp 2', choices=[('', '---')] + VINCULO_CHOICES, validators=[Optional()], render_kw={'class': 'form-control', 'id': 'inputVinculo_acomp2'})
    idade_acomp2 = IntegerField('Idade Acomp 2', validators=[Optional()], render_kw={'class': 'form-control', 'id': 'inputIdade_acomp2'})
    
    # Acomp 3
    nome_acomp3 = StringField('Nome Acomp 3', render_kw={'class': 'form-control', 'id': 'inputNome_acomp3'})
    sexo_acomp3 = SelectField('Sexo Acomp 3', choices=[('', '---')] + SEXO_CHOICES, validators=[Optional()], render_kw={'class': 'form-control', 'id': 'inputSexo_acomp3'})
    vinculo_acomp3 = SelectField('VÃ­nculo Acomp 3', choices=[('', '---')] + VINCULO_CHOICES, validators=[Optional()], render_kw={'class': 'form-control', 'id': 'inputVinculo_acomp3'})
    idade_acomp3 = IntegerField('Idade Acomp 3', validators=[Optional()], render_kw={'class': 'form-control', 'id': 'inputIdade_acomp3'})
    
    # Acomp 4
    nome_acomp4 = StringField('Nome Acomp 4', render_kw={'class': 'form-control', 'id': 'inputNome_acomp4'})
    sexo_acomp4 = SelectField('Sexo Acomp 4', choices=[('', '---')] + SEXO_CHOICES, validators=[Optional()], render_kw={'class': 'form-control', 'id': 'inputSexo_acomp4'})
    vinculo_acomp4 = SelectField('VÃ­nculo Acomp 4', choices=[('', '---')] + VINCULO_CHOICES, validators=[Optional()], render_kw={'class': 'form-control', 'id': 'inputVinculo_acomp4'})
    idade_acomp4 = IntegerField('Idade Acomp 4', validators=[Optional()], render_kw={'class': 'form-control', 'id': 'inputIdade_acomp4'})
    
    # Acomp 5
    nome_acomp5 = StringField('Nome Acomp 5', render_kw={'class': 'form-control', 'id': 'inputNome_acomp5'})
    sexo_acomp5 = SelectField('Sexo Acomp 5', choices=[('', '---')] + SEXO_CHOICES, validators=[Optional()], render_kw={'class': 'form-control', 'id': 'inputSexo_acomp5'})
    vinculo_acomp5 = SelectField('VÃ­nculo Acomp 5', choices=[('', '---')] + VINCULO_CHOICES, validators=[Optional()], render_kw={'class': 'form-control', 'id': 'inputVinculo_acomp5'})
    idade_acomp5 = IntegerField('Idade Acomp 5', validators=[Optional()], render_kw={'class': 'form-control', 'id': 'inputIdade_acomp5'})
    
    # Reserva Info
    qtde_hosp = IntegerField('Qtd HÃ³spedes', validators=[DataRequired()], render_kw={'class': 'form-control', 'id': 'id_qtde_hosp'})
    qtde_quartos = IntegerField('Qtd Quartos', validators=[Optional()], render_kw={'class': 'form-control', 'id': 'inputQtde_quartos'})
    especial = SelectField('Especial', choices=[('', '---')] + ESPECIAL_CHOICES, validators=[DataRequired()], render_kw={'class': 'form-control', 'id': 'inputEspecial'})
    status_reserva = SelectField('Status Reserva', choices=[('', '---')] + STATUS_RESERVA_CHOICES, validators=[Optional()], render_kw={'class': 'form-control', 'id': 'id_status_reserva'})
    uh = SelectField('UH', choices=[('', '---')] + UH_CHOICES, coerce=str, validators=[Optional()], render_kw={'class': 'form-control', 'id': 'inputUH'})
    motivo_viagem = SelectField('Motivo Viagem', choices=[('', '---')] + MOTIVO_VIAGEM_CHOICES, validators=[DataRequired()], render_kw={'class': 'form-control', 'id': 'motivo_viagem'})
    observacao = TextAreaField('ObservaÃ§Ã£o', render_kw={'class': 'form-control', 'id': 'observacao', 'rows': 3})
    
    # Consumo e Financeiro
    qtde_agua = IntegerField('Qtd Ãgua', validators=[Optional()], render_kw={'class': 'form-control', 'id': 'qtde_agua'})
    qtde_refri = IntegerField('Qtd Refri', validators=[Optional()], render_kw={'class': 'form-control', 'id': 'qtde_refri'})
    qtde_cerveja = IntegerField('Qtd Cerveja', validators=[Optional()], render_kw={'class': 'form-control', 'id': 'qtde_cerveja'})
    qtde_pet = IntegerField('Qtd Pet', validators=[Optional()], render_kw={'class': 'form-control', 'id': 'qtde_pet'})
    
    total_agua = DecimalField('Total Ãgua', places=2, render_kw={'class': 'form-control', 'id': 'total_agua', 'readonly': True})
    total_refri = DecimalField('Total Refri', places=2, render_kw={'class': 'form-control', 'id': 'total_refri', 'readonly': True})
    total_cerveja = DecimalField('Total Cerveja', places=2, render_kw={'class': 'form-control', 'id': 'total_cerveja', 'readonly': True})
    total_pet = DecimalField('Total Pet', places=2, render_kw={'class': 'form-control', 'id': 'total_pet', 'readonly': True})
    total_consumacao = DecimalField('Total ConsumaÃ§Ã£o', places=2, render_kw={'class': 'form-control', 'id': 'total_consumacao', 'readonly': True})
    
    valor_ajuste = DecimalField('Valor Ajuste', places=2, validators=[Optional()], render_kw={'class': 'form-control', 'id': 'valor_ajuste'})
    desc_saude = DecimalField('Desc SaÃºde', places=2, render_kw={'class': 'form-control', 'id': 'desc_saude', 'readonly': True})
    
    forma_pagamento = StringField('Forma Pagamento', render_kw={'class': 'form-control', 'id': 'forma_pagamento'})
    nome_pagante = StringField('Nome Pagante', render_kw={'class': 'form-control', 'id': 'nome_pagante'})
    cpf_pagante = StringField('CPF Pagante', render_kw={'class': 'form-control', 'id': 'cpf_pagante'})

    def _exige_quartel_funcao(self):
        return self.graduacao.data in self._GRADUACOES_COM_QUARTEL_FUNCAO

    def validate_saida(self, field):
        if self.entrada.data and field.data and self.entrada.data > field.data:
            raise ValidationError('A data de saÃ­da deve ser posterior ou igual Ã  data de entrada.')

    def validate(self, extra_validators=None):
        is_valid = super().validate(extra_validators=extra_validators)

        if self._exige_quartel_funcao():
            quartel = (self.quartel.data or '').strip()
            funcao = (self.funcao.data or '').strip()

            if not quartel:
                self.quartel.errors.append('Campo obrigatÃ³rio para esta graduaÃ§Ã£o.')
                is_valid = False
            if not funcao:
                self.funcao.errors.append('Campo obrigatÃ³rio para esta graduaÃ§Ã£o.')
                is_valid = False

        return is_valid

    def validate_uh(self, field):
        if field.data and self.entrada.data and self.saida.data:
            from services.availability import get_available_uhs

            exclude_id = self.id.data if self.id.data else None
            available_uhs = get_available_uhs(MHEX_HTM_01, self.entrada.data, self.saida.data, exclude_id)
            available_codes = [code for code, label in available_uhs]

            if exclude_id:
                reserva_atual = BaseDados.query.get(exclude_id)
                if reserva_atual and str(reserva_atual.uh).strip() == str(field.data).strip():
                    return

            if field.data not in available_codes:
                raise ValidationError(f'O quarto {field.data} nÃ£o estÃ¡ disponÃ­vel para o perÃ­odo e hotel selecionados.')

    def normalize_quartel_funcao(self):
        if self._exige_quartel_funcao():
            return
        self.quartel.data = None
        self.funcao.data = None

class AdminReservasForm(ReservasForm):
    mhex = SelectField('MHEx', choices=MHEx_CHOICES, validators=[DataRequired()], render_kw={'class': 'form-control', 'id': 'inputMhex'})
    pagante_checkbox = BooleanField('A conta nÃ£o serÃ¡ paga pelo HÃ³spede Principal.')

    def validate_uh(self, field):
        if self.mhex.data and field.data and self.entrada.data and self.saida.data:
            from services.availability import get_available_uhs

            exclude_id = int(self.id.data) if self.id.data else None
            available_uhs = get_available_uhs(self.mhex.data, self.entrada.data, self.saida.data, exclude_id)
            available_codes = [code for code, label in available_uhs]

            if exclude_id:
                reserva_atual = BaseDados.query.get(exclude_id)
                if reserva_atual and str(reserva_atual.uh).strip() == str(field.data).strip():
                    return

            if field.data not in available_codes:
                raise ValidationError(f'O quarto {field.data} nÃ£o estÃ¡ disponÃ­vel para o perÃ­odo e hotel selecionados.')


class ReservaEdicaoRapidaForm(FlaskForm):
    """
    FormulÃ¡rio mÃ­nimo para editar uma reserva jÃ¡ existente (tela "EDITAR RESERVA" simples).

    Motivo:
    - `AdminReservasForm` herda muitos campos obrigatÃ³rios do cadastro e exige que o template
      renderize todos eles. Na tela simples, exibimos vÃ¡rios campos como somente leitura
      (fora do WTForms), o que tornava a validaÃ§Ã£o impossÃ­vel e o "Salvar" nÃ£o persistia.
    """

    id = HiddenField('ID')
    entrada = DateField('Entrada', validators=[DataRequired()], render_kw={'type': 'date', 'id': 'entrada'})
    saida = DateField('SaÃ­da', validators=[DataRequired()], render_kw={'type': 'date', 'id': 'saida'})
    status_reserva = SelectField(
        'Status Reserva',
        choices=[('', '---')] + STATUS_RESERVA_CHOICES,
        validators=[Optional()],
    )
    mhex = SelectField('MHEx', choices=MHEx_CHOICES, validators=[DataRequired()], render_kw={'id': 'mhex'})
    uh = SelectField('UH', choices=[('', '---')] + UH_CHOICES, coerce=str, validators=[Optional()], render_kw={'id': 'uh'})

    def validate_saida(self, field):
        if self.entrada.data and field.data and self.entrada.data > field.data:
            raise ValidationError('A data de saÃ­da deve ser posterior ou igual Ã  data de entrada.')

    def validate_uh(self, field):
        if self.mhex.data and field.data and self.entrada.data and self.saida.data:
            from services.availability import get_available_uhs

            exclude_id = int(self.id.data) if self.id.data else None
            available_uhs = get_available_uhs(self.mhex.data, self.entrada.data, self.saida.data, exclude_id)
            available_codes = [code for code, label in available_uhs]

            if exclude_id:
                reserva_atual = BaseDados.query.get(exclude_id)
                if reserva_atual and str(reserva_atual.uh).strip() == str(field.data).strip():
                    return

            if field.data not in available_codes:
                raise ValidationError(f'O quarto {field.data} nÃ£o estÃ¡ disponÃ­vel para o perÃ­odo e hotel selecionados.')



