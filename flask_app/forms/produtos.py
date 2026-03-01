from flask_wtf import FlaskForm
from wtforms import StringField, DecimalField, SubmitField
from wtforms.validators import DataRequired


class ProdutoForm(FlaskForm):
    nome = StringField('Nome', validators=[DataRequired()])
    valor = DecimalField('Valor', places=2, validators=[DataRequired()])
    submit = SubmitField('Salvar')
