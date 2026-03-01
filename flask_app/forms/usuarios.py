from flask_wtf import FlaskForm
from wtforms import StringField, PasswordField, SelectField, BooleanField, SubmitField
from wtforms.validators import DataRequired, Email, EqualTo


class UserCreateForm(FlaskForm):
    username = StringField('CPF', validators=[DataRequired()])
    first_name = StringField('Nome', validators=[DataRequired()])
    last_name = StringField('Patente', validators=[DataRequired()])
    email = StringField('E-mail', validators=[DataRequired(), Email()])
    perfil = SelectField('Perfil', choices=[('user', 'Usuário'), ('admin', 'Admin')], validators=[DataRequired()])
    is_active = BooleanField('Ativo', default=True)
    password = PasswordField('Senha', validators=[DataRequired()])
    confirm_password = PasswordField('Confirmar Senha', validators=[DataRequired(), EqualTo('password', message='Senhas não conferem')])
    submit = SubmitField('Salvar')


class UserEditForm(FlaskForm):
    username = StringField('CPF', validators=[DataRequired()])
    first_name = StringField('Nome', validators=[DataRequired()])
    last_name = StringField('Patente', validators=[DataRequired()])
    email = StringField('E-mail', validators=[DataRequired(), Email()])
    perfil = SelectField('Perfil', choices=[('user', 'Usuário'), ('admin', 'Admin')], validators=[DataRequired()])
    is_active = BooleanField('Ativo', default=True)
    submit = SubmitField('Salvar')
