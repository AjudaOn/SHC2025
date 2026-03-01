from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_required
from passlib.hash import django_pbkdf2_sha256
from datetime import datetime
import re

from models import db
from models.user import User
from forms.usuarios import UserCreateForm, UserEditForm

bp = Blueprint('usuarios', __name__)


@bp.route('/configuracao/usuarios/')
@login_required
def consultar_usuarios():
    usuarios = User.query.order_by(User.first_name.asc(), User.username.asc()).all()
    return render_template('portal/consultar_usuario.html', usuarios=usuarios)


@bp.route('/configuracao/usuarios/novo/', methods=['GET', 'POST'])
@login_required
def cadastrar_usuario():
    form = UserCreateForm()

    if form.validate_on_submit():
        cpf_digits = re.sub(r'\D', '', form.username.data or '')
        if not cpf_digits:
            flash('CPF inválido.', 'error')
            return render_template('portal/cadastrar_usuario.html', form=form)

        existing = User.query.filter_by(username=cpf_digits).first()
        if existing:
            flash('Já existe um usuário com este CPF.', 'error')
            return render_template('portal/cadastrar_usuario.html', form=form)

        is_admin = form.perfil.data == 'admin'

        user = User(
            username=cpf_digits,
            first_name=(form.first_name.data or '').strip(),
            last_name=(form.last_name.data or '').strip(),
            email=(form.email.data or '').strip(),
            is_staff=is_admin,
            is_superuser=is_admin,
            is_active=bool(form.is_active.data),
            date_joined=datetime.utcnow(),
            password=django_pbkdf2_sha256.hash(form.password.data),
        )
        db.session.add(user)
        db.session.commit()
        flash('Usuário cadastrado com sucesso.', 'success')
        return redirect(url_for('usuarios.consultar_usuarios'))

    return render_template('portal/cadastrar_usuario.html', form=form, mode='create')


@bp.route('/configuracao/usuarios/<int:user_id>/editar/', methods=['GET', 'POST'])
@login_required
def editar_usuario(user_id):
    user = User.query.get_or_404(user_id)
    form = UserEditForm()

    if request.method == 'GET':
        form.username.data = user.username
        form.first_name.data = user.first_name
        form.last_name.data = user.last_name
        form.email.data = user.email
        form.perfil.data = 'admin' if user.is_superuser or user.is_staff else 'user'
        form.is_active.data = bool(user.is_active)

    if form.validate_on_submit():
        cpf_digits = re.sub(r'\\D', '', form.username.data or '')
        if not cpf_digits:
            flash('CPF inválido.', 'error')
            return render_template('portal/cadastrar_usuario.html', form=form, mode='edit')

        existing = User.query.filter(User.username == cpf_digits, User.id != user.id).first()
        if existing:
            flash('Já existe um usuário com este CPF.', 'error')
            return render_template('portal/cadastrar_usuario.html', form=form, mode='edit')

        is_admin = form.perfil.data == 'admin'
        user.username = cpf_digits
        user.first_name = (form.first_name.data or '').strip()
        user.last_name = (form.last_name.data or '').strip()
        user.email = (form.email.data or '').strip()
        user.is_staff = is_admin
        user.is_superuser = is_admin
        user.is_active = bool(form.is_active.data)

        db.session.commit()
        flash('Usuário atualizado com sucesso.', 'success')
        return redirect(url_for('usuarios.consultar_usuarios'))

    return render_template('portal/cadastrar_usuario.html', form=form, mode='edit')


@bp.route('/configuracao/usuarios/<int:user_id>/resetar-senha/')
@login_required
def resetar_senha_usuario(user_id):
    user = User.query.get_or_404(user_id)
    cpf_digits = re.sub(r'\D', '', user.username or '')
    if not cpf_digits:
        flash('Não foi possível resetar: CPF inválido.', 'error')
        return redirect(url_for('usuarios.consultar_usuarios'))

    user.password = django_pbkdf2_sha256.hash(cpf_digits)
    db.session.commit()
    nome = (user.first_name or '').strip() or 'Usuário'
    flash(f'Senha do usuário {nome} foi resetada. Favor usar o CPF (somente números) como senha.', 'reset')
    return redirect(url_for('usuarios.consultar_usuarios'))
