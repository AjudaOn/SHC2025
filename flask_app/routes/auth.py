from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_user, logout_user, login_required, current_user
from models.user import User
from forms.auth import LoginForm, ChangePasswordForm

bp = Blueprint('auth', __name__)

@bp.route('/login/', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        if current_user.is_superuser or current_user.is_staff:
            return redirect(url_for('recepcao.consultar_reservas'))
        return redirect(url_for('recepcao.ocupacao_hoje_geral'))
    
    form = LoginForm()
    if form.validate_on_submit():
        user = User.query.filter_by(username=form.username.data).first()
        if user and user.check_password(form.password.data):
            if user.is_active:
                login_user(user)
                next_page = request.args.get('next')
                if next_page:
                    return redirect(next_page)
                if user.is_superuser or user.is_staff:
                    return redirect(url_for('recepcao.consultar_reservas'))
                return redirect(url_for('recepcao.ocupacao_hoje_geral'))
            else:
                flash('Sua conta está desativada.', 'danger')
        else:
            flash('Usuário ou senha inválidos.', 'danger')
            
    return render_template('auth/login.html', form=form)

@bp.route('/logout/')
@login_required
def logout():
    logout_user()
    return redirect(url_for('auth.login'))


@bp.route('/alterar-senha/', methods=['GET', 'POST'])
@login_required
def alterar_senha():
    form = ChangePasswordForm()
    if form.validate_on_submit():
        if not current_user.check_password(form.current_password.data):
            flash('Senha atual inválida.', 'danger')
        else:
            current_user.set_password(form.new_password.data)
            User.query.session.commit()
            flash('Senha atualizada com sucesso.', 'success')
            if current_user.is_superuser or current_user.is_staff:
                return redirect(url_for('recepcao.consultar_reservas'))
            return redirect(url_for('recepcao.ocupacao_hoje_geral'))
    return render_template('auth/alterar_senha.html', form=form)
