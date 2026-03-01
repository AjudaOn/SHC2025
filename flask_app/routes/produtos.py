from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_required

from models import db, Produto
from forms.produtos import ProdutoForm

bp = Blueprint('produtos', __name__)


@bp.route('/configuracao/produtos/')
@login_required
def consultar_produtos():
    produtos = Produto.query.order_by(Produto.nome.asc()).all()
    return render_template('portal/consultar_produto.html', produtos=produtos)


@bp.route('/configuracao/produtos/novo/', methods=['GET', 'POST'])
@login_required
def cadastrar_produto():
    form = ProdutoForm()

    if form.validate_on_submit():
        produto = Produto(
            nome=(form.nome.data or '').strip(),
            valor=form.valor.data,
        )
        db.session.add(produto)
        db.session.commit()
        flash('Produto cadastrado com sucesso.', 'success')
        return redirect(url_for('produtos.consultar_produtos'))

    return render_template('portal/cadastrar_produto.html', form=form, mode='create')


@bp.route('/configuracao/produtos/<int:produto_id>/editar/', methods=['GET', 'POST'])
@login_required
def editar_produto(produto_id):
    produto = Produto.query.get_or_404(produto_id)
    form = ProdutoForm()

    if request.method == 'GET':
        form.nome.data = produto.nome
        form.valor.data = produto.valor

    if form.validate_on_submit():
        produto.nome = (form.nome.data or '').strip()
        produto.valor = form.valor.data
        db.session.commit()
        flash('Produto atualizado com sucesso.', 'success')
        return redirect(url_for('produtos.consultar_produtos'))

    return render_template('portal/cadastrar_produto.html', form=form, mode='edit')
