from flask import Blueprint, request, jsonify, render_template
from models import db, BaseDados
from forms.reservas import ReservasForm
from services.pricing import calculate_total
from datetime import datetime
from services.etapas_audit import log_reserva_etapa

bp = Blueprint('reservas', __name__)

@bp.route('/', methods=['GET', 'POST'])
def reserva_externa():
    form = ReservasForm()
    
    if request.method == 'POST':
        # Em Flask-WTF, usamos form.validate_on_submit() ou passamos o formdata
        # Como estamos usando AJAX com FormData, o Flask-WTF deve capturar automaticamente
        if form.validate_on_submit():
            try:
                # Criar novo objeto BaseDados a partir dos dados do form
                reserva = BaseDados()
                form.normalize_quartel_funcao()
                form.populate_obj(reserva)
                # MHEx fixo para reserva externa
                reserva.mhex = 'HTM_01'
                
                # Se o ID for string vazia, definir como None para o autoincrement do SQLite funcionar
                if not reserva.id:
                    reserva.id = None
                
                # Garantir que o status da reserva externa seja 'Pendente'
                reserva.status_reserva = 'Pendente'
                
                # Calcular totais antes de salvar
                calculate_total(reserva)
                
                db.session.add(reserva)
                db.session.commit()
                log_reserva_etapa(
                    reserva.id,
                    "RESERVA",
                    "CRIADA",
                    details={"origem": "externa", "status_reserva": reserva.status_reserva},
                    actor_name="Internet",
                )
                
                return jsonify({"success": True, "message": "Cadastro realizado com sucesso!", "id": reserva.id})
            except Exception as e:
                db.session.rollback()
                return jsonify({"success": False, "message": f"Erro ao salvar os dados: {str(e)}"}), 500
        else:
            return jsonify({"success": False, "message": "Dados do formulário inválidos!", "errors": form.errors}), 400
            
    return render_template('portal/index.html', form=form)
