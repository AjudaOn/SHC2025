from flask import Blueprint, render_template, request, redirect, url_for, jsonify, flash, abort, current_app
from flask_login import login_required
from flask_login import current_user
from models import db, BaseDados
from forms.recepcao import CheckinForm, CheckoutForm, ConsumacaoForm
from services.pricing import calculate_total
from services.reservas import expire_reservas
from datetime import datetime, date, timedelta, time
import calendar
from sqlalchemy import and_, func, or_
from constants import (
    MHEX_HTM_01,
    MHEX_HTM_02,
    STATUS_RESERVA_APROVADA,
    STATUS_RESERVA_CHECKIN,
    STATUS_RESERVA_EXPIRADA,
    STATUS_RESERVA_PAGO,
    get_uh_choices,
)
from models.etapas_audit import ReservaEtapaAudit
from services.etapas_audit import log_reserva_etapa

bp = Blueprint('recepcao', __name__)

_HOTEL_TURNOVER_TIME = time(12, 0)  # check-out até 12:00; nova entrada a partir de 12:00


def _now_local():
    tz_name = None
    try:
        tz_name = (current_app.config.get("TIMEZONE") or "").strip() or None
    except Exception:
        tz_name = None

    if not tz_name:
        return datetime.now()

    try:
        from zoneinfo import ZoneInfo  # py3.9+

        return datetime.now(tz=ZoneInfo(tz_name))
    except Exception:
        return datetime.now()


def _safe_next_url(value):
    value = (value or "").strip()
    if not value:
        return None
    if value.startswith("/"):
        return value
    return None


def _require_admin():
    if not current_user.is_authenticated:
        abort(401)
    if not (current_user.is_superuser or current_user.is_staff):
        abort(403)


def _normalize_uh(value):
    if value is None:
        return None
    text = str(value).strip()
    if text == "":
        return None
    if text.isdigit():
        return str(int(text))
    return text


def _diff_changed_fields(before: dict, after_obj, fields):
    changed = []
    for field in fields:
        before_value = before.get(field)
        after_value = getattr(after_obj, field, None)
        if field == "uh":
            before_value = _normalize_uh(before_value)
            after_value = _normalize_uh(after_value)
        elif field in {"cpf", "cpf_pagante", "telefone"}:
            def _digits(value):
                if value is None:
                    return None
                digits = "".join([c for c in str(value) if c.isdigit()])
                return digits or None
            before_value = _digits(before_value)
            after_value = _digits(after_value)
        else:
            if isinstance(before_value, str) and before_value.strip() == "":
                before_value = None
            if isinstance(after_value, str) and after_value.strip() == "":
                after_value = None
            if isinstance(before_value, str):
                before_value = before_value.strip()
            if isinstance(after_value, str):
                after_value = after_value.strip()
        if before_value != after_value:
            changed.append(field)
    return changed

def _build_ocupacao_hoje(mhex):
    now = _now_local()
    today = now.date()
    tzinfo = now.tzinfo

    reservas_candidates = BaseDados.query.filter(
        BaseDados.mhex == mhex,
        or_(
            # Mantém reservas em check-in visíveis até o checkout real (mudança de status).
            and_(
                BaseDados.status_reserva == STATUS_RESERVA_CHECKIN,
                BaseDados.entrada <= today,
                BaseDados.saida >= today,
            ),
            and_(
                BaseDados.status_reserva == STATUS_RESERVA_APROVADA,
                BaseDados.entrada <= today,
                BaseDados.saida >= today,
            ),
        ),
    ).all()

    room_keys = [choice[0] for choice in get_uh_choices(mhex)]
    rooms_map = {
        key: {
            "number": str(key).zfill(2),
            "status": "available",
        }
        for key in room_keys
    }

    reservas_by_uh = {}
    for reserva in reservas_candidates:
        room_key = _normalize_uh(reserva.uh)
        if not room_key or room_key not in rooms_map:
            continue
        reservas_by_uh.setdefault(room_key, []).append(reserva)

    def _pick_reserva_for_room(reservas):
        best = None
        best_rank = None

        for reserva in reservas:
            if not reserva.entrada or not reserva.saida:
                continue

            if reserva.status_reserva == STATUS_RESERVA_CHECKIN:
                is_active_now = True
                is_upcoming_today = False
            else:
                start_dt = datetime.combine(reserva.entrada, _HOTEL_TURNOVER_TIME, tzinfo=tzinfo)
                end_dt = datetime.combine(reserva.saida, _HOTEL_TURNOVER_TIME, tzinfo=tzinfo)

                is_active_now = start_dt <= now < end_dt
                is_upcoming_today = (reserva.entrada == today) and (now < start_dt)

            if not is_active_now and not is_upcoming_today:
                continue

            status_boost = 2 if reserva.status_reserva == STATUS_RESERVA_CHECKIN else 1
            state_boost = 2 if is_active_now else 1
            rank = (state_boost, status_boost, (reserva.id or 0))

            if best_rank is None or rank > best_rank:
                best = reserva
                best_rank = rank

        return best

    for room_key in room_keys:
        reserva = _pick_reserva_for_room(reservas_by_uh.get(room_key, []))
        if not reserva:
            continue

        room_entry = rooms_map[room_key]
        status = "occupied" if reserva.status_reserva == STATUS_RESERVA_CHECKIN else "reserved"

        if status == "occupied" or room_entry["status"] == "available":
            room_entry.update(
                {
                    "status": status,
                    "guest": reserva.nome,
                    "graduacao": reserva.graduacao,
                    "guests": reserva.qtde_hosp,
                    "check_in": reserva.entrada.strftime("%d/%m"),
                    "check_out": reserva.saida.strftime("%d/%m"),
                    "reserva_id": reserva.id,
                }
            )

    rooms = [rooms_map[key] for key in room_keys]
    return rooms


def _build_ocupacao_semanal(mhex):
    today = date.today()
    start_date = today - timedelta(days=(today.weekday() + 1) % 7)
    end_date = start_date + timedelta(days=6)

    reservas_semana = BaseDados.query.filter(
        BaseDados.mhex == mhex,
        BaseDados.status_reserva.in_([STATUS_RESERVA_APROVADA, STATUS_RESERVA_CHECKIN, STATUS_RESERVA_PAGO]),
        BaseDados.entrada <= end_date,
        BaseDados.saida > start_date,
    ).order_by(BaseDados.id.asc()).all()

    room_keys = [choice[0] for choice in get_uh_choices(mhex)]
    rooms_data = [
        {
            "room_number": str(key).zfill(2),
            "statuses": ["available"] * 7,
            "ids": [None] * 7,
        }
        for key in room_keys
    ]
    room_index = {room["room_number"]: room for room in rooms_data}

    priority = {
        "available": 0,
        "paid": 1,
        "reserved": 2,
        "occupied": 3,
    }

    for reserva in reservas_semana:
        room_key = _normalize_uh(reserva.uh)
        if not room_key:
            continue
        room_number = str(room_key).zfill(2)
        room_entry = room_index.get(room_number)
        if not room_entry:
            continue

        if reserva.status_reserva == STATUS_RESERVA_CHECKIN:
            status_value = "occupied"
        elif reserva.status_reserva == STATUS_RESERVA_PAGO:
            status_value = "paid"
        else:
            status_value = "reserved"
        current_day = start_date
        for idx in range(7):
            if reserva.entrada <= current_day < reserva.saida:
                existing = room_entry["statuses"][idx]
                existing_priority = priority.get(existing, 0)
                incoming_priority = priority.get(status_value, 0)
                existing_id = room_entry["ids"][idx] or 0
                incoming_id = reserva.id or 0

                should_replace = (
                    incoming_priority > existing_priority
                    or (incoming_priority == existing_priority and incoming_id > existing_id)
                )
                if should_replace:
                    room_entry["statuses"][idx] = status_value
                    room_entry["ids"][idx] = reserva.id
            current_day += timedelta(days=1)

    days = ["Domingo", "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"]
    day_dates = [(start_date + timedelta(days=i)).strftime("%d/%m") for i in range(7)]
    rooms = [str(key).zfill(2) for key in room_keys]

    return rooms, rooms_data, days, day_dates, start_date, end_date


def _build_ocupacao_mensal(mhex, month, year):
    last_day = calendar.monthrange(year, month)[1]
    start_date = date(year, month, 1)
    end_date = date(year, month, last_day)

    reservas_mes = BaseDados.query.filter(
        BaseDados.mhex == mhex,
        BaseDados.status_reserva.in_([STATUS_RESERVA_APROVADA, STATUS_RESERVA_CHECKIN, STATUS_RESERVA_PAGO]),
        BaseDados.entrada <= end_date,
        BaseDados.saida > start_date,
    ).all()

    room_keys = [choice[0] for choice in get_uh_choices(mhex)]
    rooms_data = [
        {
            "room_number": str(key).zfill(2),
            "statuses": ["available"] * last_day,
            "ids": [None] * last_day,
        }
        for key in room_keys
    ]
    room_index = {room["room_number"]: room for room in rooms_data}

    for reserva in reservas_mes:
        room_key = _normalize_uh(reserva.uh)
        if not room_key:
            continue
        room_number = str(room_key).zfill(2)
        room_entry = room_index.get(room_number)
        if not room_entry:
            continue

        if reserva.status_reserva == STATUS_RESERVA_CHECKIN:
            status_value = "occupied"
        elif reserva.status_reserva == STATUS_RESERVA_PAGO:
            status_value = "paid"
        else:
            status_value = "reserved"
        current_day = start_date
        for idx in range(last_day):
            if reserva.entrada <= current_day < reserva.saida:
                existing = room_entry["statuses"][idx]
                if existing != "occupied" or status_value == "occupied":
                    room_entry["statuses"][idx] = status_value
                    room_entry["ids"][idx] = reserva.id
            current_day += timedelta(days=1)

    days = [str(i) for i in range(1, last_day + 1)]
    rooms = [str(key).zfill(2) for key in room_keys]
    month_label = start_date.strftime("%m/%Y")

    return rooms, rooms_data, days, month_label, last_day

@bp.route('/reservas/etapas/')
@login_required
def consultar_etapas():
    _require_admin()
    reserva_id = request.args.get('id', type=int)

    reserva = None
    logs = []
    users_by_id = {}
    if reserva_id:
        reserva = BaseDados.query.get(reserva_id)
        if reserva:
            logs = (
                ReservaEtapaAudit.query.filter_by(reserva_id=reserva_id)
                .order_by(ReservaEtapaAudit.created_at.asc(), ReservaEtapaAudit.id.asc())
                .all()
            )
            user_ids = {log.user_id for log in logs if log.user_id}
            if user_ids:
                from models.user import User
                users = User.query.filter(User.id.in_(list(user_ids))).all()
                users_by_id = {
                    user.id: (user.first_name or "").strip()
                    for user in users
                }

    return render_template(
        'portal/consultar_etapas.html',
        reserva_id=reserva_id or '',
        reserva=reserva,
        logs=logs,
        users_by_id=users_by_id,
    )

@bp.route('/api/available-uhs/')
@login_required
def api_available_uhs():
    mhex = request.args.get('mhex')
    entrada_str = request.args.get('entrada')
    saida_str = request.args.get('saida')
    exclude_id = request.args.get('exclude_id', type=int)

    if not mhex or not entrada_str or not saida_str:
        return jsonify({'available_uhs': []})

    try:
        entrada = datetime.strptime(entrada_str, '%Y-%m-%d').date()
        saida = datetime.strptime(saida_str, '%Y-%m-%d').date()
    except ValueError:
        return jsonify({'available_uhs': []})

    from services.availability import get_available_uhs
    available_uhs = get_available_uhs(mhex, entrada, saida, exclude_id=exclude_id)
    
    return jsonify({'available_uhs': available_uhs})


@bp.route('/api/hospede')
@login_required
def api_hospede():
    cpf = (request.args.get('cpf') or '').strip()
    cpf_digits = ''.join(ch for ch in cpf if ch.isdigit())
    if len(cpf_digits) != 11:
        return jsonify({'found': False})

    def _clean(field):
        return func.replace(func.replace(func.replace(field, '.', ''), '-', ''), ' ', '')

    hospede = BaseDados.query.filter(
        or_(
            _clean(BaseDados.cpf) == cpf_digits,
            _clean(BaseDados.cpf_pagante) == cpf_digits
        )
    ).order_by(BaseDados.entrada.desc(), BaseDados.id.desc()).first()

    if not hospede:
        return jsonify({'found': False})

    return jsonify({
        'found': True,
        'data': {
            'nome': hospede.nome or '',
            'email': hospede.email or '',
            'telefone': hospede.telefone or '',
            'graduacao': hospede.graduacao or '',
            'quartel': hospede.quartel or '',
            'funcao': hospede.funcao or '',
            'status': hospede.status or '',
            'tipo': hospede.tipo or '',
            'sexo': hospede.sexo or '',
            'cidade': hospede.cidade or '',
            'uf': hospede.uf or '',
            'cpf': hospede.cpf or hospede.cpf_pagante or ''
        }
    })

@bp.route('/reservas/consultar/')
@login_required
def consultar_reservas():
    _require_admin()
    consultar = BaseDados.query.filter_by(status_reserva="Pendente").order_by(BaseDados.entrada.asc()).all()
    return render_template('portal/reservas.html', consultar=consultar)


@bp.route('/reservas/consultar-expirada/')
@login_required
def consultar_expiradas():
    _require_admin()
    hoje = _now_local().date()
    consultar = (
        BaseDados.query
        .filter(
            BaseDados.status_reserva == STATUS_RESERVA_EXPIRADA,
            func.strftime('%m', BaseDados.entrada) == f'{hoje.month:02d}',
            func.strftime('%Y', BaseDados.entrada) == str(hoje.year),
        )
        .order_by(BaseDados.id.desc())
        .all()
    )

    def _format_cpf_display(value):
        digits = ''.join(ch for ch in str(value or '') if ch.isdigit())
        if len(digits) == 11:
            return f'{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}'
        return (value or '')

    return render_template(
        'portal/consultar_expirada.html',
        consultar=consultar,
        format_cpf_display=_format_cpf_display,
    )


@bp.route('/reservas/expirar/', methods=['POST'])
@login_required
def expirar_reservas_manual():
    _require_admin()
    updated = expire_reservas(source="manual")
    flash(f'Reservas expiradas: {updated}', 'success')
    return redirect(url_for('recepcao.consultar_reservas'))

@bp.route('/ocupacao/editar/<int:id>/', methods=['GET', 'POST'])
@login_required
def editar_saida_pet(id):
    reserva = BaseDados.query.get_or_404(id)
    from forms.recepcao import SaidaForm
    from services.availability import get_available_uhs

    next_url = _safe_next_url(request.args.get('next'))
    form = SaidaForm()

    if request.method == 'GET':
        form.saida.data = reserva.saida

    if form.validate_on_submit():
        if not reserva.entrada:
            flash('Reserva sem data de entrada.', 'error')
        elif form.saida.data and form.saida.data <= reserva.entrada:
            flash('A data de saída deve ser posterior à data de entrada.', 'error')
        else:
            current_uh = (str(reserva.uh).strip() if reserva.uh is not None else "")
            if reserva.mhex and reserva.entrada and form.saida.data and current_uh:
                available_uhs = get_available_uhs(reserva.mhex, reserva.entrada, form.saida.data, exclude_id=reserva.id)
                available_codes = {code for code, _ in available_uhs}
                if current_uh not in available_codes:
                    flash('Não é possível estender/alterar a saída: UH fica indisponível no período.', 'error')
                    return render_template(
                        'portal/editar_saida_pet.html',
                        reserva=reserva,
                        form=form,
                        next_url=next_url,
                    )

            reserva.saida = form.saida.data
            reserva.diarias = max((reserva.saida - reserva.entrada).days, 0) if reserva.saida and reserva.entrada else 0
            calculate_total(reserva)

            db.session.commit()
            log_reserva_etapa(reserva.id, "RESERVA", "ALTERADA", details={"campo": "saida"})
            flash('Reserva atualizada com sucesso.', 'success')
            return redirect(next_url or (url_for('recepcao.ocupacao_hoje_htm02') if reserva.mhex == MHEX_HTM_02 else url_for('recepcao.ocupacao_hoje')))

    return render_template(
        'portal/editar_saida_pet.html',
        reserva=reserva,
        form=form,
        next_url=next_url,
    )

@bp.route('/reservas/editar/')
@login_required
def consultar_editar_reservas():
    _require_admin()
    reserva_id = request.args.get('id', type=int)
    cpf_input = (request.args.get('cpf') or '').strip()
    nome_input = (request.args.get('nome') or '').strip()
    pesquisou = bool(reserva_id or cpf_input or nome_input)
    consultar = []
    primary_reserva = None
    outras_reservas = []

    def _format_cpf_display(value):
        digits = ''.join(ch for ch in str(value or '') if ch.isdigit())
        if len(digits) == 11:
            return f'{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}'
        return (value or '')

    def _cpf_digits(value):
        return ''.join(ch for ch in (value or '') if ch.isdigit())

    def _clean(field):
        return func.replace(func.replace(func.replace(field, '.', ''), '-', ''), ' ', '')

    if reserva_id:
        reserva = BaseDados.query.filter_by(id=reserva_id).first()
        cpf_vals = []
        if reserva:
            cpf_vals.extend([_cpf_digits(reserva.cpf), _cpf_digits(reserva.cpf_pagante)])
        cpf_vals = [v for v in dict.fromkeys(cpf_vals) if len(v) == 11]
        if cpf_vals:
            consultar = BaseDados.query.filter(
                or_(_clean(BaseDados.cpf).in_(cpf_vals), _clean(BaseDados.cpf_pagante).in_(cpf_vals))
            ).all()
            primary_reserva = reserva
            outras_reservas = [r for r in consultar if r.id != reserva_id]
            status_order = {
                'Pendente': 1,
                'Aprovada': 2,
                'Checkin': 3,
                'Pago': 4,
                'Recusada': 5,
                'Expirada': 6,
            }
            outras_reservas.sort(key=lambda r: (status_order.get(r.status_reserva or '', 99), -(r.id or 0)))
        else:
            consultar = BaseDados.query.filter_by(id=reserva_id).all()
            primary_reserva = reserva
    elif cpf_input:
        cpf_digits = _cpf_digits(cpf_input)
        if len(cpf_digits) == 11:
            consultar = BaseDados.query.filter(
                or_(_clean(BaseDados.cpf) == cpf_digits, _clean(BaseDados.cpf_pagante) == cpf_digits)
            ).all()
    elif nome_input:
        consultar = (
            BaseDados.query
            .filter(BaseDados.nome.ilike(f'%{nome_input}%'))
            .order_by(BaseDados.id.desc())
            .all()
        )

    return render_template(
        'portal/editar_reservas_consulta.html',
        consultar=consultar,
        primary_reserva=primary_reserva,
        outras_reservas=outras_reservas,
        reserva_id=reserva_id or '',
        cpf=cpf_input,
        nome=nome_input,
        pesquisou=pesquisou,
        format_cpf_display=_format_cpf_display,
    )


@bp.route('/reservas/nova/', methods=['GET', 'POST'])
@login_required
def nova_reserva():
    from forms.reservas import AdminReservasForm

    mhex_param = (request.args.get('mhex') or '').strip().upper()
    if mhex_param not in [MHEX_HTM_01, MHEX_HTM_02]:
        mhex_param = MHEX_HTM_01
    mhex_locked = request.args.get('lock') == '1'

    form = AdminReservasForm()

    # Ajustar UHs conforme hotel selecionado
    selected_mhex = mhex_param if mhex_locked else (form.mhex.data or mhex_param)
    entrada_value = form.entrada.data
    saida_value = form.saida.data
    from services.availability import get_available_uhs
    if entrada_value and saida_value:
        available_uhs = get_available_uhs(selected_mhex, entrada_value, saida_value)
        form.uh.choices = [('', '---')] + available_uhs
    else:
        form.uh.choices = [('', '---')] + get_uh_choices(selected_mhex)

    if request.method == 'POST':
        if form.validate_on_submit():
            try:
                reserva = BaseDados()
                form.normalize_quartel_funcao()
                form.populate_obj(reserva)

                if not reserva.id:
                    reserva.id = None

                if not reserva.qtde_quartos:
                    reserva.qtde_quartos = 1

                reserva.mhex = mhex_param if mhex_locked else (form.mhex.data or mhex_param)
                reserva.status_reserva = STATUS_RESERVA_APROVADA

                calculate_total(reserva)
                db.session.add(reserva)
                db.session.commit()
                log_reserva_etapa(
                    reserva.id,
                    "RESERVA",
                    "CRIADA",
                    details={"origem": "admin", "status_reserva": reserva.status_reserva},
                )
                flash("Reserva criada com sucesso!", "success")
                return redirect(url_for('recepcao.ocupacao_hoje'))
            except Exception:
                db.session.rollback()
                flash("Erro ao salvar a reserva.", "error")
        else:
            flash("Dados inválidos. Verifique os campos.", "error")

    form.mhex.data = selected_mhex
    if not form.qtde_quartos.data:
        form.qtde_quartos.data = 1

    return render_template('portal/reserva_interna.html', form=form, mhex_locked=(mhex_param if mhex_locked else None))

@bp.route('/recepcao/checkin/')
@login_required
def recepcao_checkin():
    consultar = BaseDados.query.filter_by(status_reserva="Aprovada").all()
    consultar_checkin = BaseDados.query.filter_by(status_reserva="Checkin").all()
    return render_template('portal/recepcao_checkin.html', 
                           consultar=consultar, 
                           consultar_checkin=consultar_checkin)

@bp.route('/recepcao/checkout/')
@login_required
def recepcao_checkout():
    consultar = BaseDados.query.filter_by(status_reserva="Aprovada").all()
    consultar_checkin = BaseDados.query.filter_by(status_reserva="Checkin").all()
    return render_template('portal/recepcao_checkout.html', 
                           consultar=consultar, 
                           consultar_checkin=consultar_checkin)

@bp.route('/recepcao/ocupacao/hoje/')
@login_required
def ocupacao_hoje():
    expire_reservas()
    rooms = _build_ocupacao_hoje(MHEX_HTM_01)
    return render_template('portal/ocupacao_hoje.html', rooms=rooms, hotel_label='HTM 01')

@bp.route('/recepcao/ocupacao/semanal/')
@login_required
def ocupacao_semanal():
    expire_reservas()
    rooms, rooms_data, days, day_dates, start_date, end_date = _build_ocupacao_semanal(MHEX_HTM_01)

    return render_template(
        'portal/ocupacao_semanal.html',
        days=days,
        day_dates=day_dates,
        rooms=rooms,
        rooms_data=rooms_data,
        start_date=start_date.strftime("%d/%m/%Y"),
        end_date=end_date.strftime("%d/%m/%Y"),
        hotel_label='HTM 01',
    )

@bp.route('/recepcao/ocupacao/mensal/')
@login_required
def ocupacao_mensal():
    expire_reservas()
    today = date.today()
    month = request.args.get("month", type=int) or today.month
    year = request.args.get("year", type=int) or today.year

    if month < 1 or month > 12:
        month = today.month

    rooms, rooms_data, days, month_label, last_day = _build_ocupacao_mensal(MHEX_HTM_01, month, year)

    return render_template(
        'portal/ocupacao_mensal.html',
        days=days,
        rooms=rooms,
        rooms_data=rooms_data,
        month=month,
        year=year,
        month_label=month_label,
        hotel_label='HTM 01',
    )


@bp.route('/recepcao/htm02/ocupacao/hoje/')
@login_required
def ocupacao_hoje_htm02():
    expire_reservas()
    rooms = _build_ocupacao_hoje(MHEX_HTM_02)
    return render_template('portal/ocupacao_hoje.html', rooms=rooms, hotel_label='HTM 02')


@bp.route('/recepcao/ocupacao/hoje/geral/')
@login_required
def ocupacao_hoje_geral():
    expire_reservas()
    rooms = _build_ocupacao_hoje(MHEX_HTM_01)
    return render_template('portal/ocupacao_hoje.html', rooms=rooms, hotel_label='HTM 01')


@bp.route('/recepcao/ocupacao/hoje/geral/htm02/')
@login_required
def ocupacao_hoje_geral_htm02():
    expire_reservas()
    rooms = _build_ocupacao_hoje(MHEX_HTM_02)
    return render_template('portal/ocupacao_hoje.html', rooms=rooms, hotel_label='HTM 02')


@bp.route('/recepcao/htm02/ocupacao/semanal/')
@login_required
def ocupacao_semanal_htm02():
    expire_reservas()
    rooms, rooms_data, days, day_dates, start_date, end_date = _build_ocupacao_semanal(MHEX_HTM_02)
    return render_template(
        'portal/ocupacao_semanal.html',
        days=days,
        day_dates=day_dates,
        rooms=rooms,
        rooms_data=rooms_data,
        start_date=start_date.strftime("%d/%m/%Y"),
        end_date=end_date.strftime("%d/%m/%Y"),
        hotel_label='HTM 02',
    )

@bp.route('/recepcao/htm02/ocupacao/mensal/')
@login_required
def ocupacao_mensal_htm02():
    expire_reservas()
    today = date.today()
    month = request.args.get("month", type=int) or today.month
    year = request.args.get("year", type=int) or today.year

    if month < 1 or month > 12:
        month = today.month

    rooms, rooms_data, days, month_label, last_day = _build_ocupacao_mensal(MHEX_HTM_02, month, year)
    return render_template(
        'portal/ocupacao_mensal.html',
        days=days,
        rooms=rooms,
        rooms_data=rooms_data,
        month=month,
        year=year,
        month_label=month_label,
        hotel_label='HTM 02',
    )

@bp.route('/reservas/editar/<int:id>/', methods=['GET', 'POST'])
@login_required
def editar_reservas(id):
    _require_admin()
    reserva = BaseDados.query.get_or_404(id)
    from forms.reservas import ReservaEdicaoRapidaForm
    from services.availability import get_available_uhs
    
    form = ReservaEdicaoRapidaForm(obj=reserva)
    next_url = _safe_next_url(request.args.get('next'))
    
    # Filtrar UHs disponíveis para o período e MHEx atual
    mhex_value = form.mhex.data or reserva.mhex
    entrada_value = form.entrada.data or reserva.entrada
    saida_value = form.saida.data or reserva.saida
    if mhex_value and entrada_value and saida_value:
        available_uhs = get_available_uhs(mhex_value, entrada_value, saida_value, exclude_id=id)
        form.uh.choices = [('', '---')] + available_uhs
    else:
        form.uh.choices = [('', '---')]

    if reserva.uh is not None:
        current_uh = str(reserva.uh).strip()
        if current_uh and current_uh not in [c[0] for c in form.uh.choices]:
            form.uh.choices.append((current_uh, current_uh))
        if request.method == 'GET':
            form.uh.data = current_uh
    
    # Buscar acompanhantes para o contexto (lógica similar ao Django)
    acompanhantes = []
    if reserva.qtde_hosp > 1:
        for i in range(1, 6):
            nome = getattr(reserva, f'nome_acomp{i}', None)
            if nome:
                acompanhantes.append({
                    'nome': nome,
                    'idade': getattr(reserva, f'idade_acomp{i}', ''),
                    'parentesco': getattr(reserva, f'vinculo_acomp{i}', '')
                })

    if form.validate_on_submit():
        prev_status = reserva.status_reserva
        reserva.entrada = form.entrada.data
        reserva.saida = form.saida.data
        reserva.mhex = form.mhex.data

        if form.status_reserva.data:
            reserva.status_reserva = form.status_reserva.data

        if form.uh.data is not None:
            reserva.uh = (form.uh.data or '').strip() or None

        db.session.commit()
        acao = "ACEITA" if (prev_status != STATUS_RESERVA_APROVADA and reserva.status_reserva == STATUS_RESERVA_APROVADA) else "ALTERADA"
        log_reserva_etapa(reserva.id, "RESERVA", acao, details={"origem": "editar_reservas", "status_reserva": reserva.status_reserva})
        flash('Reserva atualizada com sucesso.', 'success')
        if next_url:
            return redirect(next_url)
        return redirect(url_for('recepcao.consultar_editar_reservas', id=reserva.id))
    elif request.method == 'POST':
        flash('Erro ao salvar. Verifique os campos.', 'error')
        for field, errors in form.errors.items():
            for err in errors:
                flash(f'{field}: {err}', 'error')
    
    return render_template('portal/reserva_edit.html', reserva=reserva, form=form, acompanhantes=acompanhantes, next_url=next_url)

@bp.route('/reservas/editar/admin/<int:id>/', methods=['GET', 'POST'])
@login_required
def editar_reservas_admin(id):
    _require_admin()
    reserva = BaseDados.query.get_or_404(id)
    from forms.reservas import AdminReservasForm
    from services.availability import get_available_uhs
    from models import Produto
    from constants import PRODUTO_AGUA, PRODUTO_REFRIGERANTE, PRODUTO_CERVEJA, PRODUTO_PET
    
    form = AdminReservasForm(obj=reserva)
    if request.method == 'GET':
        pagante_principal = True
        if (reserva.nome_pagante or '').strip() and (reserva.nome_pagante or '').strip() != (reserva.nome or '').strip():
            pagante_principal = False
        if (reserva.cpf_pagante or '').strip() and (reserva.cpf_pagante or '').strip() != (reserva.cpf or '').strip():
            pagante_principal = False
        form.pagante_checkbox.data = not pagante_principal
    
    mhex_value = form.mhex.data or reserva.mhex
    entrada_value = form.entrada.data or reserva.entrada
    saida_value = form.saida.data or reserva.saida
    if mhex_value and entrada_value and saida_value:
        available_uhs = get_available_uhs(mhex_value, entrada_value, saida_value, exclude_id=id)
        form.uh.choices = [('', '---')] + available_uhs
    else:
        form.uh.choices = [('', '---')]

    if reserva.uh is not None:
        current_uh = str(reserva.uh).strip()
        if current_uh and current_uh not in [c[0] for c in form.uh.choices]:
            form.uh.choices.append((current_uh, current_uh))
        if request.method == 'GET':
            form.uh.data = current_uh

    acompanhantes = []
    if reserva.qtde_hosp > 1:
        for i in range(1, 6):
            nome = getattr(reserva, f'nome_acomp{i}', None)
            if nome:
                acompanhantes.append({
                    'nome': nome,
                    'idade': getattr(reserva, f'idade_acomp{i}', ''),
                    'parentesco': getattr(reserva, f'vinculo_acomp{i}', '')
                })

    produto_agua = Produto.query.filter_by(nome=PRODUTO_AGUA).first()
    produto_refri = Produto.query.filter_by(nome=PRODUTO_REFRIGERANTE).first()
    produto_cerveja = Produto.query.filter_by(nome=PRODUTO_CERVEJA).first()
    produto_pet = Produto.query.filter_by(nome=PRODUTO_PET).first()

    if form.validate_on_submit():
        prev_status = reserva.status_reserva
        form.populate_obj(reserva)

        if not form.pagante_checkbox.data:
            reserva.nome_pagante = reserva.nome
            reserva.cpf_pagante = reserva.cpf

        if reserva.entrada and reserva.saida:
            reserva.diarias = max((reserva.saida - reserva.entrada).days, 0)
        calculate_total(reserva)

        db.session.commit()
        acao = "ACEITA" if (prev_status != STATUS_RESERVA_APROVADA and reserva.status_reserva == STATUS_RESERVA_APROVADA) else "ALTERADA"
        log_reserva_etapa(reserva.id, "RESERVA", acao, details={"origem": "editar_reservas_admin", "status_reserva": reserva.status_reserva})
        flash('Reserva atualizada com sucesso.', 'success')
        return redirect(url_for('recepcao.consultar_editar_reservas', id=reserva.id))
    elif request.method == 'POST':
        flash('Erro ao salvar. Verifique os campos.', 'error')
        for field, errors in form.errors.items():
            for err in errors:
                flash(f'{field}: {err}', 'error')

    if reserva.entrada and reserva.saida:
        reserva.diarias = max((reserva.saida - reserva.entrada).days, 0)
        calculate_total(reserva)

    return render_template(
        'portal/editar_reserva_full.html',
        reserva=reserva,
        form=form,
        acompanhantes=acompanhantes,
        preco_agua=float(produto_agua.valor) if produto_agua and produto_agua.valor is not None else 0,
        preco_refri=float(produto_refri.valor) if produto_refri and produto_refri.valor is not None else 0,
        preco_cerveja=float(produto_cerveja.valor) if produto_cerveja and produto_cerveja.valor is not None else 0,
        preco_pet=float(produto_pet.valor) if produto_pet and produto_pet.valor is not None else 0,
    )

@bp.route('/checkin/<int:id>/', methods=['GET', 'POST'])
@login_required
def editar_checkin(id):
    reserva = BaseDados.query.get_or_404(id)
    form = CheckinForm(obj=reserva)
    from services.availability import get_available_uhs

    next_url = _safe_next_url(request.args.get('next'))

    selected_mhex = (form.mhex.data or reserva.mhex or MHEX_HTM_01)
    entrada_value = form.entrada.data
    saida_value = form.saida.data
    if entrada_value and saida_value:
        available_uhs = get_available_uhs(selected_mhex, entrada_value, saida_value, exclude_id=reserva.id)
        form.uh.choices = [('', '---')] + available_uhs
    else:
        form.uh.choices = [('', '---')] + get_uh_choices(selected_mhex)

    if form.validate_on_submit():
        prev_status = reserva.status_reserva
        before = {
            "entrada": reserva.entrada,
            "saida": reserva.saida,
            "qtde_hosp": reserva.qtde_hosp,
            "qtde_quartos": reserva.qtde_quartos,
            "especial": reserva.especial,
            "nome": reserva.nome,
            "cpf": reserva.cpf,
            "telefone": reserva.telefone,
            "graduacao": reserva.graduacao,
            "status": reserva.status,
            "tipo": reserva.tipo,
            "cidade": reserva.cidade,
            "uf": reserva.uf,
            "mhex": reserva.mhex,
            "uh": reserva.uh,
        }
        form.populate_obj(reserva)
        reserva.status_reserva = STATUS_RESERVA_CHECKIN
        if reserva.entrada and reserva.saida:
            delta = (reserva.saida - reserva.entrada).days
            reserva.diarias = max(delta, 0)
        db.session.commit()
        changed_fields = _diff_changed_fields(before, reserva, list(before.keys()))
        if prev_status == STATUS_RESERVA_CHECKIN:
            if changed_fields:
                log_reserva_etapa(reserva.id, "ALTERACAO", "CHECKIN", details={"campos": changed_fields})
        else:
            if changed_fields:
                log_reserva_etapa(reserva.id, "ALTERACAO", "CHECKIN", details={"campos": changed_fields})
            log_reserva_etapa(reserva.id, "CHECKIN", "EFETUADO", details={"status_reserva": reserva.status_reserva})
        if next_url:
            return redirect(next_url)
        if not (current_user.is_superuser or current_user.is_staff):
            if reserva.mhex == MHEX_HTM_02:
                return redirect(url_for('recepcao.ocupacao_hoje_geral_htm02'))
            return redirect(url_for('recepcao.ocupacao_hoje_geral'))
        if reserva.mhex == MHEX_HTM_02:
            return redirect(url_for('recepcao.ocupacao_hoje_htm02'))
        return redirect(url_for('recepcao.ocupacao_hoje'))
    return render_template('portal/checkin.html', reserva=reserva, form=form, next_url=next_url)

@bp.route('/checkout/<int:id>/', methods=['GET', 'POST'])
@login_required
def editar_checkout(id):
    reserva = BaseDados.query.get_or_404(id)
    form = CheckoutForm(obj=reserva)
    next_url = _safe_next_url(request.args.get('next'))
    if reserva.entrada and reserva.saida:
        reserva.diarias = max((reserva.saida - reserva.entrada).days, 0)
        calculate_total(reserva)

    def _format_cpf(value):
        if not value:
            return ""
        digits = "".join([c for c in str(value) if c.isdigit()])[:11]
        if len(digits) < 11:
            return value
        return f"{digits[0:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:11]}"

    def _format_tel(value):
        if not value:
            return ""
        digits = "".join([c for c in str(value) if c.isdigit()])
        if len(digits) >= 11:
            return f"({digits[0:2]}){digits[2:7]}-{digits[7:11]}"
        if len(digits) >= 10:
            return f"({digits[0:2]}){digits[2:6]}-{digits[6:10]}"
        return value
    from models import Produto
    from constants import PRODUTO_AGUA, PRODUTO_REFRIGERANTE, PRODUTO_CERVEJA, PRODUTO_PET
    produto_agua = Produto.query.filter_by(nome=PRODUTO_AGUA).first()
    produto_refri = Produto.query.filter_by(nome=PRODUTO_REFRIGERANTE).first()
    produto_cerveja = Produto.query.filter_by(nome=PRODUTO_CERVEJA).first()
    produto_pet = Produto.query.filter_by(nome=PRODUTO_PET).first()

    if form.validate_on_submit():
        prev_status = reserva.status_reserva
        before = {
            "qtde_agua": reserva.qtde_agua,
            "qtde_refri": reserva.qtde_refri,
            "qtde_cerveja": reserva.qtde_cerveja,
            "qtde_pet": reserva.qtde_pet,
            "forma_pagamento": reserva.forma_pagamento,
            "nome_pagante": reserva.nome_pagante,
            "cpf_pagante": reserva.cpf_pagante,
        }
        form.populate_obj(reserva)

        if not form.pagante_tipo.data:
            flash("Selecione quem vai pagar.", "error")
            return render_template(
                'portal/checkout.html',
                reserva=reserva,
                form=form,
                cpf_formatado=_format_cpf(reserva.cpf),
                telefone_formatado=_format_tel(reserva.telefone),
                preco_agua=float(produto_agua.valor) if produto_agua and produto_agua.valor is not None else 0,
                preco_refri=float(produto_refri.valor) if produto_refri and produto_refri.valor is not None else 0,
                preco_cerveja=float(produto_cerveja.valor) if produto_cerveja and produto_cerveja.valor is not None else 0,
                preco_pet=float(produto_pet.valor) if produto_pet and produto_pet.valor is not None else 0,
            )

        if form.pagante_tipo.data == 'outro':
            if not (reserva.nome_pagante or '').strip() or not (reserva.cpf_pagante or '').strip():
                flash("Informe nome e CPF do pagante.", "error")
                return render_template(
                    'portal/checkout.html',
                    reserva=reserva,
                    form=form,
                    cpf_formatado=_format_cpf(reserva.cpf),
                    telefone_formatado=_format_tel(reserva.telefone),
                    preco_agua=float(produto_agua.valor) if produto_agua and produto_agua.valor is not None else 0,
                    preco_refri=float(produto_refri.valor) if produto_refri and produto_refri.valor is not None else 0,
                    preco_cerveja=float(produto_cerveja.valor) if produto_cerveja and produto_cerveja.valor is not None else 0,
                    preco_pet=float(produto_pet.valor) if produto_pet and produto_pet.valor is not None else 0,
                )
        else:
            reserva.nome_pagante = reserva.nome
            reserva.cpf_pagante = reserva.cpf

        if reserva.entrada and reserva.saida:
            reserva.diarias = max((reserva.saida - reserva.entrada).days, 0)
        calculate_total(reserva)

        reserva.status_reserva = STATUS_RESERVA_PAGO
            
        db.session.commit()
        changed_fields = _diff_changed_fields(before, reserva, list(before.keys()))
        if prev_status == STATUS_RESERVA_PAGO:
            if changed_fields:
                log_reserva_etapa(reserva.id, "ALTERACAO", "CHECKOUT", details={"campos": changed_fields})
        else:
            log_reserva_etapa(
                reserva.id,
                "CHECKOUT",
                "EFETUADO",
                details={"status_reserva": reserva.status_reserva, "forma_pagamento": reserva.forma_pagamento},
            )
        if next_url:
            return redirect(next_url)
        if not (current_user.is_superuser or current_user.is_staff):
            if reserva.mhex == MHEX_HTM_02:
                return redirect(url_for('recepcao.ocupacao_hoje_geral_htm02'))
            return redirect(url_for('recepcao.ocupacao_hoje_geral'))
        if reserva.mhex == MHEX_HTM_02:
            return redirect(url_for('recepcao.ocupacao_hoje_htm02'))
        return redirect(url_for('recepcao.ocupacao_hoje'))
    if request.method == 'GET':
        if not form.pagante_tipo.data:
            if reserva.cpf_pagante and reserva.cpf and str(reserva.cpf_pagante).strip() != str(reserva.cpf).strip():
                form.pagante_tipo.data = 'outro'
            else:
                form.pagante_tipo.data = ''
                form.nome_pagante.data = ''
                form.cpf_pagante.data = ''

        # Força o operador a selecionar a forma de pagamento na tela de checkout
        form.forma_pagamento.data = ''

    return render_template(
        'portal/checkout.html',
        reserva=reserva,
        form=form,
        next_url=next_url,
        cpf_formatado=_format_cpf(reserva.cpf),
        telefone_formatado=_format_tel(reserva.telefone),
        preco_agua=float(produto_agua.valor) if produto_agua and produto_agua.valor is not None else 0,
        preco_refri=float(produto_refri.valor) if produto_refri and produto_refri.valor is not None else 0,
        preco_cerveja=float(produto_cerveja.valor) if produto_cerveja and produto_cerveja.valor is not None else 0,
        preco_pet=float(produto_pet.valor) if produto_pet and produto_pet.valor is not None else 0,
    )

@bp.route('/consumacao/<int:id>/', methods=['GET', 'POST'])
@login_required
def editar_consumacao(id):
    reserva = BaseDados.query.get_or_404(id)
    form = ConsumacaoForm(obj=reserva)
    if form.validate_on_submit():
        form.populate_obj(reserva)
        calculate_total(reserva)
        db.session.commit()
        log_reserva_etapa(reserva.id, "CHECKOUT", "ALTERADA", details={"tipo": "consumacao"})
        return redirect(url_for('recepcao.recepcao_checkout'))
    return render_template('portal/consumacao.html', reserva=reserva, form=form)
