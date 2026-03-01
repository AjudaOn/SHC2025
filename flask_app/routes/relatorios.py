import io
from flask import Blueprint, render_template, request, send_file, redirect, url_for, flash
from flask_login import login_required
from sqlalchemy import extract, func
from models import db, BaseDados
from constants import get_uh_choices, MHEX_HTM_01
from decimal import Decimal, ROUND_HALF_UP
from datetime import date, timedelta
import calendar
from services.reports import ReportService

bp = Blueprint('relatorios', __name__)

@bp.route('/relatorio/pagamento/')
@login_required
def relatorio_pagamento():
    return render_template('portal/relatorio_pagamento.html')

@bp.route('/relatorio/mensal/')
@login_required
def relatorio_mensal():
    mes = request.args.get('mes_relatorio', type=int)
    ano = request.args.get('ano_relatorio', type=int)
    forma_pagamento = request.args.get('forma_pagamento')
    
    if not mes or not ano or not forma_pagamento:
        flash("Selecione mês, ano e forma de pagamento para gerar o relatório.", "info")
        return render_template('portal/relatorio_pagamento.html')
    
    forma_pagamento = forma_pagamento.upper()

    # Fetch data for HTML report
    consultar_htm1 = BaseDados.query.filter(
        BaseDados.status_reserva == "Pago",
        BaseDados.mhex == "HTM_01",
        BaseDados.forma_pagamento == forma_pagamento,
        extract('month', BaseDados.saida) == mes,
        extract('year', BaseDados.saida) == ano
    ).all()

    total_htm1 = db.session.query(func.sum(BaseDados.valor_total)).filter(
        BaseDados.status_reserva == "Pago",
        BaseDados.mhex == "HTM_01",
        BaseDados.forma_pagamento == forma_pagamento,
        extract('month', BaseDados.saida) == mes,
        extract('year', BaseDados.saida) == ano
    ).scalar() or 0

    consultar_htm2 = BaseDados.query.filter(
        BaseDados.status_reserva == "Pago",
        BaseDados.mhex == "HTM_02",
        BaseDados.forma_pagamento == forma_pagamento,
        extract('month', BaseDados.saida) == mes,
        extract('year', BaseDados.saida) == ano
    ).all()

    total_htm2 = db.session.query(func.sum(BaseDados.valor_total)).filter(
        BaseDados.status_reserva == "Pago",
        BaseDados.mhex == "HTM_02",
        BaseDados.forma_pagamento == forma_pagamento,
        extract('month', BaseDados.saida) == mes,
        extract('year', BaseDados.saida) == ano
    ).scalar() or 0

    return render_template('portal/relatorio_mensal.html', 
                           consultar_htm1=consultar_htm1, 
                           consultar_htm2=consultar_htm2,
                           total_htm1=total_htm1,
                           total_htm2=total_htm2)

@bp.route('/relatorio/pagamento/pix/')
@login_required
def relatorio_pagamento_pix():
    mes = request.args.get('mes_relatorio', type=int)
    ano = request.args.get('ano_relatorio', type=int)
    consultar_htm1 = []
    consultar_htm2 = []
    total_htm1 = 0
    total_htm2 = 0

    if mes and ano:
        consultar_htm1 = BaseDados.query.filter(
            BaseDados.status_reserva == "Pago",
            BaseDados.mhex == "HTM_01",
            BaseDados.forma_pagamento == "PIX",
            extract('month', BaseDados.saida) == mes,
            extract('year', BaseDados.saida) == ano
        ).order_by(BaseDados.saida).all()

        total_htm1 = db.session.query(func.sum(BaseDados.valor_total)).filter(
            BaseDados.status_reserva == "Pago",
            BaseDados.mhex == "HTM_01",
            BaseDados.forma_pagamento == "PIX",
            extract('month', BaseDados.saida) == mes,
            extract('year', BaseDados.saida) == ano
        ).scalar() or 0

        consultar_htm2 = BaseDados.query.filter(
            BaseDados.status_reserva == "Pago",
            BaseDados.mhex == "HTM_02",
            BaseDados.forma_pagamento == "PIX",
            extract('month', BaseDados.saida) == mes,
            extract('year', BaseDados.saida) == ano
        ).order_by(BaseDados.saida).all()

        total_htm2 = db.session.query(func.sum(BaseDados.valor_total)).filter(
            BaseDados.status_reserva == "Pago",
            BaseDados.mhex == "HTM_02",
            BaseDados.forma_pagamento == "PIX",
            extract('month', BaseDados.saida) == mes,
            extract('year', BaseDados.saida) == ano
        ).scalar() or 0

    meses = [
        (1, 'Janeiro'), (2, 'Fevereiro'), (3, 'Março'), (4, 'Abril'),
        (5, 'Maio'), (6, 'Junho'), (7, 'Julho'), (8, 'Agosto'),
        (9, 'Setembro'), (10, 'Outubro'), (11, 'Novembro'), (12, 'Dezembro')
    ]
    anos = list(range(2024, 2031))

    return render_template(
        'portal/relatorio_pagamento_pix.html',
        consultar_htm1=consultar_htm1,
        consultar_htm2=consultar_htm2,
        total_htm1=total_htm1,
        total_htm2=total_htm2,
        mes_relatorio=mes,
        ano_relatorio=ano,
        meses=meses,
        anos=anos,
    )

@bp.route('/relatorio/pagamento/dinheiro/')
@login_required
def relatorio_pagamento_dinheiro():
    mes = request.args.get('mes_relatorio', type=int)
    ano = request.args.get('ano_relatorio', type=int)
    consultar_htm1 = []
    consultar_htm2 = []
    total_htm1 = 0
    total_htm2 = 0

    if mes and ano:
        consultar_htm1 = BaseDados.query.filter(
            BaseDados.status_reserva == "Pago",
            BaseDados.mhex == "HTM_01",
            BaseDados.forma_pagamento == "DINHEIRO",
            extract('month', BaseDados.saida) == mes,
            extract('year', BaseDados.saida) == ano
        ).order_by(BaseDados.saida).all()

        total_htm1 = db.session.query(func.sum(BaseDados.valor_total)).filter(
            BaseDados.status_reserva == "Pago",
            BaseDados.mhex == "HTM_01",
            BaseDados.forma_pagamento == "DINHEIRO",
            extract('month', BaseDados.saida) == mes,
            extract('year', BaseDados.saida) == ano
        ).scalar() or 0

        consultar_htm2 = BaseDados.query.filter(
            BaseDados.status_reserva == "Pago",
            BaseDados.mhex == "HTM_02",
            BaseDados.forma_pagamento == "DINHEIRO",
            extract('month', BaseDados.saida) == mes,
            extract('year', BaseDados.saida) == ano
        ).order_by(BaseDados.saida).all()

        total_htm2 = db.session.query(func.sum(BaseDados.valor_total)).filter(
            BaseDados.status_reserva == "Pago",
            BaseDados.mhex == "HTM_02",
            BaseDados.forma_pagamento == "DINHEIRO",
            extract('month', BaseDados.saida) == mes,
            extract('year', BaseDados.saida) == ano
        ).scalar() or 0

    meses = [
        (1, 'Janeiro'), (2, 'Fevereiro'), (3, 'Março'), (4, 'Abril'),
        (5, 'Maio'), (6, 'Junho'), (7, 'Julho'), (8, 'Agosto'),
        (9, 'Setembro'), (10, 'Outubro'), (11, 'Novembro'), (12, 'Dezembro')
    ]
    anos = list(range(2024, 2031))

    return render_template(
        'portal/relatorio_pagamento_dinheiro.html',
        consultar_htm1=consultar_htm1,
        consultar_htm2=consultar_htm2,
        total_htm1=total_htm1,
        total_htm2=total_htm2,
        mes_relatorio=mes,
        ano_relatorio=ano,
        meses=meses,
        anos=anos,
    )


def _build_ocupacao_financeira(mes, ano, hotel):
    dias_no_mes = calendar.monthrange(ano, mes)[1]
    primeiro_dia_mes = date(ano, mes, 1)
    ultimo_dia_mes = date(ano, mes, dias_no_mes)

    room_keys = [choice[0] for choice in get_uh_choices(hotel)]
    rooms = [str(key).zfill(2) for key in room_keys]

    rooms_data = []
    for key in rooms:
        rooms_data.append(
            {
                "room_number": key,
                "statuses": ["available"] * dias_no_mes,
                "daily_values": [Decimal('0.00')] * dias_no_mes,
            }
        )

    room_index = {room["room_number"]: room for room in rooms_data}

    reservas = BaseDados.query.filter(
        BaseDados.mhex == hotel,
        BaseDados.status_reserva == "Pago",
    ).all()

    for reserva in reservas:
        entrada = reserva.entrada
        saida = reserva.saida
        if not entrada or not saida:
            continue

        saida_ajustada = saida - timedelta(days=1)
        diarias = (saida_ajustada - entrada).days + 1
        if diarias <= 0:
            continue

        try:
            quarto = str(int(reserva.uh)).zfill(2)
        except (TypeError, ValueError):
            continue

        if quarto not in room_index:
            continue

        valor_total = Decimal(str(reserva.valor_total or 0))
        valor_dia = (valor_total / Decimal(diarias)).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP)

        if entrada <= ultimo_dia_mes and saida_ajustada >= primeiro_dia_mes:
            inicio = max(entrada, primeiro_dia_mes)
            fim = min(saida_ajustada, ultimo_dia_mes)

            current = inicio
            while current <= fim:
                day_index = current.day - 1
                room_entry = room_index[quarto]
                room_entry["statuses"][day_index] = "paid"
                room_entry["daily_values"][day_index] += valor_dia
                current += timedelta(days=1)

    for room in rooms_data:
        room["daily_values"] = [
            value.quantize(Decimal('0.00'), rounding=ROUND_HALF_UP) for value in room["daily_values"]
        ]

    days = [str(i) for i in range(1, dias_no_mes + 1)]
    return days, rooms, rooms_data


def _obter_censo_mes(mes, ano, hotel):
    censo_ano = _obter_censo_ano(ano, hotel)
    return censo_ano.get(mes, [])


def _obter_censo_ano(ano, hotel):
    relatorio = {
        mes: [{'dia': dia, 'qtde': 0, 'valor': Decimal('0.00')} for dia in range(1, calendar.monthrange(ano, mes)[1] + 1)]
        for mes in range(1, 13)
    }

    inicio_ano = date(ano, 1, 1)
    fim_ano = date(ano, 12, 31)

    reservas = BaseDados.query.filter(
        BaseDados.mhex == hotel,
        BaseDados.status_reserva == "Pago",
        BaseDados.entrada <= fim_ano,
        BaseDados.saida >= inicio_ano,
    ).all()

    for reserva in reservas:
        entrada = reserva.entrada
        saida = reserva.saida
        if not entrada or not saida or not reserva.valor_total:
            continue

        dias_reserva = (saida - entrada).days
        if dias_reserva <= 0:
            continue

        valor_total = Decimal(str(reserva.valor_total))
        valor_diario = valor_total // dias_reserva
        sobra = valor_total % dias_reserva
        qtde_hospedes = reserva.qtde_hosp or 0

        data_atual = entrada
        while data_atual < saida:
            if data_atual.year == ano:
                mes = data_atual.month
                dia = data_atual.day
                valor_dia = valor_diario
                if data_atual == saida - timedelta(days=1):
                    valor_dia = valor_diario + sobra

                relatorio[mes][dia - 1]['qtde'] += qtde_hospedes
                relatorio[mes][dia - 1]['valor'] += valor_dia

            data_atual += timedelta(days=1)

    for mes in relatorio.values():
        for dia in mes:
            if dia['valor'] == 0:
                dia['qtde'] = 0

    return relatorio


def _obter_relatorio_uh_ano(ano, hotel):
    relatorio = {
        mes: [{'dia': dia, 'qtde': 0, 'uh': 0} for dia in range(1, calendar.monthrange(ano, mes)[1] + 1)]
        for mes in range(1, 13)
    }

    inicio_ano = date(ano, 1, 1)
    fim_ano = date(ano, 12, 31)

    reservas = BaseDados.query.filter(
        BaseDados.mhex == hotel,
        BaseDados.status_reserva == "Pago",
        BaseDados.entrada <= fim_ano,
        BaseDados.saida >= inicio_ano,
    ).all()

    for reserva in reservas:
        if not reserva.entrada or not reserva.saida or not reserva.qtde_quartos:
            continue

        data_inicio = max(reserva.entrada, inicio_ano)
        data_fim = min(reserva.saida, fim_ano)
        if data_inicio > data_fim:
            continue

        data_atual = data_inicio
        while data_atual < data_fim:
            mes = data_atual.month
            dia = data_atual.day
            relatorio[mes][dia - 1]['qtde'] += reserva.qtde_hosp or 0
            relatorio[mes][dia - 1]['uh'] += reserva.qtde_quartos or 0
            data_atual += timedelta(days=1)

    for mes in relatorio:
        for dia in relatorio[mes]:
            if dia['uh'] == 0:
                dia['qtde'] = 0

    return relatorio


@bp.route('/relatorio/censo/hospedes-valores/')
@login_required
def relatorio_hospedes_valores():
    ano = request.args.get('ano', type=int)
    today = date.today()
    if not ano:
        ano = today.year

    if ano < 2024 or ano > 2030:
        flash("Ano deve estar entre 2024 e 2030.", "error")
        return render_template('portal/relatorio_hospedes_valores.html', relatorios=[], ano=ano)

    meses = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    dias = list(range(1, 32))
    hoteis = ['HTM_01', 'HTM_02']
    relatorios = []

    for hotel in hoteis:
        censo_ano = _obter_censo_ano(ano, hotel)
        dados_meses = []
        total_ano = {'total_hospedes': {dia: 0 for dia in dias}, 'valor_total': {dia: Decimal('0.00') for dia in dias}}

        for mes_idx, mes_nome in enumerate(meses, start=1):
            censo_mes = censo_ano.get(mes_idx, [])
            total_hospedes = {}
            valor_total = {}
            for item in censo_mes:
                total_hospedes[item['dia']] = item['qtde']
                valor_total[item['dia']] = item['valor']
                total_ano['total_hospedes'][item['dia']] += item['qtde']
                total_ano['valor_total'][item['dia']] += item['valor']

            dados_meses.append({
                'mes': mes_nome,
                'total_hospedes': total_hospedes,
                'valor_total': valor_total,
            })

        relatorios.append({
            'nome': hotel,
            'dados_meses': dados_meses,
            'total_ano': total_ano,
        })

    return render_template(
        'portal/relatorio_hospedes_valores.html',
        relatorios=relatorios,
        meses=meses,
        dias=dias,
        ano=ano,
        anos=list(range(2024, 2031)),
    )


@bp.route('/relatorio/censo/hospedes-valores/excel/<hotel>/')
@login_required
def relatorio_hospedes_valores_excel(hotel):
    ano = request.args.get('ano', type=int)
    if not ano:
        return "Ano do relatório não foi fornecido.", 400

    if ano < 2024 or ano > 2030:
        return "Ano deve estar entre 2024 e 2030.", 400

    if hotel not in ['HTM_01', 'HTM_02']:
        return "Hotel inválido.", 400

    censo_ano = _obter_censo_ano(ano, hotel)
    workbook = ReportService.generate_censo_ano_excel(ano, hotel, censo_por_mes=censo_ano)
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    filename = f"relatorio_{hotel}_{ano}.xlsx"
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


@bp.route('/relatorio/censo/uhs-ocupadas/')
@login_required
def relatorio_uhs_ocupadas():
    ano = request.args.get('ano', type=int)
    today = date.today()
    if not ano:
        ano = today.year

    if ano < 2024 or ano > 2030:
        flash("Ano deve estar entre 2024 e 2030.", "error")
        return render_template('portal/relatorio_uhs_ocupadas.html', meses=[], dias=[], ano=ano, anos=list(range(2024, 2031)))

    meses = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    dias = list(range(1, 32))

    censo_htm_01 = _obter_relatorio_uh_ano(ano, 'HTM_01')
    censo_htm_02 = _obter_relatorio_uh_ano(ano, 'HTM_02')

    dados_meses = []
    total_ano_uh = {dia: 0 for dia in dias}
    total_ano_qtd = {dia: "" for dia in dias}

    for mes_idx, mes_nome in enumerate(meses, start=1):
        total_uh = {}
        total_qtd = {}
        for dia in dias:
            uh = 0
            if dia <= len(censo_htm_01.get(mes_idx, [])):
                uh += censo_htm_01[mes_idx][dia - 1]['uh']
            if dia <= len(censo_htm_02.get(mes_idx, [])):
                uh += censo_htm_02[mes_idx][dia - 1]['uh']
            total_uh[dia] = uh
            total_qtd[dia] = ""
            total_ano_uh[dia] += uh

        dados_meses.append({
            'mes': mes_nome,
            'total_qtd': total_qtd,
            'total_uh': total_uh,
        })

    return render_template(
        'portal/relatorio_uhs_ocupadas.html',
        meses=meses,
        dias=dias,
        ano=ano,
        anos=list(range(2024, 2031)),
        dados_meses=dados_meses,
        total_ano_qtd=total_ano_qtd,
        total_ano_uh=total_ano_uh,
    )


@bp.route('/relatorio/censo/uhs-ocupadas/excel/')
@login_required
def relatorio_uhs_ocupadas_excel():
    ano = request.args.get('ano', type=int)
    if not ano:
        return "Ano do relatório não foi fornecido.", 400

    if ano < 2024 or ano > 2030:
        return "Ano deve estar entre 2024 e 2030.", 400

    censo_htm_01 = _obter_relatorio_uh_ano(ano, 'HTM_01')
    censo_htm_02 = _obter_relatorio_uh_ano(ano, 'HTM_02')
    workbook = ReportService.generate_censo_uh_ano_excel(ano, censo_htm_01, censo_htm_02)
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    filename = f"relatorio_censo_uh_{ano}.xlsx"
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


@bp.route('/relatorio/ocupacao/financeiro/')
@login_required
def relatorio_ocupacao_financeiro():
    year = request.args.get('year', type=int)
    month = request.args.get('month', type=int)

    today = date.today()
    if not year:
        year = today.year
    if not month:
        month = today.month

    days_htm1, rooms_htm1, rooms_data_htm1 = _build_ocupacao_financeira(month, year, 'HTM_01')
    days_htm2, rooms_htm2, rooms_data_htm2 = _build_ocupacao_financeira(month, year, 'HTM_02')

    return render_template(
        'portal/ocupacao_financeira.html',
        year=year,
        month=month,
        days_htm1=days_htm1,
        rooms_htm1=rooms_htm1,
        rooms_data_htm1=rooms_data_htm1,
        days_htm2=days_htm2,
        rooms_htm2=rooms_htm2,
        rooms_data_htm2=rooms_data_htm2,
    )

@bp.route('/relatorio/gerar/<forma_pagamento>/')
@login_required
def gerar_relatorio_excel(forma_pagamento):
    mes = request.args.get('mes_relatorio')
    ano = request.args.get('ano_relatorio')
    
    try:
        workbook = ReportService.generate_payment_report(mes, ano, forma_pagamento.upper())
    except ValueError as e:
        return str(e), 400

    if workbook is None:
        flash("Nenhum dado encontrado para o período selecionado.", "warning")
        destino = 'relatorios.relatorio_pagamento_pix' if forma_pagamento.lower() == 'pix' else 'relatorios.relatorio_pagamento_dinheiro'
        return redirect(url_for(destino))

    # Save to memory and send
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    
    filename = f"relatorio_{forma_pagamento.lower()}.xlsx"
    return send_file(output, 
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True,
                     download_name=filename)
