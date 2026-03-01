from decimal import Decimal
from models import Precos_status_graduacao, Precos_graduacao_vinculo, Produto
from constants import *

def calculate_total(reserva):
    """
    Calculates all financial values for a reservation.
    Updates the reservation object in place.
    """
    status = reserva.status
    graduacao = reserva.graduacao

    # 1. Calculate Guest Value (Valor Hospede)
    try:
        if status == STATUS_CIVIL:
            preco_status_graduacao = Precos_status_graduacao.query.filter_by(status=status, graduacao=graduacao).first()
            reserva.valor_hosp = preco_status_graduacao.valor if preco_status_graduacao else 0
        elif status in [STATUS_MILITAR_ATIVA, STATUS_MILITAR_RESERVA, 'DEP ACOMPANHADO', STATUS_DEP_DESACOMPANHADO, STATUS_PENSIONISTA]:                
            preco_status_graduacao = Precos_status_graduacao.query.filter_by(status=status, graduacao=graduacao).first()
            reserva.valor_hosp = preco_status_graduacao.valor if preco_status_graduacao else 0
        else:
            reserva.valor_hosp = 0
    except Exception:
        reserva.valor_hosp = 0

    # 2. Calculate Companions Value (Valor Acompanhantes)
    vinculo_acomps = [
        reserva.vinculo_acomp1, reserva.vinculo_acomp2, 
        reserva.vinculo_acomp3, reserva.vinculo_acomp4, 
        reserva.vinculo_acomp5
    ]
    
    for i, vinculo_acomp in enumerate(vinculo_acomps):
        field_name = f'valor_acomp{i+1}'
        if vinculo_acomp:
            try:
                preco_vinculo = Precos_graduacao_vinculo.query.filter_by(graduacao=graduacao, vinculo=vinculo_acomp).first()
                setattr(reserva, field_name, preco_vinculo.valor if preco_vinculo else 0)
            except Exception:
                setattr(reserva, field_name, 0)
        else:
            setattr(reserva, field_name, 0)

    # 3. Calculate Daily Rate (Valor Dia)
    def get_val(val):
        return val if isinstance(val, (int, float, Decimal)) else 0

    reserva.valor_dia = (
        get_val(reserva.valor_hosp) + 
        get_val(reserva.valor_acomp1) + 
        get_val(reserva.valor_acomp2) + 
        get_val(reserva.valor_acomp3) + 
        get_val(reserva.valor_acomp4) + 
        get_val(reserva.valor_acomp5)
    )

    # 4. Calculate Consumption (Consumação)
    try:
        produto_agua = Produto.query.filter_by(nome=PRODUTO_AGUA).first()
        produto_refri = Produto.query.filter_by(nome=PRODUTO_REFRIGERANTE).first()
        produto_cerveja = Produto.query.filter_by(nome=PRODUTO_CERVEJA).first()   
        produto_pet = Produto.query.filter_by(nome=PRODUTO_PET).first()
    except Exception:
        return

    # Helper to calculate item total
    def calc_item(qty, price):
        return (Decimal(qty) * price) if qty is not None else Decimal(0)

    reserva.total_agua = calc_item(reserva.qtde_agua, produto_agua.valor if produto_agua else 0)
    reserva.total_refri = calc_item(reserva.qtde_refri, produto_refri.valor if produto_refri else 0)
    reserva.total_cerveja = calc_item(reserva.qtde_cerveja, produto_cerveja.valor if produto_cerveja else 0)
    
    # 5. Calculate Totals based on Trip Reason (Motivo Viagem)
    diarias = reserva.diarias if reserva.diarias is not None else 0
    
    if reserva.motivo_viagem == MOTIVO_SAUDE:
        reserva.desc_saude = reserva.valor_dia / 2
        reserva.subtotal = reserva.desc_saude * diarias
             
    else:
        reserva.desc_saude = 0
        reserva.subtotal = reserva.valor_dia * diarias

    # Pet cobra uma única vez por animal (não por diária)
    if reserva.qtde_pet and reserva.qtde_pet != 0:
        reserva.total_pet = Decimal(reserva.qtde_pet) * (produto_pet.valor if produto_pet else 0)
    else:
        reserva.total_pet = 0

    # Final Sums
    reserva.total_consumacao = (
        reserva.total_agua + 
        reserva.total_refri + 
        reserva.total_cerveja
    )
    
    valor_ajuste = reserva.valor_ajuste if reserva.valor_ajuste is not None else 0
    
    reserva.valor_total = (
        reserva.subtotal + 
        valor_ajuste + 
        reserva.total_consumacao + 
        reserva.total_pet
    )
