from django.urls import path
from . import views

urlpatterns = [    
    path('reservas/consultar/', views.consultar_reservas, name='consultar_reservas'),
    path('reservas/fazer/', views.fazer_reservas, name='fazer_reservas'),
    path('reservas/editar/<int:reservas_pk>/', views.editar_reservas, name='editar_reservas'), 
    path('checkin/<int:reservas_pk>/', views.editar_checkin, name='editar_checkin'),
    path('checkout/<int:reservas_pk>/', views.editar_checkout, name='editar_checkout'),
    path('consumacao/<int:reservas_pk>/', views.editar_consumacao, name='editar_consumacao'), 
    path('recepcao/checkin/', views.recepcao_checkin, name='recepcao_checkin'),
    path('recepcao/checkout/', views.recepcao_checkout, name='recepcao_checkout'),
    path('relatorio/', views.relatorio_mensal, name='relatorio_mensal'),    
    path('relatorio_pagamento/', views.relatorio_pagamento, name='relatorio_pagamento'),
    path('relatorio_pagamento_excel/', views.relatorio_pagamento_excel, name='relatorio_pagamento_excel'),
    path('gerar_relatorio/', views.RelatorioPix.as_view(), name='gerar_relatorio'),   
    path('relatorio_pagamento_pix/', views.relatorio_pagamento_pix, name='relatorio_pagamento_pix'),
    path('gerar_relatorio_pix/', views.RelatorioPix.as_view(), name='gerar_relatorio_pix'),
    path('relatorio_pagamento_dinheiro/', views.relatorio_pagamento_dinheiro, name='relatorio_pagamento_dinheiro'),
    path('gerar_relatorio_dinheiro/', views.RelatorioDinheiro.as_view(), name='gerar_relatorio_dinheiro'),
    path('', views.reserva_externa, name='reserva_externa'),
    path('relatorio_ocupacao_financeira/', views.relatorio_ocupacao_finaceiro, name='relatorio_ocupacao_finaceiro'),
    path('ocupacao/', views.Visualizar_Ocupacao_Hotel.as_view(), name='calculo_ocupacao'),
    path('agenda_ocupacao/', views.agenda_ocupacao, name='agenda_ocupacao'),
    path('agenda/', views.Visualizar_Agenda_Ocupacao.as_view(), name='obter_agenda'),
    

    path('relatorio_todos/', views.relatorio_todos, name='relatorio_todos'),
    path('registros/consultar/', views.consultar_todos, name='consultar_todos'),    
    path('registros/editar/<int:reservas_pk>/', views.editar_todos, name='editar_todos'), 
    path('relatorio-anual-hoteis/', views.relatorio_anual_hoteis, name='relatorio_anual_hoteis'),

    path('relatorio_censo_mes/', views.Visualizar_Censo_Ano.as_view(), name='relatorio_censo_ano'),
    path('relatorio_censo_ano_excel/', views.Visualizar_Censo_Ano_Excel.as_view(), name='relatorio_censo_ano_excel'),
    path('relatorio_censo_ano/', views.Visualizar_Censo_Ano.as_view(), name='relatorio_censo_ano'),
]


