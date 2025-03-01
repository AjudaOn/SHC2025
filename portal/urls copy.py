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
]

