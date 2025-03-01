from django.urls import path
from . import views

urlpatterns = [
    path('cadastro/', views.cadastro, name='cadastro'),
    path('login/', views.custom_login, name='custom_login'),
    path('home/', views.home, name='home'),    
    path('logout/', views.logout_view, name='logout'),
]