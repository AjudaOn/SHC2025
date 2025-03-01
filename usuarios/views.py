from django.http import HttpResponse
from django.shortcuts import redirect, render
from django.contrib.auth.models import User
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required



def cadastro (request):
    if request.method == 'GET':        
        return render(request, 'portal/cadastro.html')
    else:
        nome  = request.POST.get('nome')
        patente = request.POST.get('patente')
        email = request.POST.get('email')
        username = request.POST.get('cpf')               
        senha = request.POST.get('senha')

        user = User.objects.filter(username=username).first()
        
        if user:
            return HttpResponse('Usuário já cadastrado!')
        
        user = User.objects.create_user(first_name=nome, last_name=patente, username=username, email=email, password=senha)
        user.save()

        return HttpResponse('Usuário cadastrado com sucesso!')
        
def custom_login(request):
    if request.method == 'GET':
        return render(request, 'portal/login.html')
    else:
        username = request.POST.get('cpf')
        senha = request.POST.get('senha')

        user = authenticate(username=username, password=senha)

        if user is not None:
            login(request, user)  # Autenticação bem-sucedida, usuário logado
            return render(request, 'portal/home.html')  # Redireciona para a home
        else:
            return HttpResponse('CPF ou Senha incorretos')  # Credenciais inválidas

@login_required
def home(request):
    if request.user.is_superuser:
        return render(request, 'portal/home.html')  # Se for superusuário, renderiza home.html
    else:
        return render(request, 'portal/home2.html')  # Se não for superusuário, renderiza home2.html

  
     
def logout_view(request):
    logout(request)
    # Redirecionar para a página de login após o logout
    return redirect('custom_login')    