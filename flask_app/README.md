# Flask App - Sistema de Reservas Hoteleiras

Migração do sistema Django para Flask.

## 🚀 Setup Inicial

### 1. Criar Virtual Environment

```bash
cd /home/edir/Documentos/SCH_V02/flask_app
python3 -m venv .venv-flask
source .venv-flask/bin/activate
```

### 2. Instalar Dependências

```bash
pip install -r requirements.txt
```

### 3. Configurar Ambiente

```bash
cp .env.example .env
# Editar .env se necessário
```

### 4. Rodar Aplicação

```bash
python app.py
```

Acesse: http://localhost:5000

## 📁 Estrutura

```
flask_app/
├── app.py                 # Aplicação principal
├── config.py              # Configurações
├── requirements.txt       # Dependências
├── models/                # SQLAlchemy models
├── routes/                # Blueprints (rotas)
├── services/              # Lógica de negócio
├── forms/                 # WTForms
├── templates/             # Templates Jinja2
└── static/                # CSS, JS, imagens
```

## 🔗 Banco de Dados

Este app Flask usa o **mesmo banco de dados** (`db.sqlite3`) que o Django.

**IMPORTANTE:** Não rode migrations do Flask. O banco já existe e está sendo usado pelo Django.

## ✅ Verificação

### Health Check
```bash
curl http://localhost:5000/health
```

### Testar Conexão com Banco
```bash
python -c "from models import db, BaseDados; print(BaseDados.query.count())"
```

## 📋 Status da Migração

- [x] Fase 0: Setup e Estrutura
- [ ] Fase 1: Models
- [ ] Fase 2: Services
- [ ] Fase 3: Forms
- [ ] Fase 4: Primeira Rota
- [ ] Fase 5: Recepção
- [ ] Fase 6: Autenticação
- [ ] Fase 7: Relatórios
- [ ] Fase 8: Testes

## 🔧 Desenvolvimento

### Rodar em modo debug
```bash
export FLASK_DEBUG=1
python app.py
```

### Rodar testes
```bash
pytest tests/ -v
```

## 📚 Documentação

Ver documentos de planejamento em `/home/edir/.gemini/antigravity/brain/*/`
