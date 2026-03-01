# Mapeamento do Sistema (Flask) — SCH_V02

Este documento é o **mapa único** do sistema Flask dentro de `SCH_V02/flask_app/`.
Use quando precisar entender rapidamente **menus**, **rotas**, **templates**, **banco de dados**, **modelos**, **serviços** e **pontos de deploy/operação**.

---

## 1) Onde fica o quê (visão rápida)

- App Flask: `SCH_V02/flask_app/`
- Banco SQLite (compartilhado com Django): `SCH_V02/db.sqlite3`
  - O Flask aponta para esse arquivo via `flask_app/config.py` (variável `DB_PATH`).
- Entrada do Flask (factory + blueprints): `flask_app/app.py`
- Configs (DB, cookies, expiração automática): `flask_app/config.py`
- Constantes (choices, UH por hotel, status): `flask_app/constants.py`
- Models (SQLAlchemy, mapeando tabelas existentes): `flask_app/models/`
- Rotas/blueprints: `flask_app/routes/`
- Forms (Flask-WTF/WTForms): `flask_app/forms/`
- Serviços (regras de negócio): `flask_app/services/`
- Templates (Jinja2): `flask_app/templates/`
- Static (CSS/JS/imagens): `flask_app/static/`

---

## 2) Como rodar (dev)

### Rodar o Flask
Dentro de `SCH_V02/flask_app`:

```bash
python app.py
```

Endpoints úteis:
- Health check: `GET /health` (definido em `flask_app/app.py`)
- Login: `GET /login/`
- App público (reserva externa): `GET /` (definido em `flask_app/routes/reservas.py`)

### Variáveis de ambiente relevantes
- `SECRET_KEY`: chave de sessão (produção)
- `FLASK_ENV`: `development` | `production` | `testing` (lida por `flask_app/app.py`)
- `RESERVA_EXPIRER_INTERVAL`: intervalo (segundos) do expirer de reservas (default 300)

---

## 3) Banco de Dados (SQLite)

### Caminho do banco
- **Arquivo**: `SCH_V02/db.sqlite3`
- **Config**: `flask_app/config.py` define `DB_PATH = ../db.sqlite3` e usa em `SQLALCHEMY_DATABASE_URI`.

Observação:
- Este banco é **compartilhado** com o sistema Django do repositório (não criar/rodar migrations do Flask).

### Tabelas principais usadas no Flask
- `base_dados` (model: `flask_app/models/base_dados.py::BaseDados`)
  - Reserva/estadia, hóspedes, acompanhantes, status, UH/MHEx, consumação e valores.
- `auth_user` (model: `flask_app/models/user.py::User`)
  - Usuários e senha no formato do Django (hash PBKDF2 do Django).
- `produto` (model: `flask_app/models/produto.py::Produto`)
  - Itens de consumo (Água, Refrigerante, Cerveja, Pet).
- `precos_por_status_e_graduacao` (model: `flask_app/models/precos.py::Precos_status_graduacao`)
- `precos_por_graduacao_e_vinculo` (model: `flask_app/models/precos.py::Precos_graduacao_vinculo`)

---

## 4) Autenticação e perfis

### Como funciona
- Login via `Flask-Login` em `flask_app/app.py`.
- Model de usuário: `flask_app/models/user.py` lê `auth_user` do banco.
- Validação de senha: `passlib.hash.django_pbkdf2_sha256` (mesmo padrão do Django).

Observação de navegação pós-login:
- Em `flask_app/routes/auth.py`, usuário **não-admin** redireciona para `recepcao.ocupacao_hoje` (URL `/recepcao/ocupacao/hoje/`), apesar do menu “Usuário comum” apontar para `recepcao.ocupacao_hoje_geral` (URL `/recepcao/ocupacao/hoje/geral/`).

### Perfis (menu e permissões)
O layout define `is_admin = is_superuser OR is_staff` em `flask_app/templates/layouts/sch_base.html`.

- **Admin** (is_superuser ou is_staff):
  - Vê menus: Reservas (análise), Recepção HTM 01/02 (hoje/semanal/mensal), Relatórios, Configuração, Editar reservas.
- **Usuário comum**:
  - Vê apenas: Recepção → Ocupação Hoje (geral).

---

## 5) Menus do sistema (sidebar) e para onde apontam

Fonte de verdade do menu: `flask_app/templates/layouts/sch_base.html`.

### Menu (Admin)

| Menu | Item | Endpoint (url_for) | URL |
|---|---|---|---|
| RESERVA | Analisar | `recepcao.consultar_reservas` | `/reservas/consultar/` |
| RECEPÇÃO - HTM 01 | Ocupação - Hoje | `recepcao.ocupacao_hoje` | `/recepcao/ocupacao/hoje/` |
| RECEPÇÃO - HTM 01 | Ocupação - Semanal | `recepcao.ocupacao_semanal` | `/recepcao/ocupacao/semanal/` |
| RECEPÇÃO - HTM 01 | Ocupação - Mensal | `recepcao.ocupacao_mensal` | `/recepcao/ocupacao/mensal/` |
| RECEPÇÃO - HTM 02 | Ocupação - Hoje | `recepcao.ocupacao_hoje_htm02` | `/recepcao/htm02/ocupacao/hoje/` |
| RECEPÇÃO - HTM 02 | Ocupação - Semanal | `recepcao.ocupacao_semanal_htm02` | `/recepcao/htm02/ocupacao/semanal/` |
| RECEPÇÃO - HTM 02 | Ocupação - Mensal | `recepcao.ocupacao_mensal_htm02` | `/recepcao/htm02/ocupacao/mensal/` |
| RELATÓRIOS | Em Dinheiro | `relatorios.relatorio_pagamento_dinheiro` | `/relatorio/pagamento/dinheiro/` |
| RELATÓRIOS | Em Pix | `relatorios.relatorio_pagamento_pix` | `/relatorio/pagamento/pix/` |
| RELATÓRIOS | Quartos - Financeiro | `relatorios.relatorio_ocupacao_financeiro` | `/relatorio/ocupacao/financeiro/` |
| RELATÓRIOS | Hóspedes e Valores | `relatorios.relatorio_hospedes_valores` | `/relatorio/censo/hospedes-valores/` |
| RELATÓRIOS | UHs Ocupadas | `relatorios.relatorio_uhs_ocupadas` | `/relatorio/censo/uhs-ocupadas/` |
| CONFIGURAÇÃO | Consultar Usuário | `usuarios.consultar_usuarios` | `/configuracao/usuarios/` |
| CONFIGURAÇÃO | Consultar Produto | `produtos.consultar_produtos` | `/configuracao/produtos/` |
| CONFIGURAÇÃO | Editar reservas | `recepcao.consultar_editar_reservas` | `/reservas/editar/` |

### Menu (Usuário comum)
| Menu | Item | Endpoint | URL |
|---|---|---|---|
| RECEPÇÃO | Ocupação - Hoje | `recepcao.ocupacao_hoje_geral` | `/recepcao/ocupacao/hoje/geral/` |

---

## 6) Rotas e telas (blueprints)

### 6.1) Blueprint: `auth` (login/sessão)
Arquivo: `flask_app/routes/auth.py`

| Endpoint | URL | Métodos | Template |
|---|---|---|---|
| `auth.login` | `/login/` | GET/POST | `auth/login.html` |
| `auth.logout` | `/logout/` | GET | (redirect) |
| `auth.alterar_senha` | `/alterar-senha/` | GET/POST | `auth/alterar_senha.html` |

### 6.2) Blueprint: `reservas` (público)
Arquivo: `flask_app/routes/reservas.py`

| Endpoint | URL | Métodos | Template |
|---|---|---|---|
| `reservas.reserva_externa` | `/` | GET/POST | `portal/index.html` |

Notas:
- Reserva externa fixa `mhex = HTM_01` e `status_reserva = Pendente`.
- Retorno do POST é JSON (AJAX).

### 6.3) Blueprint: `recepcao` (fluxos internos)
Arquivo: `flask_app/routes/recepcao.py`

| Endpoint | URL | Métodos | Template/Obs |
|---|---|---|---|
| `recepcao.api_available_uhs` | `/api/available-uhs/` | GET | JSON (`available_uhs`) |
| `recepcao.api_hospede` | `/api/hospede` | GET | JSON (busca por CPF) |
| `recepcao.consultar_reservas` | `/reservas/consultar/` | GET | `portal/reservas.html` (admin) |
| `recepcao.expirar_reservas_manual` | `/reservas/expirar/` | POST | expira reservas manualmente |
| `recepcao.consultar_editar_reservas` | `/reservas/editar/` | GET | `portal/editar_reservas_consulta.html` |
| `recepcao.nova_reserva` | `/reservas/nova/` | GET/POST | `portal/reserva_interna.html` |
| `recepcao.editar_reservas` | `/reservas/editar/<id>/` | GET/POST | `portal/reserva_edit.html` |
| `recepcao.editar_reservas_admin` | `/reservas/editar/admin/<id>/` | GET/POST | `portal/editar_reserva_full.html` |
| `recepcao.editar_saida_pet` | `/ocupacao/editar/<id>/` | GET/POST | `portal/editar_saida_pet.html` (edição rápida: saída) |
| `recepcao.recepcao_checkin` | `/recepcao/checkin/` | GET | `portal/recepcao_checkin.html` |
| `recepcao.recepcao_checkout` | `/recepcao/checkout/` | GET | `portal/recepcao_checkout.html` |
| `recepcao.editar_checkin` | `/checkin/<id>/` | GET/POST | `portal/checkin.html` |
| `recepcao.editar_checkout` | `/checkout/<id>/` | GET/POST | `portal/checkout.html` |
| `recepcao.editar_consumacao` | `/consumacao/<id>/` | GET/POST | `portal/consumacao.html` |
| `recepcao.ocupacao_hoje` | `/recepcao/ocupacao/hoje/` | GET | `portal/ocupacao_hoje.html` (HTM_01) |
| `recepcao.ocupacao_semanal` | `/recepcao/ocupacao/semanal/` | GET | `portal/ocupacao_semanal.html` (HTM_01) |
| `recepcao.ocupacao_mensal` | `/recepcao/ocupacao/mensal/` | GET | `portal/ocupacao_mensal.html` (HTM_01) |
| `recepcao.ocupacao_hoje_htm02` | `/recepcao/htm02/ocupacao/hoje/` | GET | `portal/ocupacao_hoje.html` (HTM_02) |
| `recepcao.ocupacao_semanal_htm02` | `/recepcao/htm02/ocupacao/semanal/` | GET | `portal/ocupacao_semanal.html` (HTM_02) |
| `recepcao.ocupacao_mensal_htm02` | `/recepcao/htm02/ocupacao/mensal/` | GET | `portal/ocupacao_mensal.html` (HTM_02) |
| `recepcao.ocupacao_hoje_geral` | `/recepcao/ocupacao/hoje/geral/` | GET | `portal/ocupacao_hoje.html` (default HTM_01 + `allow_mhex_select=True`) |

#### Edição rápida (lápis) na Ocupação Hoje

Objetivo: permitir que o atendente ajuste **somente a data de saída** sem precisar abrir o checkout completo.

- Onde aparece: `portal/ocupacao_hoje.html` → quando `room.status == 'occupied'`, aparece um ícone de lápis ao lado do botão **Check-out**.
- Para onde vai: `GET/POST /ocupacao/editar/<id>/` (`recepcao.editar_saida_pet`) com `next` apontando de volta para a ocupação.
- O que edita:
  - Apenas `SAÍDA` (e o sistema recalcula `DIARIAS` e `VALOR_TOTAL`).
- Validações:
  - `Saída` precisa ser **maior** que `Entrada`.
  - Se a UH ficar indisponível no período, a alteração é bloqueada (reaproveita a lógica de disponibilidade).

#### Busca de CPF e cadastro de cliente novo (Nova Reserva)

Tela: `portal/reserva_interna.html` (rota `recepcao.nova_reserva` — `/reservas/nova/`)

- **Seção "BUSCAR HÓSPEDE"**: campo CPF + botão "Buscar" que consulta `GET /api/hospede?cpf=...`.
  - Se encontrado, preenche automaticamente todos os dados do hóspede (nome, e-mail, telefone, graduação, status, tipo, sexo, cidade, UF e CPF).
- **Seção "DADOS DO HÓSPEDE PRINCIPAL"**: contém um campo CPF visível (`cpf_hospede`) para que o atendente possa digitar manualmente o CPF de **clientes novos** (sem cadastro prévio).
- **Sincronização bidirecional** (JS no template):
  - Campo de busca → campo de dados do hóspede (e vice-versa).
  - O campo real do formulário WTForms é `form.cpf` (id `cpf`), que tem `DataRequired`. O campo `cpf_hospede` é visual e sincroniza via JS com o `cpf`.

### 6.4) Blueprint: `relatorios`
Arquivo: `flask_app/routes/relatorios.py`

| Endpoint | URL | Template/Obs |
|---|---|---|
| `relatorios.relatorio_pagamento` | `/relatorio/pagamento/` | `portal/relatorio_pagamento.html` |
| `relatorios.relatorio_mensal` | `/relatorio/mensal/` | `portal/relatorio_mensal.html` |
| `relatorios.relatorio_pagamento_pix` | `/relatorio/pagamento/pix/` | `portal/relatorio_pagamento_pix.html` |
| `relatorios.relatorio_pagamento_dinheiro` | `/relatorio/pagamento/dinheiro/` | `portal/relatorio_pagamento_dinheiro.html` |
| `relatorios.relatorio_ocupacao_financeiro` | `/relatorio/ocupacao/financeiro/` | `portal/ocupacao_financeira.html` |
| `relatorios.relatorio_hospedes_valores` | `/relatorio/censo/hospedes-valores/` | `portal/relatorio_hospedes_valores.html` |
| `relatorios.relatorio_hospedes_valores_excel` | `/relatorio/censo/hospedes-valores/excel/<hotel>/` | download Excel |
| `relatorios.relatorio_uhs_ocupadas` | `/relatorio/censo/uhs-ocupadas/` | `portal/relatorio_uhs_ocupadas.html` |
| `relatorios.relatorio_uhs_ocupadas_excel` | `/relatorio/censo/uhs-ocupadas/excel/` | download Excel |
| `relatorios.gerar_relatorio_excel` | `/relatorio/gerar/<forma_pagamento>/` | download Excel |

Serviço de relatórios:
- `flask_app/services/reports.py` (classe `ReportService`)

### 6.5) Blueprint: `usuarios`
Arquivo: `flask_app/routes/usuarios.py`

| Endpoint | URL | Template |
|---|---|---|
| `usuarios.consultar_usuarios` | `/configuracao/usuarios/` | `portal/consultar_usuario.html` |
| `usuarios.cadastrar_usuario` | `/configuracao/usuarios/novo/` | `portal/cadastrar_usuario.html` |
| `usuarios.editar_usuario` | `/configuracao/usuarios/<id>/editar/` | `portal/cadastrar_usuario.html` |
| `usuarios.resetar_senha_usuario` | `/configuracao/usuarios/<id>/resetar-senha/` | (redirect) |

### 6.6) Blueprint: `produtos`
Arquivo: `flask_app/routes/produtos.py`

| Endpoint | URL | Template |
|---|---|---|
| `produtos.consultar_produtos` | `/configuracao/produtos/` | `portal/consultar_produto.html` |
| `produtos.cadastrar_produto` | `/configuracao/produtos/novo/` | `portal/cadastrar_produto.html` |
| `produtos.editar_produto` | `/configuracao/produtos/<id>/editar/` | `portal/cadastrar_produto.html` |

---

## 7) Regras de negócio (serviços)

### Disponibilidade de UH (quartos vagos por período)
- Implementação: `flask_app/services/availability.py::get_available_uhs(mhex, entrada, saida, exclude_id=None)`
- Usado em:
  - Cadastro interno: `recepcao.nova_reserva` (template `portal/reserva_interna.html`)
  - Edição: `recepcao.editar_reservas` e `recepcao.editar_reservas_admin`
  - API: `GET /api/available-uhs/`

Front-end:
- Script: `flask_app/static/js/availability.js`
- Incluído em:
  - `flask_app/templates/portal/reserva_edit.html`
  - `flask_app/templates/portal/editar_reserva_full.html`

### Cálculo financeiro (diárias + consumo + descontos)
- Implementação: `flask_app/services/pricing.py::calculate_total(reserva)`
- Usa:
  - Tabelas de preço (`precos_*`)
  - Produtos (`produto`: Água/Refrigerante/Cerveja/Pet)
- Chamado em vários pontos de `flask_app/routes/recepcao.py` e `flask_app/routes/reservas.py`.

### Expiração automática de reservas
Objetivo: marcar reservas como **Expirada** quando `ENTRADA < hoje` e status ainda bloqueante.

- Serviço: `flask_app/services/reservas.py::expire_reservas(...)`
- Execução automática: thread em background criada por `flask_app/app.py` (config `RESERVA_EXPIRER_INTERVAL`)
- Execução manual: `POST /reservas/expirar/` (`recepcao.expirar_reservas_manual`)

---

## 8) Constantes e “domínio”

Arquivo: `flask_app/constants.py`

- `MHEx_CHOICES`: `HTM_01` e `HTM_02`
- UHs por hotel:
  - `HTM_01`: 12 quartos (1..12)
  - `HTM_02`: 3 quartos (1..3) — ver `UH_CHOICES_BY_MHEX`
- Status de reserva: `Pendente`, `Aprovada`, `Checkin`, `Pago`, `Recusada`, `Expirada`, ...
- Produtos: `Água`, `Refrigerante`, `Cerveja`, `Pet`

---

## 9) Static e layout

### CSS
- Layout principal: `flask_app/static/css/sch-layout.css` (incluído em `layouts/sch_base.html`)

### JS
- Disponibilidade de UHs: `flask_app/static/js/availability.js`
- Outros (legado/demo): `flask_app/static/js/principal.js`, `popup.js`, etc.

### Layout base (shell + sidebar)
- `flask_app/templates/layouts/sch_base.html`

---

## 10) Deploy (referência rápida)

Guia existente: `SCH_V02/migracao_flask.txt` (na raiz do repo).

Pontos principais:
- Subir `flask_app/` + `db.sqlite3` na raiz do projeto no servidor.
- Ajustar `DB_PATH` em `flask_app/config.py` se o banco não estiver no mesmo nível.
- Definir `SECRET_KEY` e `FLASK_ENV=production`.
- Garantir dependência `passlib` instalada (necessária para senha do Django).

---

## 11) Checklist rápido (dev/ops)

- Banco existe e está no lugar certo? `SCH_V02/db.sqlite3`
- Usuário consegue logar? `GET /login/`
- Menu “Admin” aparece para usuários `is_staff`/`is_superuser`?
- API de UHs funciona? `GET /api/available-uhs/?mhex=HTM_01&entrada=YYYY-MM-DD&saida=YYYY-MM-DD`
- Expiração automática ativa? ver `ENABLE_RESERVA_EXPIRER` e logs do app.

### Checklist pós-deploy (produção)

1) `git pull` no servidor (PythonAnywhere) e Web → **Reload**
2) Validar login (`/login/`)
3) Validar ocupação:
   - HTM 01: `/recepcao/ocupacao/hoje/`
   - HTM 02: `/recepcao/htm02/ocupacao/hoje/`
4) Validar edição rápida (lápis):
   - em um quarto **ocupado**, clicar no lápis → alterar `Saída` → salvar → voltar para a ocupação e confirmar as datas
5) Validar checkout completo:
   - clicar em **Check-out** → salvar pagamento → confirmar status `Pago` nos relatórios/pesquisa
