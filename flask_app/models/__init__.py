"""
SQLAlchemy Models
Mapeia as tabelas existentes do Django no db.sqlite3
"""
from flask_sqlalchemy import SQLAlchemy

# Instância do SQLAlchemy
db = SQLAlchemy()

# Importar models quando forem criados
from .base_dados import BaseDados
from .produto import Produto
from .precos import Precos_status_graduacao, Precos_graduacao_vinculo
from .user import User
from .etapas_audit import ReservaEtapaAudit
