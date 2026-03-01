"""
Flask Configuration
Aponta para o mesmo db.sqlite3 usado pelo Django
"""
import os

# Diretório base do projeto Flask
basedir = os.path.abspath(os.path.dirname(__file__))

# Caminho para o db.sqlite3 do Django (um nível acima)
DB_PATH = os.path.join(os.path.dirname(basedir), 'db.sqlite3')


class Config:
    """Configuração base do Flask"""
    
    # Secret key para sessões (alterar em produção)
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'dev-secret-key-change-in-production'
    
    # Conexão com o banco de dados SQLite existente
    SQLALCHEMY_DATABASE_URI = f'sqlite:///{DB_PATH}'
    
    # Desabilita tracking de modificações (economiza memória)
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    
    # Configurações de sessão
    SESSION_COOKIE_SECURE = False  # True em produção com HTTPS
    SESSION_COOKIE_HTTPONLY = True
    SESSION_COOKIE_SAMESITE = 'Lax'
    
    # Configurações de upload (se necessário)
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max
    
    # Timezone
    TIMEZONE = 'America/Sao_Paulo'

    # Expiração automática de reservas
    ENABLE_RESERVA_EXPIRER = True
    RESERVA_EXPIRER_INTERVAL = int(os.environ.get('RESERVA_EXPIRER_INTERVAL', 300))


class DevelopmentConfig(Config):
    """Configuração de desenvolvimento"""
    DEBUG = True
    TESTING = False


class ProductionConfig(Config):
    """Configuração de produção"""
    DEBUG = False
    TESTING = False
    # Em produção, usar variável de ambiente para SECRET_KEY
    SECRET_KEY = os.environ.get('SECRET_KEY')


class TestingConfig(Config):
    """Configuração de testes"""
    TESTING = True
    SQLALCHEMY_DATABASE_URI = 'sqlite:///:memory:'  # Banco em memória para testes


# Configuração padrão
config = {
    'development': DevelopmentConfig,
    'production': ProductionConfig,
    'testing': TestingConfig,
    'default': DevelopmentConfig
}
