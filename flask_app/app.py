"""
Aplicação Flask Principal
Sistema de Reservas Hoteleiras - Migração do Django
"""
from flask import Flask
from config import config
import os
import threading
import time


def create_app(config_name=None):
    """
    Factory function para criar a aplicação Flask
    
    Args:
        config_name: Nome da configuração ('development', 'production', 'testing')
    
    Returns:
        app: Instância configurada do Flask
    """
    if config_name is None:
        config_name = os.environ.get('FLASK_ENV', 'development')
    
    app = Flask(__name__)
    app.config.from_object(config[config_name])
    
    # Inicializar extensões
    from models import db
    from flask_wtf.csrf import CSRFProtect
    from flask_login import LoginManager
    
    db.init_app(app)
    csrf = CSRFProtect(app)
    
    login_manager = LoginManager()
    login_manager.login_view = 'auth.login'
    login_manager.login_message = 'Por favor, faça login para acessar esta página.'
    login_manager.login_message_category = 'info'
    login_manager.init_app(app)
    
    from models.user import User
    @login_manager.user_loader
    def load_user(user_id):
        return User.query.get(int(user_id))
    
    # Registrar blueprints (rotas)
    from routes import reservas, recepcao, auth, relatorios, usuarios, produtos
    app.register_blueprint(reservas.bp)
    app.register_blueprint(recepcao.bp)
    app.register_blueprint(auth.bp)
    app.register_blueprint(relatorios.bp)
    app.register_blueprint(usuarios.bp)
    app.register_blueprint(produtos.bp)

    # Criar tabelas novas adicionadas pelo Flask (não altera tabelas existentes).
    # Usado para recursos novos (ex.: auditoria de etapas).
    try:
        with app.app_context():
            db.create_all()
    except Exception:
        app.logger.exception('Falha ao executar db.create_all()')

    def start_reserva_expirer():
        if app.config.get('TESTING'):
            return
        if not app.config.get('ENABLE_RESERVA_EXPIRER', True):
            return

        interval = app.config.get('RESERVA_EXPIRER_INTERVAL', 300)
        from services.reservas import expire_reservas

        def worker():
            while True:
                with app.app_context():
                    try:
                        expire_reservas()
                    except Exception:
                        app.logger.exception('Falha ao expirar reservas')
                time.sleep(interval)

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()

    start_reserva_expirer()
    
    # Rota de teste inicial
    @app.route('/health')
    def health_check():
        """Endpoint de health check"""
        return {
            'status': 'ok',
            'message': 'Flask app is running',
            'database': 'connected' if db.engine else 'not connected'
        }
    
    return app


if __name__ == '__main__':
    app = create_app()
    app.run(host='0.0.0.0', port=5000, debug=True)
