from . import db

class Precos_status_graduacao(db.Model):
    __tablename__ = 'precos_por_status_e_graduacao'
    
    id = db.Column(db.Integer, primary_key=True)
    status = db.Column(db.String(100), nullable=False)
    graduacao = db.Column(db.String(100), nullable=False)
    valor = db.Column(db.Numeric(10, 2), nullable=False)

class Precos_graduacao_vinculo(db.Model):
    __tablename__ = 'precos_por_graduacao_e_vinculo'
    
    id = db.Column(db.Integer, primary_key=True)
    graduacao = db.Column(db.String(100), nullable=False)
    vinculo = db.Column(db.String(100), nullable=False)
    valor = db.Column(db.Numeric(10, 2), nullable=False)
