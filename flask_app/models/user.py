from . import db
from flask_login import UserMixin
from passlib.hash import django_pbkdf2_sha256

class User(db.Model, UserMixin):
    __tablename__ = 'auth_user'
    
    id = db.Column(db.Integer, primary_key=True)
    password = db.Column(db.String(128), nullable=False)
    last_login = db.Column(db.DateTime)
    is_superuser = db.Column(db.Boolean, nullable=False)
    username = db.Column(db.String(150), unique=True, nullable=False)
    first_name = db.Column(db.String(150), nullable=False)
    last_name = db.Column(db.String(150), nullable=False)
    email = db.Column(db.String(254), nullable=False)
    is_staff = db.Column(db.Boolean, nullable=False)
    is_active = db.Column(db.Boolean, nullable=False)
    date_joined = db.Column(db.DateTime, nullable=False)

    def check_password(self, password):
        """Verifica a senha usando o hash do Django"""
        try:
            return django_pbkdf2_sha256.verify(password, self.password)
        except Exception:
            return False

    def set_password(self, password):
        self.password = django_pbkdf2_sha256.hash(password)

    def get_id(self):
        return str(self.id)
