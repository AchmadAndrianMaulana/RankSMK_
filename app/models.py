from datetime import datetime
from . import db, login_manager
from flask_login import UserMixin

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

class User(db.Model, UserMixin):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class Upload(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    original_filename = db.Column(db.String(255), nullable=False)
    saved_path = db.Column(db.String(255), nullable=False)  # raw file path
    processed_path = db.Column(db.String(255), nullable=True)  # processed CSV
    n_rows = db.Column(db.Integer, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class Criteria(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    ctype = db.Column(db.String(20), nullable=False)  # 'benefit' or 'cost'
    display_order = db.Column(db.Integer, default=1)  # used by ROC

class Subcriteria(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    criteria_id = db.Column(db.Integer, db.ForeignKey('criteria.id'), nullable=False)
    name = db.Column(db.String(200), nullable=False)
    min_val = db.Column(db.Float, nullable=True)
