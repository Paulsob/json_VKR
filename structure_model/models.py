from .extensions import db
from datetime import datetime


class Absence(db.Model):
    __tablename__ = "absences"
    id = db.Column(db.Integer, primary_key=True)
    tab_no = db.Column(db.String(50), nullable=False, index=True)
    shift = db.Column(db.Integer, nullable=False)
    day = db.Column(db.Integer, nullable=False)
    reason = db.Column(db.String(255))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class Route(db.Model):
    __tablename__ = "routes"
    id = db.Column(db.Integer, primary_key=True)
    route_number = db.Column(db.Integer, nullable=False, unique=True, index=True)
    name = db.Column(db.String(255))
    sheet_name_workday = db.Column(db.String(255), nullable=False)
    sheet_name_weekend = db.Column(db.String(255), nullable=False)
    file_workday = db.Column(db.String(500))
    file_weekend = db.Column(db.String(500))
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
