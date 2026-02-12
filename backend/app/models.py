from sqlalchemy import Column, Integer, String, DateTime, Text, Boolean
from datetime import datetime
from app.database import Base


class Auth(Base):
    __tablename__ = "auth"

    id = Column(Integer, primary_key=True, index=True)
    email = Column(String(255), unique=True, index=True, nullable=False)
    access_token = Column(Text, nullable=False)  # Fernet encrypted
    refresh_token = Column(Text, nullable=False)  # Fernet encrypted
    expires_at = Column(DateTime, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


class BackgroundJob(Base):
    __tablename__ = "background_jobs"

    id = Column(String(36), primary_key=True)  # UUID
    job_type = Column(String(50), nullable=False)
    status = Column(String(20), default="pending")  # pending, running, complete, error
    progress = Column(Integer, default=0)
    total = Column(Integer, default=0)
    result = Column(Text, nullable=True)  # JSON
    error = Column(Text, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


class ApiKey(Base):
    __tablename__ = "api_keys"

    id = Column(Integer, primary_key=True, index=True)
    key_hash = Column(String(64), unique=True, index=True, nullable=False)  # SHA256 hex
    name = Column(String(255), unique=True, nullable=False)
    permissions = Column(Text, nullable=False)  # JSON array, e.g. '["read:mail","read:calendar"]'
    created_at = Column(DateTime, default=datetime.utcnow)
    last_used_at = Column(DateTime, nullable=True)
    is_active = Column(Boolean, default=True, nullable=False)
