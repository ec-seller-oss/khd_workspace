from __future__ import annotations

from fastapi import Request, HTTPException, Depends
from itsdangerous import URLSafeTimedSerializer, BadSignature
from passlib.context import CryptContext
from sqlalchemy.orm import Session
from typing import Optional
import os
from database import get_db
from models import User

SECRET_KEY = os.getenv("SECRET_KEY", "dev-secret-key-change-in-production")
SESSION_COOKIE = "landscanner_session"

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")
serializer = URLSafeTimedSerializer(SECRET_KEY)


def hash_password(password: str) -> str:
    return pwd_context.hash(password)


def verify_password(plain: str, hashed: str) -> bool:
    return pwd_context.verify(plain, hashed)


def create_session_token(user_id: int) -> str:
    return serializer.dumps({"user_id": user_id})


def decode_session_token(token: str, max_age: int = 86400 * 30) -> dict:
    try:
        data = serializer.loads(token, max_age=max_age)
        return data
    except BadSignature:
        return None


def get_current_user_optional(request: Request, db: Session = Depends(get_db)) -> Optional[User]:
    token = request.cookies.get(SESSION_COOKIE)
    if not token:
        return None
    data = decode_session_token(token)
    if not data:
        return None
    user = db.query(User).filter(User.id == data["user_id"]).first()
    return user


def get_current_user(request: Request, db: Session = Depends(get_db)) -> User:
    user = get_current_user_optional(request, db)
    if not user:
        raise HTTPException(status_code=302, headers={"Location": "/login"})
    return user


def require_auth(request: Request, db: Session = Depends(get_db)) -> User:
    """ページルート用: 未ログインはリダイレクト"""
    token = request.cookies.get(SESSION_COOKIE)
    if not token:
        from fastapi.responses import RedirectResponse
        return None  # ページルートで処理
    data = decode_session_token(token)
    if not data:
        return None
    return db.query(User).filter(User.id == data["user_id"]).first()
