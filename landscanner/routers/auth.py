from fastapi import APIRouter, Depends, Request, Form, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session
from database import get_db
from models import User
from auth import hash_password, verify_password, create_session_token, SESSION_COOKIE
import os

router = APIRouter()
templates = Jinja2Templates(directory="templates")


@router.get("/login", response_class=HTMLResponse)
async def login_page(request: Request):
    return templates.TemplateResponse("login.html", {"request": request})


@router.post("/login")
async def login(
    request: Request,
    email: str = Form(...),
    password: str = Form(...),
    db: Session = Depends(get_db),
):
    user = db.query(User).filter(User.email == email).first()
    if not user or not verify_password(password, user.password_hash):
        return templates.TemplateResponse(
            "login.html",
            {"request": request, "error": "メールアドレスまたはパスワードが違います"},
            status_code=401,
        )
    token = create_session_token(user.id)
    response = RedirectResponse(url="/", status_code=302)
    response.set_cookie(
        key=SESSION_COOKIE,
        value=token,
        httponly=True,
        max_age=86400 * 30,
        samesite="lax",
    )
    return response


@router.post("/logout")
async def logout():
    response = RedirectResponse(url="/login", status_code=302)
    response.delete_cookie(SESSION_COOKIE)
    return response


@router.get("/logout")
async def logout_get():
    response = RedirectResponse(url="/login", status_code=302)
    response.delete_cookie(SESSION_COOKIE)
    return response


@router.get("/register", response_class=HTMLResponse)
async def register_page(request: Request):
    return templates.TemplateResponse("register.html", {"request": request})


@router.post("/register")
async def register(
    request: Request,
    email: str = Form(...),
    password: str = Form(...),
    password_confirm: str = Form(...),
    db: Session = Depends(get_db),
):
    def error(msg: str):
        return templates.TemplateResponse(
            "register.html",
            {"request": request, "error": msg, "email": email},
            status_code=400,
        )

    if len(password) < 8:
        return error("パスワードは8文字以上で設定してください")
    if password != password_confirm:
        return error("パスワードが一致しません")
    if db.query(User).filter(User.email == email).first():
        return error("このメールアドレスはすでに登録されています")

    # 最初のユーザーを管理者にする
    is_first_user = db.query(User).count() == 0
    user = User(
        email=email,
        password_hash=hash_password(password),
        is_admin=is_first_user,
    )
    db.add(user)
    db.commit()
    db.refresh(user)

    token = create_session_token(user.id)
    response = RedirectResponse(url="/settings", status_code=302)
    response.set_cookie(
        key=SESSION_COOKIE,
        value=token,
        httponly=True,
        max_age=86400 * 30,
        samesite="lax",
    )
    return response
