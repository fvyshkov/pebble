"""Authentication endpoints."""
from fastapi import APIRouter
from pydantic import BaseModel
from backend.db import get_db
from backend.auth import create_token

router = APIRouter(prefix="/api/auth", tags=["auth"])


class LoginIn(BaseModel):
    username: str
    password: str


@router.post("/login")
async def login(body: LoginIn):
    db = get_db()
    user = await db.execute_fetchall(
        "SELECT id, username, password, can_admin FROM users WHERE username = ?",
        (body.username,),
    )
    if not user:
        return {"error": "Неверный логин или пароль"}
    if user[0]["password"] != body.password:
        return {"error": "Неверный логин или пароль"}

    token = create_token(user[0]["id"], user[0]["username"])
    return {
        "token": token,
        "user": {
            "id": user[0]["id"],
            "username": user[0]["username"],
            "can_admin": bool(user[0]["can_admin"]),
        },
    }


@router.get("/me")
async def get_me():
    """Placeholder — token verification is in middleware."""
    return {"ok": True}
