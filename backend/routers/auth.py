"""Authentication endpoints."""
import bcrypt
from fastapi import APIRouter
from pydantic import BaseModel
from backend.db import get_db
from backend.auth import create_token


def hash_password(plain: str) -> str:
    return bcrypt.hashpw(plain.encode(), bcrypt.gensalt()).decode()


def check_password(plain: str, hashed: str) -> bool:
    # Support both bcrypt hashes and legacy plain-text passwords
    if hashed.startswith("$2"):
        return bcrypt.checkpw(plain.encode(), hashed.encode())
    return plain == hashed  # legacy: plain-text match

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
    if not check_password(body.password, user[0]["password"]):
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
