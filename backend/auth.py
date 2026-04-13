"""JWT authentication for Pebble."""
import jwt
import os
from datetime import datetime, timedelta
from fastapi import Request, HTTPException
from starlette.middleware.base import BaseHTTPMiddleware

SECRET = os.environ.get("PEBBLE_JWT_SECRET", "pebble-dev-secret-change-in-prod")
ALGORITHM = "HS256"
TOKEN_EXPIRE_HOURS = 24 * 7  # 1 week

# Paths that don't require auth
PUBLIC_PATHS = {"/api/auth/login", "/api/auth/me"}


def create_token(user_id: str, username: str) -> str:
    payload = {
        "user_id": user_id,
        "username": username,
        "exp": datetime.utcnow() + timedelta(hours=TOKEN_EXPIRE_HOURS),
    }
    return jwt.encode(payload, SECRET, algorithm=ALGORITHM)


def verify_token(token: str) -> dict:
    try:
        return jwt.decode(token, SECRET, algorithms=[ALGORITHM])
    except jwt.ExpiredSignatureError:
        raise HTTPException(status_code=401, detail="Token expired")
    except jwt.InvalidTokenError:
        raise HTTPException(status_code=401, detail="Invalid token")


class AuthMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        path = request.url.path

        # Skip auth for public paths, static, and OPTIONS
        if (path in PUBLIC_PATHS or
            not path.startswith("/api/") or
            request.method == "OPTIONS"):
            return await call_next(request)

        # Check Authorization header
        auth = request.headers.get("Authorization", "")
        if auth.startswith("Bearer "):
            token = auth[7:]
            try:
                payload = verify_token(token)
                request.state.user_id = payload["user_id"]
                request.state.username = payload["username"]
            except HTTPException:
                pass  # Allow request but without user context

        return await call_next(request)
