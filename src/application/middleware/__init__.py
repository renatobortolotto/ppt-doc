from __future__ import annotations

from base64 import b64encode
import logging
import os

from src.infrastructure.framework_compat import request

try:
    import jwt as jwt_mid
except ModuleNotFoundError:
    jwt_mid = None


JWT_SECRET = os.getenv("JWT_SECRET_KEY")
JWT_ALGORITHM = os.getenv("JWT_ALGORITHM")
logger = logging.getLogger(__name__)


def jwt_middleware(request_obj=None):
    if jwt_mid is None:
        logger.info("Biblioteca jwt indisponivel localmente; validacao JWT sera ignorada")
        return None

    if not JWT_SECRET or not JWT_ALGORITHM:
        logger.warning("JWT nao configurado; validacao sera ignorada neste ambiente")
        return None

    active_request = request_obj or request
    auth_header = (getattr(active_request, "headers", None) or {}).get("Authorization")
    if not auth_header or not auth_header.startswith("Bearer "):
        return {"error": "Token de autenticacao ausente ou invalido"}, 401

    token = auth_header.split(" ")[1]

    try:
        secret = b64encode(JWT_SECRET.encode("utf-8"))
        decoded = jwt_mid.decode(token, secret, algorithms=[JWT_ALGORITHM])
        setattr(active_request, "jwt_payload", decoded)
    except jwt_mid.InvalidTokenError as e:
        return {"error": "Token de autenticacao invalido", "details": str(e)}, 401
    return None
