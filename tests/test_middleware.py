import types
import unittest
from unittest.mock import patch

from src.application.middleware import jwt_middleware


class TestJwtMiddleware(unittest.TestCase):
    def test_jwt_middleware_skips_options_requests(self):
        request_obj = types.SimpleNamespace(
            method="OPTIONS",
            headers={},
        )

        with patch("src.application.middleware.jwt_mid", object()):
            with patch("src.application.middleware.JWT_SECRET", "secret"):
                with patch("src.application.middleware.JWT_ALGORITHM", "HS256"):
                    result = jwt_middleware(request_obj)

        self.assertIsNone(result)

    def test_jwt_middleware_rejects_post_without_bearer_token(self):
        request_obj = types.SimpleNamespace(
            method="POST",
            headers={},
        )

        with patch("src.application.middleware.jwt_mid", object()):
            with patch("src.application.middleware.JWT_SECRET", "secret"):
                with patch("src.application.middleware.JWT_ALGORITHM", "HS256"):
                    result = jwt_middleware(request_obj)

        self.assertEqual(result, ({"error": "Token de autenticacao ausente ou invalido"}, 401))


if __name__ == "__main__":
    unittest.main()
