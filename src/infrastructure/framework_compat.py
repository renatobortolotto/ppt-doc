from __future__ import annotations

import logging
import os
from typing import Any, Callable


try:
    from OPatomic import Environments, OPAtomic, request

    CORPORATE_FRAMEWORK_AVAILABLE = True
except ModuleNotFoundError:
    CORPORATE_FRAMEWORK_AVAILABLE = False

    class Environments:
        def get(self, key: str, default: Any = None) -> Any:
            return os.getenv(key, default)

    class _LocalRequestProxy:
        path = ""
        headers: dict[str, str] = {}
        files: dict[str, Any] = {}
        form: dict[str, Any] = {}
        args: dict[str, Any] = {}
        _json_payload: Any = None

        def get_json(self, silent: bool = False) -> Any:
            return self._json_payload

    request = _LocalRequestProxy()

    class _LocalApp:
        def __init__(self) -> None:
            self.config: dict[str, Any] = {}
            self.logger = logging.getLogger("ppt-doc.pyweb")
            self.before_request_funcs: list[Callable[..., Any]] = []
            self.registered_routes: list[dict[str, Any]] = []

        def before_request(self, fn: Callable[..., Any]) -> Callable[..., Any]:
            self.before_request_funcs.append(fn)
            return fn

        def run(self, host: str, port: int) -> None:
            raise RuntimeError(
                "OPatomic nao esta disponivel neste ambiente local. "
                "Use o ambiente corporativo para subir a aplicacao PyWeb."
            )

    class OPAtomic:
        def __init__(self, import_name: str, *, prefix: str = "") -> None:
            self.import_name = import_name
            self.prefix = prefix.rstrip("/")
            self._app = _LocalApp()

        def create_app(self) -> _LocalApp:
            return self._app

        def create_route(self, handler: Callable[..., Any], path: str) -> Callable[..., Any]:
            full_path = f"{self.prefix}{path}" if self.prefix else path
            self._app.registered_routes.append(
                {
                    "path": full_path,
                    "handler": handler,
                }
            )
            return handler
