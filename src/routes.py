from src.controller.app import app, op_app
from src.application.build_pptx import BuildPptxResource
from src.infrastructure.config import prefix


def _full_route_path(path: str) -> str:
    normalized_prefix = prefix.rstrip("/")
    return f"{normalized_prefix}{path}" if normalized_prefix else path


def _build_pptx_route_registered() -> bool:
    full_path = _full_route_path("/build-pptx")

    registered_routes = getattr(app, "registered_routes", None)
    if isinstance(registered_routes, list):
        return any(
            route.get("path") == full_path and route.get("handler") is BuildPptxResource
            for route in registered_routes
            if isinstance(route, dict)
        )

    url_map = getattr(app, "url_map", None)
    view_functions = getattr(app, "view_functions", None)
    if url_map is None or view_functions is None:
        return False

    for rule in url_map.iter_rules():
        if rule.rule != full_path:
            continue
        view = view_functions.get(rule.endpoint)
        if getattr(view, "view_class", None) is BuildPptxResource:
            return True
    return False


def create_routes():
    """Creates the routes for the application."""
    if _build_pptx_route_registered():
        return
    op_app.create_route(BuildPptxResource, "/build-pptx")
