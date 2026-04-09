from src.controller.app import op_app
from src.application.build_pptx import build_pptx_route

def create_routes():
    """Creates the routes for the application."""
    op_app.create_route(build_pptx_route, "/build-pptx")
