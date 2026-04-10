from src.controller.app import op_app
from src.application.build_pptx import BuildPptxResource

def create_routes():
    """Creates the routes for the application."""
    op_app.create_route(BuildPptxResource, "/build-pptx")
