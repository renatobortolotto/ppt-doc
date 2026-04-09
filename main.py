from src.controller.app import app
from src.infrastructure.config import host, port
from src.routes import create_routes

create_routes()

if __name__ == "__main__":
    app.logger.info(f"Starting application on {host}:{port}...")
    app.run(host=host, port=port)
