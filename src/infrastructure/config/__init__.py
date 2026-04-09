from __future__ import annotations

import os


prefix = str(os.getenv("SERVER_ROOT", "/api"))
host = str(os.getenv("HOST", "127.0.0.1"))
port = int(os.getenv("PORT", "8080"))
