from src.infrastructure.framework_compat import OPAtomic, request

from src.infrastructure.config import prefix
from src.infrastructure.config.business_action import (
    BusinessActionDefinitionCustom,
)
from src.application.middleware import jwt_middleware

op_app = OPAtomic(__name__, prefix=prefix)
app = op_app.create_app()
app.config.update(
    {
        "business_action_providers": [BusinessActionDefinitionCustom()]
    }
)

@app.before_request
def before_request():
    protected_path = f"{prefix.rstrip('/')}/build-pptx" if prefix else "/build-pptx"
    if getattr(request, "path", "") == protected_path:
        result = jwt_middleware()
        if result:
            return result
        return
    return
