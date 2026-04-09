from __future__ import annotations

try:
    from OPaudit import OPBusinessActionInterface
except ModuleNotFoundError:
    class OPBusinessActionInterface:
        def get_business_action(self, service):
            return "UNDEFINED_ACTION"

PATH = "/build-pptx"

class BusinessActionDefinitionCustom(OPBusinessActionInterface):
    def get_business_action(self, service):
        if PATH in service and "POST" in service:
            return "build_pptx_POST_ACTION"
        return 'UNDEFINED_ACTION'
