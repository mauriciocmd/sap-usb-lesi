import json
from typing import Dict, Any

class CommandDTO:
    def __init__(self, comando: str, variables: Dict[str, Any], modulo: str = "unknown"):
        self.comando = comando
        self.variables = variables
        self.modulo = modulo

    def to_dict(self) -> Dict[str, Any]:
        return {
            "comando": self.comando,
            "variables": self.variables,
            "modulo": self.modulo
        }
    
    def to_json(self) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False)

    def __repr__(self):
        return self.to_json()