import json
from typing import Dict, Any
from ..styling.models import StylingConfigModel

def load_config(config_path: str) -> Dict[str, Any]:
    """Loads the main configuration from a JSON file."""
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def load_styling_config(sheet_config: Dict[str, Any]) -> StylingConfigModel:
    """Parses the styling portion of the config into Pydantic models."""
    print(f"DEBUG: sheet_config in load_styling_config: {sheet_config.get('styling', {})}")
    return StylingConfigModel(**sheet_config.get('styling', {}))
