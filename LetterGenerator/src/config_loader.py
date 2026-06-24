import json
from pathlib import Path


def load_template_config(config_path: Path) -> dict:
    with open(config_path, encoding="utf-8") as f:
        return json.load(f)


def get_template_dir(config_path: Path) -> Path:
    return config_path.parent
