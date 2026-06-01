from __future__ import annotations

from pathlib import Path


def default_output_path() -> Path:
    desktop = Path.home() / "Desktop"
    output_dir = desktop if desktop.exists() else Path.home()
    return output_dir / "extracted_discharge_data.xlsx"
