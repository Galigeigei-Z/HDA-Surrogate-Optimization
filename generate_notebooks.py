from __future__ import annotations

import importlib.util
import json
from pathlib import Path


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "notebook_sources" / "supertargeting_demo.py"
OUTPUT = ROOT / "supertargeting_demo.ipynb"


def load_module(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Failed to load module from {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main() -> None:
    if OUTPUT.exists() and OUTPUT.stat().st_mtime > SOURCE.stat().st_mtime:
        raise RuntimeError(
            f"Notebook {OUTPUT.name} is newer than its source. Sync notebook edits back into {SOURCE.name} before regeneration."
        )

    module = load_module(SOURCE, "supertargeting_demo_source")
    notebook = module.build_notebook()
    OUTPUT.write_text(json.dumps(notebook, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"Generated: {OUTPUT}")


if __name__ == "__main__":
    main()
