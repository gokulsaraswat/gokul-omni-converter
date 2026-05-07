from __future__ import annotations

import json
from pathlib import Path

from release_validation import run_release_checks


def main() -> int:
    root = Path(__file__).resolve().parent
    result = run_release_checks(root)
    print(json.dumps(result, indent=2))
    if result.get("status") != "passed":
        return 1
    print("Final release validation passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
