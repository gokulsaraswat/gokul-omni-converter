from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from packaging_core import create_portable_source_bundle


def main() -> None:
    parser = argparse.ArgumentParser(description="Create a source bundle ZIP for Gokul Omni Convert Lite")
    parser.add_argument("destination", nargs="?", default="release_output", help="Destination directory for the ZIP bundle")
    parser.add_argument("--name", default="", help="Optional release archive base name")
    args = parser.parse_args()

    output = create_portable_source_bundle(PROJECT_ROOT, Path(args.destination), release_name=args.name)
    print(output)


if __name__ == "__main__":
    main()
