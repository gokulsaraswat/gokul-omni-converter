"""Headless tests for archive_core and rename_organizer_core.

These two modules are pure domain logic with no Tk dependency, so they are the
part of the app that can be tested properly. Run directly:

    python tests/test_cores.py
"""

from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import archive_core as A  # noqa: E402
import rename_organizer_core as R  # noqa: E402
from converter_core import ConversionError  # noqa: E402

FAILURES: list[str] = []


def check(label: str, condition: bool, detail: str = "") -> None:
    print(f"  {'ok  ' if condition else 'FAIL'} {label}{'  ' + detail if detail else ''}")
    if not condition:
        FAILURES.append(label)


def make_tree(root: Path) -> Path:
    src = root / "src"
    (src / "sub").mkdir(parents=True)
    for index, name in enumerate(["a.txt", "b.txt", "photo.jpg", "sheet.csv"]):
        (src / name).write_text("x" * (100 * (index + 1)), encoding="utf-8")
    (src / "sub" / "deep.md").write_text("deep", encoding="utf-8")
    (src / ".DS_Store").write_text("junk", encoding="utf-8")  # must be skipped
    return src


def test_archive(tmp: Path) -> None:
    print("archive_core")
    src = make_tree(tmp)

    bundle = A.create_archive([src], tmp / "bundle.zip")
    summary = A.read_archive(bundle)
    names = [entry.name for entry in summary.entries]
    check("skips OS junk files", not any(".DS_Store" in n for n in names))
    check("archives every real file", summary.file_count == 5, f"({summary.file_count})")

    dest = A.extract_archive(bundle, tmp / "extracted")
    check("round-trips content", (dest / "src" / "sub" / "deep.md").read_text() == "deep")

    outputs = A.batch_compress([src, src / "sub"], tmp / "batch")
    check("batch makes one zip per folder", len(outputs) == 2 and all(p.exists() for p in outputs))

    try:
        A.read_archive(src / "a.txt")
        check("rejects non-zip input", False)
    except ConversionError:
        check("rejects non-zip input", True)

    check("suggests a sensible name", A.suggest_archive_name([src / "a.txt"]) == "a.zip")


def test_rename(tmp: Path) -> None:
    print("rename_organizer_core")
    src = make_tree(tmp / "rename")
    state = tmp / "state"
    files = sorted(src.glob("*.txt")) + [src / "photo.jpg"]

    plan = R.plan_rename(files, mode="pattern", pattern="doc_{n}_{name}", number_padding=3)
    check("pattern rename plans every file", len(plan.applicable) == 3, f"({len(plan.applicable)})")

    # Two .txt files collapsing to one name is a collision; the .jpg keeps its own extension.
    collide = R.plan_rename(files, mode="pattern", pattern="same")
    check("detects name collisions", len(collide.problems) == 1, f"({len(collide.problems)})")

    try:
        R.plan_rename(files, mode="regex", regex_pattern="([")
        check("rejects invalid regex", False)
    except ConversionError:
        check("rejects invalid regex", True)

    applied, manifest = R.apply_plan(plan, state)
    check("applies the rename", (src / "doc_001_a.txt").exists())
    check("writes an undo manifest", manifest.exists() and len(applied) == 3)

    restored = R.undo_last(state)
    check("undo restores originals", (src / "a.txt").exists() and restored == 3)

    organize = R.plan_organize(src, mode="type")
    R.apply_plan(organize, state)
    check("organises by file type", (src / "Documents" / "a.txt").exists() and (src / "Images" / "photo.jpg").exists())

    R.undo_last(state)
    check("undo reverses organise", (src / "a.txt").exists())

    try:
        R.undo_last(state)
        check("refuses undo with no history", False)
    except ConversionError:
        check("refuses undo with no history", True)


def main() -> int:
    tmp = Path(tempfile.mkdtemp())
    try:
        test_archive(tmp)
        test_rename(tmp)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    if FAILURES:
        print(f"\n{len(FAILURES)} FAILED: {', '.join(FAILURES)}")
        return 1
    print("\nAll core tests passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
