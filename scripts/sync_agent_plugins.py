"""Synchronize the canonical Agent Skill into platform plugin packages."""

from __future__ import annotations

import argparse
import filecmp
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SKILL_NAME = "tochi-satei-kun"
SOURCE = REPO_ROOT / "skills" / SKILL_NAME
IGNORED_NAMES = {"__pycache__", ".DS_Store", ".pytest_cache", "output", "tests", "PR_HANDOFF.md"}
IGNORED_SUFFIXES = {".pyc", ".pyo"}


@dataclass(frozen=True)
class PackageTarget:
    name: str
    destination: Path


TARGETS = (
    PackageTarget("codex", REPO_ROOT / "plugins" / SKILL_NAME / "skills" / SKILL_NAME),
    PackageTarget("claude", REPO_ROOT / ".claude-plugin" / "plugins" / SKILL_NAME / "skills" / SKILL_NAME),
)


def included_files(root: Path) -> set[Path]:
    return {
        path.relative_to(root)
        for path in root.rglob("*")
        if path.is_file()
        and not any(part in IGNORED_NAMES for part in path.relative_to(root).parts)
        and path.suffix not in IGNORED_SUFFIXES
    }


def differences(target: PackageTarget) -> list[str]:
    if not SOURCE.is_dir():
        return [f"canonical skill is missing: {SOURCE}"]
    if not target.destination.is_dir():
        return [f"{target.name} plugin skill is missing: {target.destination}"]

    source_files = included_files(SOURCE)
    destination_files = included_files(target.destination)
    messages = [
        f"{target.name}: missing from package: {path}"
        for path in sorted(source_files - destination_files)
    ]
    messages += [
        f"{target.name}: extra in package: {path}"
        for path in sorted(destination_files - source_files)
    ]
    messages += [
        f"{target.name}: content differs: {path}"
        for path in sorted(source_files & destination_files)
        if not filecmp.cmp(SOURCE / path, target.destination / path, shallow=False)
    ]
    return messages


def sync(target: PackageTarget) -> None:
    if not SOURCE.is_dir():
        raise FileNotFoundError(f"canonical skill is missing: {SOURCE}")
    if target.destination.exists():
        shutil.rmtree(target.destination)
    shutil.copytree(
        SOURCE,
        target.destination,
        ignore=shutil.ignore_patterns(
            "__pycache__",
            "*.pyc",
            "*.pyo",
            ".DS_Store",
            ".pytest_cache",
            "output",
            "tests",
            "PR_HANDOFF.md",
        ),
    )


def selected_targets(names: list[str] | None) -> tuple[PackageTarget, ...]:
    if not names:
        return TARGETS
    selected = {name.lower() for name in names}
    unknown = selected - {target.name for target in TARGETS}
    if unknown:
        raise ValueError(f"unknown target(s): {', '.join(sorted(unknown))}")
    return tuple(target for target in TARGETS if target.name in selected)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="exit non-zero when any package differs")
    parser.add_argument("--target", action="append", choices=[target.name for target in TARGETS])
    args = parser.parse_args()

    targets = selected_targets(args.target)
    if args.check:
        mismatches = [message for target in targets for message in differences(target)]
        if mismatches:
            print("Agent plugin skills are out of sync:", file=sys.stderr)
            for mismatch in mismatches:
                print(f"- {mismatch}", file=sys.stderr)
            return 1
        print("Agent plugin skills are in sync.")
        return 0

    for target in targets:
        sync(target)
        print(f"Synchronized {SOURCE} -> {target.destination}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
