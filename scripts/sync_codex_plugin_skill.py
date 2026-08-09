"""Compatibility wrapper for synchronizing the Codex plugin package."""

from __future__ import annotations

import sys

from sync_agent_plugins import main


if __name__ == "__main__":
    sys.argv.extend(["--target", "codex"])
    raise SystemExit(main())
