# Changelog

## v1.5.0 - 2026-08-09

- Canonicalized `skills/tochi-satei-kun/` as the single Agent Skill source for Codex and Claude Code.
- Added `scripts/sync_agent_plugins.py` to synchronize both Codex and Claude plugin packages from the canonical Skill.
- Kept `scripts/sync_codex_plugin_skill.py` as a Codex-only compatibility wrapper.
- Added Claude Code plugin marketplace/package files under `.claude-plugin/`.
- Added shared activation test cases and packaging checks for both platforms.
- Updated README and platform requirements documentation for cross-platform installation and public submission preparation.
- No valuation logic or valuation result changes.
