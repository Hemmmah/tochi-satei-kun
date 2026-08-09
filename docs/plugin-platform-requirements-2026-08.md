# Agent Skill / Plugin platform requirements as of 2026-08-09

Primary sources checked:

- OpenAI Developers, Package your plugin: https://developers.openai.com/plugins/build/plugins
- OpenAI Developers, Build skills: https://developers.openai.com/plugins/build/skills
- OpenAI / ChatGPT Learn, Plugins overview: https://learn.chatgpt.com/docs/plugins
- OpenAI Developers, Plugin submission errors: https://developers.openai.com/plugins/deploy/submission-errors
- Anthropic Claude Code Docs, Skills: https://code.claude.com/docs/en/skills
- Anthropic Claude Code Docs, Plugin marketplaces: https://code.claude.com/docs/en/plugin-marketplaces

| 項目 | 共通 | OpenAI固有 | Claude固有 |
| -- | -- | -------- | -------- |
| Skill正本 | `skills/tochi-satei-kun/SKILL.md` を唯一の手編集対象にする。 | Plugin内には同期生成した `skills/tochi-satei-kun/SKILL.md` を含める。 | Plugin内には同期生成した `skills/tochi-satei-kun/SKILL.md` を含める。 |
| Skill仕様 | `SKILL.md` はfrontmatterと本文で構成し、作業手順、必須入力、禁止事項、実行CLIを明示する。 | Skillはplugin内の `skills/` 配下に配置できる。 | SkillはClaude Codeが関連時または `/skill-name` で利用する。 |
| Plugin manifest | version, name, description, author, license等をSemVer更新と合わせて管理する。 | `.codex-plugin/plugin.json` がmanifest。skills-only pluginでは `skills` が `./skills/` を指す。 | `.claude-plugin/plugin.json` がmanifest。Claude marketplaceのplugin source先に置く。 |
| Marketplace / directory | 公開説明は、鑑定評価・投資判断ではなく媒介査定支援であることを明記する。 | ChatGPT / CodexはUniversal Plugin Directoryを使う。公開提出ではmanifest、skill、listing fields、review material等の追加チェックがある。 | Claude Codeは `.claude-plugin/marketplace.json` をGit等で配布するmarketplace方式が公式文書化されている。 |
| Skills-only package | 外部MCPやconnectorを使わないlocal-first packageにする。 | skills-only ZIPではMCP/app/screenshot設定を含めない。 | Claude pluginにはskillsを含め、marketplaceからinstallできる構成にする。 |
| Installation | Python依存関係はSkillの `requirements.txt` を使う。 | Codex desktop / CLIのPlugins UIまたはlocal marketplaceからinstallして新規taskでテストする。 | `claude plugin marketplace add` と `claude plugin install` でmarketplaceからinstallしてテストする。 |
| Testing | canonical Skillとpackaged Skillの内容一致、manifest JSON、version一致、activation cases、valuation regressionをCIで検証する。 | `scripts/sync_agent_plugins.py --check` でCodex package一致を検証する。 | `scripts/sync_agent_plugins.py --check` でClaude package一致を検証する。 |
| Legal / privacy | 入力データを外部送信しないlocal-first実装であること、出力が鑑定評価ではないこと、最終判断が利用者に残ることを明記する。 | 公開提出時はsupport URL、privacy policy、terms、developer verification等が必要になる可能性がある。 | 公式marketplace配布ではowner、description、source等を明示する。公開registry制度が追加されている場合はAnthropic公式手順に従う。 |
| Versioning | engine version、Codex manifest、Claude manifestを一致させる。既存tagは上書きしない。 | 新規公開・更新ではmanifest versionを更新する。 | version指定時は更新ごとにversion bumpする。 |

## Current package decision

This repository remains local-first and skills-only. Codex-specific and Claude-specific files are thin wrappers around the canonical Skill:

- Canonical Skill: `skills/tochi-satei-kun/`
- Codex package: `plugins/tochi-satei-kun/`
- Claude marketplace/package: `.claude-plugin/`

Synchronize both packages with:

```bash
python scripts/sync_agent_plugins.py
python scripts/sync_agent_plugins.py --check
```

The legacy Codex-only command remains as a compatibility wrapper:

```bash
python scripts/sync_codex_plugin_skill.py --check
```
