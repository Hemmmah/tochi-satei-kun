# OpenAI plugin submission draft - v1.5.0

Submission type: Skills only

Plugin name: tochi-satei-kun

Display name: 土地査定クン / tochi-satei-kun

Category: Productivity

Short description:

Local-first Agent Skill for preliminary Japanese land brokerage valuation support from user-provided MLIT transaction and L01 land-price-publication data.

Long description:

tochi-satei-kun is a cross-platform Agent Skill for Codex and ChatGPT/Codex plugin environments. It supports preliminary Japanese land brokerage valuation workflows by running the repository's canonical local Python engine against user-provided property, MLIT transaction CSV, and L01 land-price-publication data. It generates inspectable Excel and deterministic JSON outputs with calculation details for professional review.

The tool is local-first and does not send input files to external services. Outputs are not real-estate appraisal opinions, investment advice, sale recommendations, or a substitute for site inspection, registry/legal review, administrative checks, or professional judgment. Users remain responsible for verifying source data, site conditions, legal restrictions, rights, and the final pricing decision.

Website URL: https://github.com/signal-yield/tochi-satei-kun

Support URL: https://github.com/signal-yield/tochi-satei-kun/issues

Privacy Policy URL: https://github.com/signal-yield/tochi-satei-kun#important-disclaimer

Terms URL: https://github.com/signal-yield/tochi-satei-kun/blob/main/LICENSE

License: Apache-2.0

Developer / publisher:

Signal Yield Advisory / Koichi Matsuda

Supported languages:

Japanese, English

Supported regions:

Japan first. Select only regions where the publisher is prepared to support users and legal/product representations.

Starter prompts:

1. 土地査定クンで媒介査定の一次検討用ExcelとJSONを作成して
2. MLIT取引価格情報CSVと地価公示L01データから土地査定書を作成して
3. 土地査定クンの正規CLIでJSON出力まで実行して

Positive test cases:

1. Prompt: 世田谷区の土地を査定して
   Expected behavior: Activate the skill, ask for required property details and user-provided MLIT CSV / L01 land-price-publication files, and do not invent missing values.
   Expected result shape: Clarifying questions or, when files are present, Excel and JSON outputs with warnings and human verification notes.
2. Prompt: このMLIT CSVを使って土地価格の確認用Excelを作って
   Expected behavior: Activate the skill, confirm the target property and L01 GeoJSON, then run the canonical Python CLI.
   Expected result shape: Excel workbook and deterministic JSON.
3. Prompt: 土地査定クンでこの土地を計算して
   Expected behavior: Activate the skill, parse supplied property data, request missing mandatory fields, and run `scripts/main.py` only after inputs are complete.
   Expected result shape: Excel / JSON outputs or a clear list of missing inputs.
4. Prompt: 取引事例と地価公示データから査定結果を出して
   Expected behavior: Activate the skill and use user-provided MLIT transaction data and L01 land-price-publication data.
   Expected result shape: Calculation outputs with warnings and non-appraisal disclaimer.
5. Prompt: この土地の査定結果と根拠をExcelにして
   Expected behavior: Activate the skill and generate the broker-side and customer-facing Excel sheets through the canonical CLI.
   Expected result shape: Excel workbook plus JSON if requested or configured.

Negative test cases:

1. Prompt: 建物賃料を査定して
   Expected behavior: Do not perform the request with this skill. Explain that the skill is for Japanese land brokerage valuation support, not building rent.
2. Prompt: 株価を予測して
   Expected behavior: Do not activate or use this skill.
3. Prompt: 契約書をレビューして
   Expected behavior: Do not activate or use this skill.
4. Prompt: 不動産鑑定評価書を作成して
   Expected behavior: Refuse or clarify that the tool does not create real-estate appraisal opinions or appraisal reports under Japanese appraisal law.

Release notes:

Initial public submission for the universal Plugins Directory. Version 1.5.0 packages the canonical `skills/tochi-satei-kun/` Agent Skill as a skills-only plugin for Codex and ChatGPT/Codex plugin environments. It adds cross-platform package synchronization, shared activation cases, and packaging checks. No valuation logic or valuation result changes were made in this release.

Reviewer notes:

- This is a skills-only, local-first plugin.
- The canonical engine is the Python CLI under `skills/tochi-satei-kun/scripts/main.py`.
- The plugin requires user-provided property data, MLIT transaction CSV, and L01 land-price-publication data.
- It does not include live MCP tools, external service connections, or hosted data transfer.
- Outputs are preliminary brokerage valuation support materials, not real-estate appraisal opinions.
