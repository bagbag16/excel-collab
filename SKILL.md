---
name: excel-collab
description: Use when Codex needs to read, inspect, compare, modify, write, validate, or explain Excel workbooks (.xlsx, .xlsm, or workbook-derived CSVs) in human-AI collaboration, especially when a workbook is both source data, generated output, review surface, or versioned state. Enforce active workbook/sheet selection, data-contract clarity, version lineage, formula-vs-value handling, user-edit protection, open-file/write-lock checks, and post-write validation.
---

# Excel Collab

Use this skill to keep Excel work safe and recoverable when a workbook carries data, logic, output, and user intent at the same time.

Core stance: Excel is a review and delivery surface unless explicitly stated otherwise. Do not treat it as the only state root when scripts, CSVs, Markdown notes, or formal state files define the current mouth.

## Progressive Use

Do not expand every rule for every task. First classify the user request into behavior tags and object tags, then apply only the sections attached to those tags.

## Tagging Model

Every Excel task starts by assigning a compact tag set. Tags are internal routing labels; do not ask the user to choose them.

Use Occam's razor: create a tag only when it changes at least one of these:

- execution order
- implementation/tool choice
- safety precheck
- validation/readback method
- user-edit protection policy

Do not create tags for ordinary parameters such as exact sheet name, exact range, desired value, row count, source path, or output name. Put those in the active workbook contract.

Behavior tags describe the action, in user-request order:

- `read`: inspect, explain, compare, or audit without writing.
- `remove`: delete rows, columns, sheets, values, formulas, names, tables, or generated artifacts.
- `add`: insert or create rows, columns, sheets, formulas, notes, generated outputs, or source data.
- `overwrite`: replace existing cells/ranges/sheets with new values or formulas.
- `derive`: compute a new table/output from source data.
- `style`: change presentation only.
- `calculate`: require Excel's calculation engine, macro execution, external-link refresh, or pivot/chart refresh.
- `repair`: recover a corrupt workbook, broken OOXML, broken formulas, or a bad edit.
- `toolize`: create or improve a deterministic helper for repeated Excel work.

Validation is not usually a behavior tag; it is a required phase selected from the action tags. Use a separate validation/audit step only when the user explicitly asks for verification without another action.

Object tags describe the minimum object conditions that change implementation or risk:

- Storage: `local` or `cloud`.
- Format: `xlsx`, `xlsm`, or `grid-text` for CSV/TSV.
- Role: `source`, `working`, `generated`, `published`, or `state`.
- Surface: `cells`, `sheet`, `workbook-structure`, `presentation`, or `ooxml`.
- Risk flags only when relevant: `formulas`, `macros`, `external-links`, `calc-chain`, `hidden-filtered`, `merged-cells`, `protected`, `large-file`, `locked`, `non-ascii`, `human-edited`.

Use behavior tags to choose action-specific cautions. Use object tags to choose implementation and validation. When multiple behavior tags appear, preserve order, for example `remove -> add -> validate-by-readback`.

Tag-driven expansion rules:

- `read` on an unfamiliar workbook: use `Operating Model`, `Active Workbook Contract`, and `Reading Workflow`.
- `read` + `formulas`: use `Drift Guards` and formula-vs-cache checks.
- `remove`: require exact target surface, backup unless disposable, adjacent sentinel validation, and checks for formulas, tables, names, dimensions, calc-chain, and human-edited ranges that may reference the removed area.
- `add`: require source of inserted content, placement semantics, formula-fill/format policy, uniqueness/order expectations, and validation of row/column counts plus sample cells.
- `overwrite`: require overwrite policy, formula-vs-value policy, human-edit protection, and before/after sampled verification.
- `derive`: keep source data separate from generated output unless explicitly told to overwrite; preserve a machine-readable companion when useful.
- `style`: protect underlying values/formulas; validate visible presentation only where relevant.
- `calculate`: use Excel COM or cloud Excel services when Excel's calculation engine is required; do not trust stale caches.
- `repair`: inspect package/XML invariants first, preserve or restore from backup, and avoid broad rewrites until the damaged part is known.
- `toolize`: use `Toolbook Growth`; design deterministic scripts instead of rewriting ad hoc code.
- `local` + `xlsx`/`xlsm` + `ooxml`: use `Implementation Selection`, `OOXML Safety`, `Edit Workflow`, and `Validation Workflow`.
- `cloud`: use the Graph/Office Scripts branch in `Implementation Selection`.
- `non-ascii`: use `Windows Unicode Hygiene` before running shell or inline scripts.
- `xlsm` or `macros`: preserve VBA parts and avoid tools that silently drop macros.

## Operating Model

Before reading or editing, classify each workbook or sheet:

- `source`: original data; preserve unless the user explicitly asks to edit it.
- `working`: intermediate analysis or scratch calculations; may be regenerated.
- `generated`: machine-produced output; update only from declared inputs.
- `published`: user-facing delivery sheet; preserve presentation and manual comments.
- `state`: a workbook/sheet the user treats as current truth; require stronger readback and validation.

If roles are unclear, infer from context and state the assumption before making durable edits.

## Active Workbook Contract

For any non-trivial Excel task, establish a compact contract:

- workbook path and whether it is the current active file
- target sheet name(s), not just workbook name
- sheet role: source, working, generated, published, or state
- source of truth: workbook cells, CSV, script output, external data, or user instruction
- key field meanings and units, especially time window, per-hit, per-round, per-second, totals, and final-vs-theory values
- formula policy: preserve formulas, overwrite formulas, write values only, or export formulas separately
- human-edit policy: columns/ranges/notes that must not be overwritten
- write mode: dry-run, sidecar output, append sheet, replace generated sheet, or edit in place

Do not ask the user to fill this contract mechanically. Build it from files and context; ask only for missing choices that would be risky to assume.

## Reading Workflow

When inspecting a workbook:

1. List sheets first when the active sheet is not explicitly known.
2. Inspect dimensions, headers, row counts, and modified time before interpreting values.
3. Identify hidden rows/columns, merged cells, filters, tables, named ranges, formulas, and data validations if they may affect the task.
4. Distinguish formulas from cached values. Do not trust cached formula values as recalculated unless Excel or a calculation engine has refreshed them.
5. Normalize field meanings before calculations. Same-looking names such as `damage`, `total`, `round`, or `value` may refer to different layers.
6. Prefer exporting bounded sheet ranges to CSV for analysis; keep the workbook as the review surface.

## Speed Policy

Separate first-pass discovery from follow-up execution:

- On the first tool-design or unfamiliar workbook request, inspect structure enough to identify sheets, ranges, formulas, and data roles.
- On follow-up edits to the same workbook/sheet where the user gives a clear target, reuse the already established contract and execute directly.
- For simple row/column deletion, tail truncation, bounded value replacement, or generated-sheet refresh, do not repeat broad workbook analysis. Validate only the touched sheet, affected row/column counts, and sentinel cells.
- Keep user-facing work proportional. If a human could do the edit in a minute, the automated path should be a short deterministic edit plus readback, not a full exploratory pass.

## Implementation Selection

Choose the fastest stable implementation that preserves the required workbook surface:

- Use direct OOXML zip/XML edits for local `.xlsx` / `.xlsm` structural operations that are simple and bounded, especially tail-row deletion, removing stale calc chains, changing worksheet dimensions, or editing known XML attributes. This avoids Excel startup and dependency overhead.
- Use `openpyxl` or a similar workbook library for ordinary cell/range edits when preserving workbook structure is sufficient and formulas do not need Excel recalculation during the task.
- Use streaming or optimized read/write modes for very large generated workbooks or CSV-like transformations.
- Use Microsoft Graph Excel APIs or Office Scripts when the workbook lives in OneDrive/SharePoint and the task benefits from Excel's service-side calculation/session model.
- Use Excel COM only when Excel itself is required: macro execution, external-link refresh, live recalculation that no library can provide, charts/pivots that must be updated by Excel, or visual/manual automation. Treat COM as slower and more failure-prone; set a short timeout and clean up orphaned Excel processes if it hangs. For the local workspace recalculation helper, prefer `tools\recalculate_excel_workbook.cmd` over calling the `.ps1` directly so PowerShell execution policy is bypassed consistently, and validate workbook paths before starting Excel COM.
- If Excel COM reports a sharing violation, save conflict, or locked file after
  one controlled retry, stop and return a read-only diagnosis or ask the user to
  close the workbook. Do not keep trying multiple write paths against the same
  locked workbook and expand the damage surface.
  不要在共享违规后连续尝试多种写入方式。
- After deleting formula-heavy regions by OOXML, remove `xl/calcChain.xml` and its workbook/content-type references, and set workbook calc properties such as `fullCalcOnLoad` / `forceFullCalc` when formulas remain.

## OOXML Safety

Treat `.xlsx` / `.xlsm` as Open XML packages, not as generic XML documents. Excel may reject a workbook even when every XML part is well-formed.

- Avoid full-tree parse-and-serialize rewrites of Excel XML with generic serializers such as Python `xml.etree.ElementTree` unless namespace prefixes, markup-compatibility attributes, and extension namespaces are explicitly preserved.
- Preserve original namespace prefixes such as `mc`, `x14ac`, `xr`, `xr2`, and `xr3` when a part contains `mc:Ignorable`. The value of `mc:Ignorable` is prefix text; generic XML serializers may rename declarations to `ns1/ns2` without updating the attribute value, producing an Excel repair dialog.
- Prefer minimal text/streaming patches for simple bounded structural edits: remove known `<row ...>` blocks, adjust `dimension ref`, clamp visible selection references, or remove specific calc-chain entries without reserializing unrelated XML.
- If a serializer must be used, inspect the root element before and after. Verify that namespace prefixes referenced by `mc:Ignorable` are still declared with the same names.
- Validation must include Excel-specific invariants, not only zip integrity and XML parse success.

## Toolbook Growth

Treat this skill as the AI-facing Excel modification toolbook. The model should decide intent, scope, and safety checks; deterministic scripts should perform repeated fragile edits.

Promote a workflow into `scripts/` when it recurs, is easy to parameterize, or has failure modes that should not be re-solved from scratch. Good candidates include sheet inspection, bounded range export, tail-row truncation, row/column deletion, formula-cache and calc-chain cleanup, generated-sheet refresh, workbook diffing, lock detection, and post-write validation.

For future tools:

- Keep inputs explicit: workbook path, sheet name or index, range/row bounds, write mode, backup tag, and dry-run flag.
- Emit JSON summaries with path, sheet, changed ranges, backup path, and validation results.
- Avoid requiring non-ASCII inline source. Accept Unicode paths and sheet names as command arguments, and keep the tool source ASCII-safe where practical.
- Prefer idempotent operations and dry-runs for destructive edits.
- Keep human-facing reports compact; detailed mechanics belong in tool output or logs.

## Edit Workflow

Before writing:

- Create a timestamped backup unless the edit is a disposable generated copy.
- Check whether the workbook appears open or locked. If locked, ask the user to close it or write a sidecar/copy.
- Avoid overwriting source sheets or human-editable ranges.
- Prefer adding or refreshing clearly named generated sheets over modifying raw data sheets.
- If replacing a generated sheet, record the producing inputs and script/config in a note sheet or adjacent artifact.
- If using scripts, make outputs deterministic and re-runnable.
- For row-copy/insert tasks, identify the business maintenance predicate before
  choosing source rows, such as a text marker in column B plus a numeric group
  key. Insert directly below the maintained source row/block when formulas or
  formatting are copied from adjacent rows. Do not select by the numeric key
  alone or append below the last global matching key unless that is the explicit
  business rule.
  不要只凭数字列匹配决定维护范围或插入位置。
- Before changing a formatted config value, compare the same column's existing
  values and adjacent source row. Preserve separators, prefixes, suffixes, and
  text shape; change only the intended numeric part. Do not replace established
  symbols such as `@` with `~` without explicit instruction.
  不要在只修改数值时破坏同列旧值的符号和文本格式。
- If the user manually edited the workbook after an automated plan, discard
  stale planned row numbers and rediscover the current workbook surface before
  validating or writing again. Do not validate a user-edited workbook only
  against an old `new_rows` list.
  不要拿旧自动化计划校验用户手工提交后的当前表。

When editing manually, preserve workbook structure unless the user asked for restructuring. For fragile formatting, use a workbook library that preserves existing sheets where possible, or write a separate generated workbook.

For known simple edits, use this short path:

1. Reconfirm the target workbook, sheet, and exact row/column/range from cached context or a bounded read.
2. Create a timestamped backup.
3. Apply the smallest deterministic edit.
4. Reopen or re-read the workbook and verify only the changed surface plus adjacent sentinels.
5. Report path, sheet, edit range, backup, and validation.

## Windows Unicode Hygiene

On Windows PowerShell, non-ASCII paths, sheet names, or inline Python source can be corrupted by console/pipeline encodings. Avoid making decisions from the rendered PowerShell output when Chinese text matters; verify via UTF-8 file reads or ASCII-safe JSON output instead.

- Before running any inline script that mentions a non-ASCII workbook path or sheet name, perform this as a preflight, not as a fallback after failure.
- Prefer `scripts\utf8_text_probe.py` for Chinese text/file checks. Pass Chinese
  search terms as Unicode escapes, for example `--contains-escape
  "\u56db\u65e5"`, and inspect its ASCII-safe JSON instead of the terminal's
  rendered Chinese.
- Prefer resolving paths by filesystem enumeration plus ASCII filters, or use Unicode escape literals such as `\u7b49\u7ea7\u6a21\u677f` inside Python source.
- Do not pipe here-string Python source containing Chinese or other non-ASCII text to `python -` unless the command first sets `$OutputEncoding`, `[Console]::InputEncoding`, and `[Console]::OutputEncoding` to UTF-8 and runs Python with `-X utf8`.
- Prefer ASCII-only inline scripts that compare discovered workbook/sheet names, or store non-ASCII names as Unicode escapes.
- For readable command output, set console output to UTF-8 or emit ASCII-safe JSON with escaped non-ASCII.
- If a path or sheet name appears as `????` or mojibake, stop and fix encoding before writing; do not proceed with a guessed target.
- For Chinese path, sheet, or cell-text predicates, prefer JSON arguments,
  workbook-discovered strings, or Unicode escapes such as `\u56db\u65e5`.
  Do not hard-code Chinese literals in a PowerShell here-string or inline Python
  script and then trust the result without an encoding preflight.
  不要在 PowerShell inline 脚本里直接硬写中文路径或中文匹配常量后继续相信结果。
- Do not explain away repeated mojibake after the fact; change the command shape
  so the check never depends on the PowerShell rendering layer.
  不要反复用“PowerShell 输出层编码偏差”解释问题，要改用 UTF-8 读取和 ASCII-safe JSON 校验。

## Validation Workflow

Writing the workbook is not completion. Reopen or re-read the file and verify:

- target sheet exists with the intended name
- row count and header count match expected outputs
- key cells or sampled rows match the source calculation
- formulas were preserved or intentionally replaced according to the contract
- copied formulas came from the intended adjacent/source row, not from an unrelated global match elsewhere in the sheet
- formatted config values keep the same separator/text shape as same-column existing values unless the user explicitly requested a format change
- no protected human ranges or source sheets were overwritten
- backup path exists if a backup was required
- generated outputs are linked to the source CSV/script/config in the answer

For generated analysis sheets, also produce or preserve a machine-readable CSV/JSON/Markdown companion when useful for future recovery.

## Drift Guards

Use these checks whenever the user says "latest", "current", "this sheet", "the table", "write it in", "why is this value weird", or "continue from the workbook":

- Verify the current sheet, not just the current workbook.
- Verify whether the workbook reflects the latest script/CSV output.
- Verify whether old experimental sheets or archived files are being mistaken for active results.
- Verify whether visible Excel formatting hides rows, columns, or filters.
- Verify whether a value is per-hit, per-skill, per-round, per-cycle, per-second, or full-window.
- Verify whether a formula cell was read as a cached value.
- Verify whether the user edited the workbook after the last generated output.

If a conflict exists, report it as a state conflict instead of silently choosing one side.

## Communication Standard

When reporting Excel work, include:

- workbook path
- sheet name(s)
- whether the workbook was read-only inspected, copied, or edited
- what source data produced the result
- what validation was performed
- remaining risks, such as formulas not recalculated, workbook open/locked, or old sheets still present

Keep the user-facing answer compact. Do not expose low-level XML or library mechanics unless the user asks.

## Common Failure Patterns

Avoid these:

- treating an old sheet, old workbook copy, archived output, or experiment tab as current
- mixing source data, generated output, and manual review notes in one overwrite operation
- using field names without confirming their unit and layer
- recalculating from Excel display formatting instead of underlying values
- assuming formula caches are fresh
- losing user edits by regenerating a sheet in place
- claiming success after file write without re-reading the workbook
- giving the user an output without path and sheet name
- maintaining rows outside the declared business predicate just because a numeric key matches
- changing configuration syntax while only intending to update numeric values
- retrying Excel COM writes repeatedly after a lock/share violation
- 不要维护业务谓词之外的行，即使数字 key 看起来匹配。
- 不要只想改数值却改坏配置符号。
- 不要在 Excel 锁表或共享违规后反复写入。
