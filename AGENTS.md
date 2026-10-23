# AGENTS.md

## Project source of truth

- GitHub `main` is the canonical source of truth for VBA and documentation.
- Files under `サンプルデータ` are reference data for headers/layout/verification and are not implementation targets unless explicitly requested.
- Do not treat a local workbook copy as the canonical implementation when it differs from `main`.
- Every managed VBA component has one authoritative paste-ready source at `vba/<component name>.txt`: standard module, class module, UserForm code, or worksheet code. The sheet file `vba/新ファイル基準表.txt` uses the sheet tab name; its workplace VBE codename has not been confirmed. Each source begins with `Option Explicit`, omits export-only `Attribute VB_Name`, and is CP932-representable. Copy the complete file to the corresponding existing VBE code window; do not import `.txt` or add duplicate components.
- For future functional changes, edit the relevant `vba/*.txt` authoritative source directly, not only a workbook, export, restore artifact, fixture, or documentation. Before commit confirm the source appears in the diff and run relevant tests, body comparison, and CP932 checks; after push read back `origin/main` and confirm the committed files. Report changed filenames, VBE destinations, and reasons.
- Keep `frmSeal` (old), `frmSealNew`, and `frmSyncReview` as separate forms. Do not alter their designer, controls, or event behavior for a source-location change. Root `frmSyncReview.frm` is a reference/restore artifact and `frmSyncReview-native.frm` plus `frmSyncReview-native.frx` are an adjacent import/restore pair, not alternative paste-ready authorities. Preserve them; when changing the canonical form code later, verify embedded code in restore artifacts remains aligned.
- `Module2` exists in the user's Excel but has no verified repository source. It remains outside GitHub management: do not infer its contents, fabricate a source, or rename/delete it.

## Project-specific Excel integration verification

- Follow `safe-app-development` and `excel-vba-safe-development` for common model, time/attempt ledger, retry, and stop rules; keep only this project's source mapping and workbook side effects here.
- This repository's sync mapping is `FILES` in `サンプルデータ/スキル用/sync_github_vba.py`; follow `excel-vba-sync` for invocation and source-snapshot procedure. Ensure a local snapshot includes every source in `FILES`. The helper's `SOURCE` is a template, not automatically the designated data-bearing original; use the source workbook explicitly identified for the work item.
- The current VBA has external side effects: `Module1` creates recovery copies in `%LOCALAPPDATA%\file-index-sync\CSVImportBackups`; `modExcelRowId` uses `%LOCALAPPDATA%\file-index-sync\ExcelRowIdBackups`; CSV import stages files under `%TEMP%\CSVImportStage_*.csv`; `固定印刷設定` stores `土木課C4476R.dat` under `%LOCALAPPDATA%\file-index-sync\固定印刷設定`. Isolate `LOCALAPPDATA` and `TEMP` for applicable tests and verify them before/after; if a path cannot be isolated, do not run the macro.
- For generation changes, inspect relevant results in `新ファイル基準表`, `コード管理CSV`, and `差分インポートCSV`; for CSV-import changes, inspect `目録CSV` and `ログ`. When matching/year behavior is affected, choose 3–5 relevant cases such as exact-ID match, unique-name match, new entry, ambiguous conflict, or an era/year pair; compare expected and actual before/operation/after values. A value visible in an InputBox is not proof of the value received by VBA.

## Cross-PC development ownership

### Default ownership and assignment

- Normal development from implementation through Excel runtime verification is owned by the laptop's Hermes profile `hermes-dev`, where Microsoft Excel is installed.
- Desktop Hermes `default` may take development or split work only when the user explicitly assigns it. Do not self-assign, change PC ownership, or start/control the laptop Hermes on the user's behalf.
- Before desktop work, check shared Git/handoff information for overlapping notebook work. If concurrent ownership cannot be established, ask the user before editing.

### Desktop split-work scope

- On explicit assignment, desktop may change only the requested source, run static/regression checks, and prepare the laptop/user handoff. Commit/push to GitHub `main` only when the task explicitly includes main reflection.
- Desktop has no Microsoft Excel. Do not attempt Excel COM, VBE compilation, or Excel runtime macro verification. Static success is not Excel verification and must never be reported as feature-complete or passed.
- An explicit split-work task that includes main reflection may push source after static checks, but must label it `Excel実機検証待ち`. The laptop Hermes or user performs Excel verification. If it fails, fix with a new commit and use a safe revert if necessary.
- Do not create verification branches by default. GitHub `main` is the sole code source. Fetch/check latest `origin/main` before work and again before push. Preserve dirty and untracked files. On conflicts or unclear divergence, stop; never force-push or hard-reset.

### Handoff

Record the target commit SHA; implemented function; static/regression test commands and results; Excel runtime verification status; 3–5 concrete Excel checks; remaining tasks and known issues; autonomous start time, elapsed time, and verification-attempt counts.

### Autonomous-development limits

- Follow common safe-development and Excel-safety Skills when available. Do not bypass them. If a required rule is unavailable in the active profile, stop and ask when the limits below do not resolve the safety question.
- Start in Luna; escalate to Sol only if root-cause analysis requires it. Do not force-stop solely because 3 wall-clock hours have passed since autonomous start; perform a safety review after each 3 cumulative hours of active work, excluding approval/user-response waits and Telegram interruptions. Review progress, uncommitted changes, verification results, and risks; if clear, continue without renewed approval.
- Limit terminal real-world/integration verification to 3 attempts per stage. Use CUA/GUI only if necessary, at most 2 attempts. Do not reset cumulative active-work time or attempt counts on PC switch, `/goal` continuation, interruptions, or session resume.
- Do not repeat a failed attempt for the same cause without a new hypothesis. Stop for serious risk, an important unresolved specification decision, or an unsafe result.
- Keep Excel runtime verification explicitly pending until it is actually performed on the laptop or by the user.

## Operational manual synchronization rule

The end-user operational manual is:

- `新ファイル基準表.md`

Whenever a code change affects any user-visible workflow, operational rule, status, judgment, warning, sheet structure, button/macro usage, color meaning, or the relationship between Excel and the management system, update `新ファイル基準表.md` in the same work item.

Examples that require a manual update:

- Step1–Step4 behavior or execution order changes
- addition / update / deletion workflows change
- sync-status semantics change
- diff matching or registration behavior changes in a way users need to understand
- guide-seat creation or classification-code numbering changes
- `コード管理CSV` columns or states change
- yellow / red / gray display meanings change
- new conflict, warning, abort, or recovery behavior is added
- users need to perform a new manual step

If the code change is purely internal and does not alter user-visible behavior, a manual update is not required.

## Documentation roles

- `README.md`: developer/project overview
- `新ファイル基準表.md`: end-user operational manual and current business workflow
- `README_固定印刷設定.md`: fixed-printing-specific manual

## Documentation quality

When updating the operational manual:

- describe both Excel-side and management-system-side workflows where applicable
- explain "if you do X, Y happens"
- keep addition, update, and deletion easy to find
- keep status/color meanings current
- update Mermaid flowcharts when the workflow changes
- avoid documenting behavior that the current code does not actually support
- explicitly note system-side assumptions that cannot be verified from the exported CSV
