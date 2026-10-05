# AGENTS.md

## Project source of truth

- GitHub `main` is the canonical source of truth for VBA and documentation.
- Files under `サンプルデータ` are reference data for headers/layout/verification and are not implementation targets unless explicitly requested.
- Do not treat a local workbook copy as the canonical implementation when it differs from `main`.

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
