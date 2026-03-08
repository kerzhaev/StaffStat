# Completion Summary: tech debt refactor wave 1

**Date:** 2026-03-09
**Scope:** safe schema initialization, search debug throttling, service/UI decoupling, import diagnostics, import profile self-heal, import wizard UI extraction, reports decoupling, mapping alignment, linked backend schema self-heal

## Why this refactor started

The project accumulated several technical debt clusters that were blocking safe production use and later refactors:
- startup initialization could recreate buffer structures destructively
- search debug logging was too noisy for normal runtime
- service modules mixed business logic with `MsgBox` / `InputBox`
- `Full Update` hid the real import failure reason
- `tbl_Import_Profiles` could disappear or become empty, breaking auto-detect import
- the interactive import wizard made decisions inside `mod_Import_Logic`, which prevented clean service/UI separation

## Iteration history

### Iteration 1: Safe startup schema alignment

Implemented in:
- `modules/mod_App_Init.bas`
- `modules/mod_Schema_Manager.bas`

Changes:
- `InitializeApp` now calls `InitDatabaseStructure True`
- `InitDatabaseStructure` no longer drops `tbl_Import_Buffer`
- buffer initialization now uses safe schema alignment instead of destructive recreate
- `mod_Schema_Manager` gained:
  - `EnsureBufferStructure`
  - expanded canonical field lists for buffer/master
  - `GetCanonicalFieldType`

Outcome:
- startup became idempotent for the import buffer
- startup no longer risks deleting production buffer data as part of initialization
- schema evolution moved closer to a single center in `mod_Schema_Manager`

### Iteration 2: Search debug load reduction

Implemented in:
- `forms/uf_Search.cls`

Changes:
- search debug file logging is now enabled only when `LogLevel = DEBUG`
- normal typing/search no longer writes to `debug_search.log` on every keystroke

Outcome:
- reduced hot-path file I/O during search
- lower runtime noise in normal INFO/ERROR environments

### Iteration 3: Service/UI decoupling for maintenance and top-level import flow

Implemented in:
- `modules/mod_Maintenance_Logic.bas`
- `modules/mod_Import_Logic.bas`
- `forms/uf_Dashboard.cls`
- `forms/uf_Settings.cls`

Changes:
- added result-style service APIs:
  - `RunDataHealthCheckResult`
  - `ExportValidationLogResult`
  - `CreateDatabaseBackupResult`
  - `ClearValidationLogResult`
  - `FactoryResetDataResult`
  - `ImportExcelDataResult`
- maintenance and top-level import logic no longer show user dialogs directly
- forms now own message display and confirmations for:
  - backup
  - clear validation log
  - health check
  - validation export
  - import success/failure summary

Outcome:
- maintenance services now return status/message/error instead of mixing in UI
- the project moved to a clearer service contract without breaking existing flows
- old boolean wrappers stayed in place for compatibility where needed

### Iteration 4: Import diagnostics and profile self-heal

Implemented in:
- `modules/mod_Analysis_Logic.bas`
- `modules/mod_Import_Logic.bas`
- `forms/uf_Settings.cls`

Changes:
- `RunFullSyncProcess` now shows real `Import details` when import fails
- `SelectExcelFile` treats `ImportFolderPath = N/A` as an empty configured path
- `DetectBestProfile` now ensures `tbl_Import_Profiles` exists and default profiles are restored
- `uf_Settings.Form_Load` also ensures import profiles exist

Regression found and fixed:
- import started failing because `tbl_Import_Profiles` had disappeared while `tbl_Import_Mapping` still existed
- auto-detect profile therefore returned `0` and import stopped before sync

Outcome:
- import failures became diagnosable from UI
- default profiles `1/2/3` are restored automatically when missing
- the system recovered without touching `StaffState.accdb` manually in git

### Iteration 5: Import wizard UI extraction

Implemented in:
- `modules/mod_Import_Logic.bas`
- `forms/uf_Dashboard.cls`

Changes:
- removed direct `AskUserYesNo` / `InputBox` usage from `mod_Import_Logic`
- import logic now returns action payload when a user decision is required:
  - `RequiresUserAction`
  - `ActionType`
  - `ActionMessage`
  - `ActionTitle`
  - `ProfileID`
  - `ExcelField`
  - `SuggestedFieldName`
- introduced import decision helpers:
  - `CreateImportDecisionStore`
  - `SetSkipImportDecision`
  - `SetRestoreImportDecision`
  - `SetMapImportFieldDecision`
- `uf_Dashboard` now runs an import loop:
  - call import
  - inspect required action
  - ask the user in the form layer
  - store the decision
  - rerun import with the same file path

Outcome:
- import wizard decisions are now taken in UI instead of service logic
- the same selected Excel file is reused across repeated wizard steps
- skipping a new column remains a valid user choice without aborting the full import flow

### Iteration 6: Reports service/UI decoupling

Implemented in:
- `modules/mod_Reports_Logic.bas`
- `forms/uf_Dashboard.cls`

Changes:
- removed direct report-layer UI from `mod_Reports_Logic`
- added result-style APIs:
  - `GenerateAuditReportResult`
  - `GenerateCurrentStaffReportResult`
- kept legacy `GenerateAuditReport` / `GenerateCurrentStaffReport` as compatibility wrappers
- `uf_Dashboard` now owns user-facing status/messages for:
  - audit report generation
  - current staff snapshot generation

Outcome:
- report services now follow the same `result/status/message/error` contract used in maintenance/import flows
- UI dialogs for reports are isolated to form code
- `mod_Reports_Logic` is no longer a blocker for continued service/UI separation

### Iteration 7: Default mapping alignment and linked schema self-heal

Implemented in:
- `modules/mod_Schema_Manager.bas`
- `modules/mod_Import_Logic.bas`

Changes:
- aligned default profile 1 mapping with the current canonical master fields:
  - `BirthDate_Text` instead of `BirthDate`
  - `WorkStatus`
  - `PosName` instead of `Position` for the main job title field
- upgraded `AddMappingIfNotExists` so reseed updates stale `TargetField` values for existing `ExcelHeader` rows
- fixed `EnsureFieldExists` for split databases:
  - if table is linked, DDL is executed against the physical back-end database
  - link is refreshed afterward
- added schema self-heal call before import:
  - `ImportExcelDataResult` now runs `InitDatabaseStructure True` before buffer work starts

Regression found and fixed:
- import still failed after mapping changes because schema alignment was attempting `ALTER TABLE` against linked front-end tables
- this left fields such as `PosName` absent in the physical back-end even though code and mapping expected them

Outcome:
- default mapping profile now matches the current canonical schema more closely
- stale profile 1 mappings can be corrected by reseed without manual DB edits
- schema drift in linked `tbl_Import_Buffer` / `tbl_Personnel_Master` is healed through the back-end instead of silently failing in the front-end

## Files affected across the wave

- `forms/uf_Dashboard.cls`
- `forms/uf_Search.cls`
- `forms/uf_Settings.cls`
- `modules/mod_App_Init.bas`
- `modules/mod_Analysis_Logic.bas`
- `modules/mod_Import_Logic.bas`
- `modules/mod_Maintenance_Logic.bas`
- `modules/mod_Reports_Logic.bas`
- `modules/mod_Schema_Manager.bas`

## User-visible effects

- startup is safer and less noisy
- search debug logging no longer floods disk in normal mode
- backup / health-check / import feedback is still visible to users, but now originates from forms
- `Full Update` shows real import failure reasons
- missing import profiles are restored automatically
- import wizard prompts still work, but are now owned by `uf_Dashboard`
- report success/no-data/error feedback is now also owned by forms
- reseeding the default mapping can repair stale `TargetField` values
- schema alignment now works for linked back-end tables instead of failing on front-end DDL

## Architecture effects

- `mod_Schema_Manager` became more central for schema evolution
- `mod_Maintenance_Logic` moved from UI-coupled procedures toward result-returning services
- `mod_Import_Logic` now exposes service state plus action requests instead of directly running dialogs
- `uf_Dashboard` became the first orchestrator for interactive import decisions
- `mod_Reports_Logic` now follows the same result-based service contract direction
- linked-table schema maintenance now uses a physical back-end path instead of assuming monolithic Access DDL

## Remaining follow-up work

1. Continue service/UI decoupling in `mod_Analysis_Logic`.
2. Align `uf_Settings` mapping field list with the canonical master/buffer field set so UI choices cannot drift from schema.
3. Reduce dangerous `On Error Resume Next` in import/sync/schema/export code paths.
4. Optimize `SyncBufferToMaster` to avoid row-by-row `FindFirst`.
5. Add regression coverage for import, reseed alignment, linked schema self-heal, and full update diagnostics.
