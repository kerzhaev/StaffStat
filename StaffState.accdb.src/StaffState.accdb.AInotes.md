# StaffState.accdb — Notes

**Project:** StaffState (Дозор). Source of truth: `StaffState.accdb.src/`.  
**Phase:** 19–20 (Universal Mapping + English schema).

## Tables (from tbldefs)

- **tbl_Personnel_Master** — Main personnel table. PK: PersonUID. English-only columns. Definition: `tbldefs/tbl_Personnel_Master.sql`. Recreated via `mod_Schema_Manager.ForceRebuildSchema`.
- **tbl_Import_Buffer** — Temp import from Excel. PK: ID (COUNTER). English-only columns only; definition: `tbldefs/tbl_Import_Buffer.sql`. Import logic is **strict**: only columns present in tbl_Import_Mapping are filled; no new Cyrillic columns are created.
- **tbl_Import_Mapping** — ProfileID, ExcelHeader (e.g. ID_Сотрудника, ФИО_Полное, ДР), TargetField (English). Profile 1 must map: ID_Сотрудника→PersonUID, ФИО_Полное→FullName, ДР→BirthDate_Text. Seeded in `tbldefs/tbl_Import_Mapping.sql`.
- **tbl_Import_Profiles** — ProfileID (PK), ProfileName, IdStrategy.
- **tbl_Import_Meta** — ExportFileDate, ImportRunAt, SourceFilePath.
- **tbl_History_Log** — Audit log. LogID (PK), PersonUID, ChangeDate, FieldName, OldValue, NewValue.
- **tbl_Settings** — SettingKey (PK), SettingValue, SettingGroup, Description.
- **tbl_Validation_Log** — LogID (PK), RecordID, TableName, ErrorType, ErrorMessage, CheckDate.

## Restore English-only state (Phase 1 cleanup)

If tbl_Import_Buffer already has **Cyrillic columns** from an older import:
1. Close all tables in Access.
2. Run **mod_Schema_Manager.CleanupImportBuffer** (or **DeepCleanSchema** for both Master and Buffer). This DROPs Cyrillic columns via VBA (MCP cannot run DDL like ALTER TABLE DROP COLUMN on MS Access).

## After schema reset (tables dropped)

1. Run **mod_Schema_Manager.ForceRebuildSchema** (F5) — recreates tbl_Personnel_Master and tbl_Import_Buffer from VCS definitions.
2. Click **Import from Source** in Access to load data. Only mapped English columns are filled; unmapped Excel columns are skipped and logged as "Unmapped column skipped".

## MCP usage

- Connect to binary .accdb (e.g. parent folder).
- After schema change: update `tbldefs/*.sql` and VBA in `.accdb.src/`.

Plan: `.spec/phase-19-20-universal-mapping-and-english-schema.md`.
