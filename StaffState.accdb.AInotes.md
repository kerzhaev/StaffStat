# StaffState.accdb — Notes

**Project:** StaffState (Дозор). Source of truth: `StaffState.accdb.src/`.  
**Phase:** 19–20 (Universal Mapping + English schema), **Phase 21 (Legacy migration).**

## Tables (from tbldefs)

- **tbl_Personnel_Master** — Main personnel table. PK: PersonUID. English-only columns. Definition: `tbldefs/tbl_Personnel_Master.sql`. Recreated via `mod_Schema_Manager.ForceRebuildSchema`.
- **tbl_Import_Buffer** — Temp import from Excel. PK: ID (COUNTER). English-only columns; definition: `tbldefs/tbl_Import_Buffer.sql`. Import logic is **strict**: only columns present in tbl_Import_Mapping are filled; unmapped Excel columns are **skipped**; PersonUID **must** be mapped to proceed.
- **tbl_Import_Mapping** — ProfileID, ExcelHeader, TargetField. Created by `CreateImportMappingTable`; Profile 1 seeded by **SeedImportMappingProfile1** (all 44 mappings hardcoded in `mod_Schema_Manager.bas` with ChrW/CyrStr; groups: Personal, Contract, Order, Banking, Sizes + English aliases). **Persistence:** .sql stores schema only; data is restored by running SeedImportMappingProfile1 in Access after Build.
- **tbl_Import_Profiles** — ProfileID (PK), ProfileName, IdStrategy. Created by `CreateImportProfilesTable`; default ProfileID=1.
- **tbl_Import_Meta** — ExportFileDate, ImportRunAt, SourceFilePath.
- **tbl_History_Log** — Audit log. LogID (PK), PersonUID, ChangeDate, FieldName, OldValue, NewValue. **Phase 21:** All FieldName values are English only; legacy Cyrillic values were migrated via UPDATE (40 mapping pairs from Phase 20). Current DB had only "_System" in FieldName at migration time.
- **tbl_Settings** — SettingKey (PK), SettingValue, SettingGroup, Description.
- **tbl_Validation_Log** — LogID (PK), RecordID, TableName, ErrorType, ErrorMessage, CheckDate.

## Phase 21 — tbl_History_Log FieldName migration (2026-02-07)

UPDATEs were run for the following Cyrillic → English pairs (Phase 20 mapping). If your DB had legacy data, these rows were updated; otherwise 0 rows affected.

| Cyrillic (old) | English (new) |
|----------------|---------------|
| Дата_рождения | BirthDate_Text |
| Возраст_сотрудника | EmployeeAge |
| Пол | Gender |
| Семейное_положение | MaritalStatus |
| Количество_детей | ChildrenCount |
| Национальность | Nationality |
| Гражданство | Citizenship |
| Дата_увольнения | DismissalDate |
| Группа_сотрудников | EmployeeGroup |
| Вид_контракта | ContractType |
| Тип_контракта | ContractKind |
| Дата_начала_контракта | ContractStartDate |
| Дата_окончания_контракта | ContractEndDate |
| Срок_договора_год | ContractYears |
| Срок_договора_месяц | ContractMonths |
| Номер_приказа | OrderNumber |
| Дата_приказа | OrderDate_Text |
| Вид_мероприятия | EventType |
| Причина_мероприятия | EventReason |
| Начало_срока_действия | ValidFromDate |
| Конец_срока_действия | ValidToDate |
| Штатная_должность | StaffPosition |
| Должность | Position |
| Единица_расчета | CalculationUnit |
| Рассчитано_по | CalculatedBy |
| Штатная_должность1 | StaffPosition1 |
| ВУС | VUS |
| Тарифный_разряд | SalaryGrade |
| Дата_приказа_Л_С | OrderDate_LS |
| Номер_приказа1 | OrderNumber1 |
| Чей_приказ_должность | PositionOrderIssuer |
| Раздел_персонала | PersonnelDivision |
| Раздел_персонала1 | PersonnelDivision1 |
| Статус_занятости | EmploymentStatus |
| Контрольный_банковский_ключ | BankControlKey |
| Номер_счета_в_банке | BankAccountNumber |
| Получатель | Payee |
| Ключ_банка | BankKey |
| Размер_Сапог | BootSize |
| Охват_головы | HeadSize |
| ФИО | FullName |

## After Build from Source (tbl_Import_Mapping empty)

1. Click **Import from Source** in Access (msaccess-vcs-addin) to load updated VBA.
2. Run **mod_Schema_Manager.SeedImportMappingProfile1** (F5 in mod_Schema_Manager). This CLEARs Profile 1 and INSERTs all 44 mappings.

## Adding a new mapping (persistence)

Add the mapping in **two** places: (1) in the live DB (INSERT into tbl_Import_Mapping) for immediate use, and (2) in **SeedImportMappingProfile1** in `modules/mod_Schema_Manager.bas` (use ChrW/CyrStr for Russian ExcelHeader) so it persists after the next Build from Source.

## Restore English-only state (Buffer cleanup)

If tbl_Import_Buffer has Cyrillic columns: close all tables, run **mod_Schema_Manager.CleanupImportBuffer** or **DeepCleanSchema**. MCP cannot DROP COLUMN on Access; use VBA.

## After schema reset (tables dropped)

1. Run **mod_Schema_Manager.ForceRebuildSchema** (F5).
2. Run **CreateImportMappingTable** and **SeedImportMappingProfile1** (or InitDatabaseStructure).

Plan: `.spec/phase-19-20-universal-mapping-and-english-schema.md`. Phase 21: `.spec/PROJECT_CONTEXT.md` (version 0.12).
