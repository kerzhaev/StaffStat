# Phase 19–20: Universal Mapping Engine & 100% English Schema

**Branch:** `feature/phase-19-20-mcp-core`  
**Status:** Plan  
**MCP:** AccessDB (schema inspection, SQL, data validation)

---

## 1. Goals (Combined)

| Goal | Description |
|------|--------------|
| **Universal mapping** | Engine to translate ANY Excel header (Cyrillic) → DB field via `tbl_Import_Mapping` + `tbl_Import_Profiles`. |
| **English core** | Refactor entire DB schema and VBA to 100% English field/column names (fix encoding issues). |
| **Encoding safety** | All Russian UI strings via `ChrW()`; no Cyrillic in .bas/.cls source. |

---

## 2. Current State (from tbldefs & code)

### 2.1 tbl_Personnel_Master

- **Already English:** PersonUID, SourceID, FullName, RankName, BirthDate, WorkStatus, PosCode, PosName, OrderDate, OrderNum, LastUpdated, IsActive, ID, and all *_Raw columns.
- **Cyrillic (to rename):**  
  `Дата_рождения`, `Возраст_сотрудника`, `Пол`, `Семейное_положение`, `Количество_детей`, `Национальность`, `Гражданство`, `Дата_увольнения`, `Группа_сотрудников`, `Вид_контракта`, `Тип_контракта`, `Дата_начала_контракта`, `Дата_окончания_контракта`, `Срок_договора_год`, `Срок_договора_месяц`, `Номер_приказа`, `Дата_приказа`, `Вид_мероприятия`, `Причина_мероприятия`, `Начало_срока_действия`, `Конец_срока_действия`, `Штатная_должность`, `Должность`, `Единица_расчета`, `Рассчитано_по`, `Штатная_должность1`, `ВУС`, `Тарифный_разряд`, `Дата_приказа_Л_С`, `Номер_приказа1`, `Чей_приказ_должность`, `Раздел_персонала`, `Раздел_персонала1`, `Статус_занятости`, `Контрольный_банковский_ключ`, `Номер_счета_в_банке`, `Получатель`, `Ключ_банка`, `Размер_Сапог`, `Охват_головы`, `F44`.

### 2.2 tbl_History_Log

- **Already English:** LogID, PersonUID, ChangeDate, FieldName, OldValue, NewValue.  
- **No Cyrillic columns.** Values in FieldName/OldValue/NewValue may reference old Cyrillic field names → after Master rename, history may show old names; consider migration or display mapping.

### 2.3 tbl_Import_Mapping (tbldefs)

- **Structure:** MappingID (PK), ProfileID, ExcelHeader (VARCHAR 255), TargetField (VARCHAR 100).
- **Seed data:** ProfileID=1 with ExcelHeader (Cyrillic/English) → TargetField (English: e.g. BirthDate_Raw, EmployeeAge, Gender, MaritalStatus, …).

### 2.4 tbl_Import_Profiles (tbldefs)

- **Structure:** ProfileID (PK), ProfileName (VARCHAR 100), IdStrategy (VARCHAR 20).

---

## 3. English Field Mapping (Cyrillic → English)

Target **single canonical set** for Master/Buffer (and Mapping TargetField). Use existing `tbl_Import_Mapping.sql` TargetField names where they map to “final” Master fields:

| Cyrillic (current) | English (target) |
|--------------------|------------------|
| Дата_рождения | BirthDate (or keep as-is if already used for normalized date; raw stays BirthDate_Raw) |
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
| Дата_приказа | OrderDate (already exists; duplicate logic in Buffer) |
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
| F44 | F44 (keep or rename to e.g. ExtraCode) |

---

## 4. Implementation Plan

### Step 1: MCP & tbldefs — Mapping tables

1. **Ensure DB has tables** (if using MCP on live .accdb):
   - Create `tbl_Import_Profiles` (ProfileID, ProfileName, IdStrategy).
   - Create `tbl_Import_Mapping` (MappingID, ProfileID, ExcelHeader, TargetField).
   - Insert default profile (e.g. ProfileID=1, ProfileName="Default", IdStrategy="PersonUID").
   - Insert rows from `tbldefs/tbl_Import_Mapping.sql` (INSERTs) for ProfileID=1.
2. **Keep tbldefs in sync:** `tbldefs/tbl_Import_Profiles.sql`, `tbldefs/tbl_Import_Mapping.sql` are source of truth; any MCP schema change must be reflected there.

### Step 2: Import engine uses Mapping table

1. In **mod_Import_Logic**:
   - Add dependency on default profile (e.g. GetDefaultProfileID() → 1 or from tbl_Import_Profiles).
   - Replace **MapFieldName()** logic: first look up `ExcelHeader` in `tbl_Import_Mapping` for current profile; if found, use `TargetField`; else fallback to SanitizeFieldName (and optionally still IsPersonUIDColumn for PersonUID_Raw).
   - PersonUID detection: keep IsPersonUIDColumn() or add mapping row "Личный номер" / "PersonUID" → PersonUID_Raw.
2. **Buffer columns:** TargetField values from Mapping are Buffer column names (e.g. BirthDate_Raw, EmployeeAge, Gender). So Buffer schema is extended by EnsureFieldExists for each TargetField.

### Step 3: Schema refactor — Master/Buffer to English only

1. **Order of operations (no data loss):**
   - Add new English columns to `tbl_Personnel_Master` (and Buffer if needed).
   - One-time data migration: UPDATE Master SET NewEnglishField = OldCyrillicField.
   - Update all VBA references to use English names.
   - Drop Cyrillic columns from Master (and Buffer).
   - Update tbldefs (tbl_Personnel_Master.sql, tbl_Import_Buffer.sql).
2. **SyncMasterStructure:** Must only create English field names (from Buffer’s TargetField set or from a fixed list). No longer copy “any” Buffer field name (which could be Cyrillic) into Master.
3. **mod_Schema_Manager:** EnsureFieldExists / SyncMasterStructure use only English names; mapping from Buffer column name (which is already English after Step 2) to Master is 1:1 or via BufferFieldToMasterName.

### Step 4: BufferFieldToMasterName & CopyAllFields (Analysis)

1. **BufferFieldToMasterName:** Extend to map all new English Buffer fields to English Master fields (e.g. EmployeeAge → EmployeeAge, Gender → Gender). Remove Cyrillic branch; all names English.
2. **CopyAllFields:** Same mapping; no Cyrillic field names.

### Step 5: tbl_History_Log

- **FieldName** stores which field changed. After schema rename, existing rows may still have Cyrillic names. Options:
  - **A)** Leave as-is; display layer (e.g. uf_PersonCard) translates Cyrillic FieldName → Russian label via ChrW() when displaying.
  - **B)** One-time UPDATE tbl_History_Log SET FieldName = EnglishName WHERE FieldName = CyrillicName for each pair.
- Recommendation: **A** for backward compatibility and simpler rollout; optional **B** later for consistency.

### Step 6: UI and ChrW()

- **mod_UI_Helpers / forms:** All user-visible Russian text via ChrW() (or central constants) so .bas/.cls contain no Cyrillic.
- **uf_PersonCard TranslateFieldName:** Input can stay Russian; output for filtering must be **English** FieldName (Master/tbl_History_Log). So TranslateFieldName returns English field name; display labels from a separate function (ChrW-based).

### Step 7: Reports & Export

- **mod_Reports_Logic, mod_Export_Logic:** Column titles for Excel: use English or ChrW() Russian; internal field names English.

---

## 5. File Checklist

| Item | Action |
|------|--------|
| tbldefs/tbl_Import_Profiles.sql | Already exists; ensure structure and seed (e.g. ProfileID=1). |
| tbldefs/tbl_Import_Mapping.sql | Already exists with many mappings; align TargetField with final English Master names. |
| tbldefs/tbl_Personnel_Master.sql | Replace Cyrillic column names with English; keep same types. |
| tbldefs/tbl_Import_Buffer.sql | Replace Cyrillic column names with English. |
| mod_Import_Logic.bas | Use tbl_Import_Mapping for MapFieldName; ChrW() for any new UI strings. |
| mod_Schema_Manager.bas | SyncMasterStructure: only English fields; EnsureFieldExists called with English names. |
| mod_Analysis_Logic.bas | BufferFieldToMasterName + CopyAllFields: full English mapping; no Cyrillic. |
| mod_Maintenance_Logic.bas | Any field references → English. |
| mod_Reports_Logic.bas | Field names English; UI strings ChrW(). |
| mod_Export_Logic.bas | Same. |
| uf_PersonCard (TranslateFieldName) | Return English FieldName; labels via ChrW(). |
| uf_Dashboard, uf_Search, uf_Settings | ChrW() for all Russian. |

---

## 6. MCP Protocol Reminder

- If schema is changed via MCP (create/alter tables), **always** update the corresponding `.sql` in `tbldefs/` and the VBA that references those objects.
- After MCP or code changes, remind user to run **Import from Source** in Access (msaccess-vcs-addin) to sync binary .accdb with .accdb.src.

---

## 7. VBA-POWERED CLEANUP (Emergency)

If MCP DROP COLUMN did not persist (e.g. table locked), run **FinalSchemaCleanup** in Access:

1. **Import from Source** (msaccess-vcs-addin) to load the new VBA.
2. **Close all tables** in MS Access (including tbl_Personnel_Master).
3. Open **mod_Schema_Manager**, run **FinalSchemaCleanup** (F5).
4. The procedure: (a) Ensures all English columns exist; (b) Drops all 40 Cyrillic columns (names built via ChrW() in VBA).

**MCP verification after user runs VBA:** Reconnect to `C:\Users\Nachfin\Desktop\Projets\StaffState\StaffState.accdb` and run:
- `SELECT TOP 1 * FROM [tbl_Personnel_Master]` — result keys must be only English; no Cyrillic column names.
- `SELECT TOP 1 * FROM [tbl_Import_Buffer]` — result keys must be only English; no Cyrillic column names.
- Ensure **ContractYears** and **ContractMonths** exist (LONG); **Срок_договора_год** and **Срок_договора_месяц** must be dropped.

**Profile 1 mappings for ContractYears/ContractMonths:** In `tbldefs/tbl_Import_Mapping.sql`: ExcelHeader "Срок договора год" -> TargetField "ContractYears", "Срок договора месяц" -> "ContractMonths". Run these INSERTs in Access if tbl_Import_Mapping exists and lacks these rows.

---

## 8. Execution Order (Recommended)

1. Create branch (done: `feature/phase-19-20-mcp-core`).
2. Apply tbl_Import_Profiles + tbl_Import_Mapping in DB (via MCP or manual), and ensure tbldefs match.
3. Implement mapping-driven MapFieldName in mod_Import_Logic (keep fallback).
4. Add English columns to Master/Buffer; migrate data; switch VBA to English; drop Cyrillic columns; update tbldefs.
5. Refactor BufferFieldToMasterName / CopyAllFields / SyncMasterStructure.
6. ChrW() for UI and TranslateFieldName output to English.
7. Test import with Cyrillic Excel headers; test sync and history display.
