# Implementation Plan: Phase 11 - Data Integrity & System Settings
**Spec:** `.spec/feature.md`

## 1. File Structure Changes (Изменения в файлах)

- **Create:** `StaffState.accdb.src/modules/mod_Maintenance_Logic.bas`
- **Modify:** `StaffState.accdb.src/modules/mod_Schema_Manager.bas` - add settings table management functions
- **Modify:** `StaffState.accdb.src/modules/mod_App_Init.bas` - add initialization for settings and validation log tables
- **Modify:** `StaffState.accdb.src/forms/uf_Dashboard.cls` - add system administration section with buttons
- **Modify:** `StaffState.accdb.src/modules/mod_Export_Logic.bas` - add system summary export function

## 2. Detailed Steps (Пошаговый план)

### Step 1: Database Schema Updates

- [ ] **1.1:** Update `mod_Schema_Manager.bas` to add `CreateSettingsTable()` function
  - Use DAO to create `tbl_System_Settings` with fields: SettingKey (PK), SettingValue, Category, Description, LastModified, ModifiedBy
  - Add `EnsureSettingsTableExists()` helper called from InitDatabaseStructure
  - Use CP1251 comments (English ASCII) in code

- [ ] **1.2:** Update `mod_App_Init.bas` to add `CreateValidationLogTable()` function
  - Use DAO to create `tbl_Validation_Log` with all required fields (ID, LogDate, PersonUID, FieldName, ValidationRule, ActualValue, Severity, ImportMetaID, Resolved, ResolvedDate, ResolvedBy)
  - Add call to `CreateValidationLogTable()` in `InitDatabaseStructure()`
  - Ensure proper error handling with On Error GoTo ErrorHandler

- [ ] **1.3:** Update `mod_Schema_Manager.bas` to add settings access functions
  - `GetSetting(strKey As String, Optional strDefault As String = "") As String` - retrieves setting value via DAO query
  - `SetSetting(strKey As String, strValue As String, Optional strCategory As String = "General", Optional strDescription As String = "")` - inserts or updates setting via DAO Execute
  - Both functions use DAO with proper error handling and SQL escaping

### Step 2: Maintenance Logic Module (`mod_Maintenance_Logic.bas`)

- [ ] **2.1:** Create new module `mod_Maintenance_Logic.bas` with Option Explicit
  - Add module header with English comments (CP1251 encoding)
  - Define constants for maintenance check types

- [ ] **2.2:** Implement `CheckOrphanedRecords()` function
  - Query `tbl_History_Log` for PersonUID values not in `tbl_Personnel_Master`
  - Return collection or array of orphaned record IDs
  - Use DAO Recordset with proper error handling
  - Log findings via `mod_App_Logger.LogInfo()`

- [ ] **2.3:** Implement `CheckIndexIntegrity()` function
  - Verify expected indexes exist on key tables (PersonUID index on Master, etc.)
  - Compare against expected index list
  - Return list of missing indexes
  - Use DAO TableDef.Indexes collection

- [ ] **2.4:** Implement `CheckSchemaConsistency()` function
  - Compare Buffer and Master table schemas using TableDefs
  - Identify fields in Buffer not in Master (potential sync issues)
  - Identify fields in Master not in Buffer (potential schema drift)
  - Return collection of inconsistencies

- [ ] **2.5:** Implement `GetDatabaseStatistics()` function
  - Count records in each major table (Master, History_Log, Import_Buffer, Import_Meta)
  - Count active vs inactive personnel
  - Count validation log entries by severity
  - Return dictionary or custom type with statistics

- [ ] **2.6:** Implement `RunMaintenanceChecks()` main function
  - Call all check functions sequentially
  - Aggregate results into maintenance report structure
  - Log summary via `mod_App_Logger.LogInfo()`
  - Return collection or custom type with all findings
  - Include error handling for each check (continue on failure)

### Step 3: Validation Logging Integration

- [ ] **3.1:** Add `LogValidationIssue()` function to `mod_App_Logger.bas` or create helper in `mod_Maintenance_Logic.bas`
  - Parameters: PersonUID, FieldName, ValidationRule, ActualValue, Severity, Optional ImportMetaID
  - Insert record into `tbl_Validation_Log` via DAO Execute
  - Use proper SQL escaping and error handling
  - Set LogDate to Now() automatically

- [ ] **3.2:** Update existing validation code (e.g., PersonUID validation in `mod_Analysis_Logic.bas`) to call `LogValidationIssue()`
  - Replace or supplement existing LogInfo calls with structured validation logging
  - Ensure ImportMetaID is passed when available

### Step 4: Export Logic Updates (`mod_Export_Logic.bas`)

- [ ] **4.1:** Add `ExportSystemSummary()` function
  - Use late binding Excel (CreateObject("Excel.Application"))
  - Create workbook with multiple sheets: Summary, Validation Log, Maintenance Report, Statistics
  - Query `tbl_Validation_Log` for recent entries (last 30 days or configurable)
  - Query maintenance check results
  - Format Excel cells (headers bold, auto-width columns)
  - Save to user-selected path or default location
  - Proper error handling and cleanup (Quit Excel, release objects)

- [ ] **4.2:** Add helper function `FormatValidationLogForExport()` if needed
  - Transform validation log records into export-friendly format
  - Include filtering options (date range, severity)

### Step 5: UI Integration (`uf_Dashboard.cls`)

- [ ] **5.1:** Add new section to Dashboard form: "System Administration"
  - Add label "Системное администрирование" (CP1251 string)
  - Position below existing sections

- [ ] **5.2:** Add button "Настройки системы" (System Settings)
  - Click handler: Show message box with current settings (or placeholder for future settings form)
  - Use `GetSetting()` to display key configuration values

- [ ] **5.3:** Add button "Проверка целостности" (Run Maintenance Checks)
  - Click handler: Call `mod_Maintenance_Logic.RunMaintenanceChecks()`
  - Display results in message box with summary
  - Show "Found X orphaned records, Y missing indexes, Z schema inconsistencies"
  - Log completion via `mod_App_Logger.LogInfo()`

- [ ] **5.4:** Add button "Журнал валидации" (View Validation Log)
  - Click handler: Show message box with recent validation issues count
  - Display: "Recent validation issues: X errors, Y warnings in last 30 days"
  - Placeholder for future detailed viewer form

- [ ] **5.5:** Add button "Экспорт отчета" (Export System Summary)
  - Click handler: Call `mod_Export_Logic.ExportSystemSummary()`
  - Show file save dialog or use default path
  - Display success message: "Системный отчет экспортирован в [путь]"
  - Handle errors gracefully with user-friendly messages

### Step 6: Initialization Updates (`mod_App_Init.bas`)

- [ ] **6.1:** Update `InitDatabaseStructure()` to call new table creation functions
  - Add `CreateSettingsTable()` call
  - Add `CreateValidationLogTable()` call
  - Ensure tables are created only if they don't exist (use TableExists helper)

- [ ] **6.2:** Add default settings initialization (optional)
  - After creating settings table, insert default settings if table is empty
  - Examples: "ImportPath", "ValidationStrictMode", "MaintenanceSchedule"

## 3. Definition of Done (Завершение)

- [ ] All database tables (`tbl_System_Settings`, `tbl_Validation_Log`) created successfully via DAO
- [ ] Settings access functions (`GetSetting`, `SetSetting`) work correctly with proper SQL escaping
- [ ] Maintenance checks (`RunMaintenanceChecks`) execute without errors and return accurate results
- [ ] Validation logging integrated into existing validation code
- [ ] System summary export generates Excel file with all required sheets using late binding
- [ ] Dashboard UI updated with all four system administration buttons
- [ ] All error handlers implemented (On Error GoTo ErrorHandler) in every procedure
- [ ] Code comments in English (ASCII), UI strings in Russian (CP1251)
- [ ] Manual review of types, DAO usage, and late-binding safety
- [ ] **ОБЯЗАТЕЛЬНО:** Update `.spec/PROJECT_CONTEXT.md` (version and history) after coding is complete

## 4. Final Cleanup (Завершение)

"Никогда не перемещай в архив файлы PROJECT_CONTEXT.md и ROADMAP.md."

- [ ] **Context:** Update `.spec/PROJECT_CONTEXT.md` (version and history).
- [ ] **Archive:** Move this plan and its related spec to `.spec/archive/` after completion.

## Technical Notes

- **DAO Usage:** All database operations use DAO (CurrentDb, TableDefs, Recordsets, Execute)
- **Error Handling:** Every procedure must have On Error GoTo ErrorHandler with proper cleanup
- **Comments:** VBA code comments in English (ASCII), UI messages in Russian (CP1251 encoding)
- **Excel Binding:** Use late binding (`CreateObject("Excel.Application")`) for Excel export, never early binding
- **SQL Formatting:** Use UPPERCASE SQL keywords in all queries
- **Type Safety:** Use Long (not Integer), Option Explicit everywhere
