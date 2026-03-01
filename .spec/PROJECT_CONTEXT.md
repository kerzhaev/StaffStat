# PROJECT CONTEXT: StaffState (Штаты) - MS Access/VBA

## Current State
- **Phase 31 (Multi-Profile Mapping & Interactive Wizard)** is completed.
  - Внедрена поддержка нескольких профилей маппинга (`tbl_Import_Profiles`). По умолчанию создаются 3 профиля: Основной, Снабжение, Финансы.
  - В `mod_Import_Logic` добавлен умный авто-детект (`DetectBestProfile`): система перед импортом сканирует заголовки Excel и сама выбирает подходящий профиль, где есть привязка `PersonUID` и максимальное совпадение колонок.
  - Реализован «Интерактивный мастер импорта»: при нахождении неизвестной колонки система предлагает пользователю прямо на лету создать новое поле в базе (`ALTER TABLE`), назначить ему тип данных и добавить связь в маппинг.
  - Добавлен вывод списка пропущенных (неразмеченных) колонок по итогам импорта.
  - Оптимизирована функция `GetFirstSheetName`: тяжелый запуск `Excel.Application` заменен на мгновенное чтение схемы через DAO.
- **Phase 30 (UI Localization Engine)** is completed.
  - Created `tbl_Localization` table for storing Russian UI translations (DB encoding protection).
  - Implemented a fast caching engine based on `Scripting.Dictionary` (Late Binding) in `mod_UI_Helpers`, loaded into memory at app start (`InitLocalization`).
  - Added `GetLoc(strKey)` function for instantaneous translation retrieval from cache.
  - Added `SeedLocalizationTable` seeder in `mod_Schema_Manager` for basic dictionary population.
- **Phase 29 (Batch Transactions & Performance)** is completed.
  - Implemented batch transactions in `SyncBufferToMaster` (commit every 2000 records) for datasets > 30,000.
  - Added `DoEvents` for UI responsiveness.
  - Expanded `dbMaxLocksPerFile` limit to 200,000.
- **Phase 28 (Explicit UI Error Handling)** is completed.
  - Full removal of `On Error Resume Next` in `uf_PersonCard`.
  - Explicit `HasField` check for Recordset columns before access.
  - Complete error logging via `mod_App_Logger` for 100% predictability and data protection.
- **Phase 27 (Smart DDL Typing & Dynamic Schema)** is completed.
  - UI-driven data type selection (Text, Date, Number) for new fields.
  - Dynamic type shifting via double-click in `uf_Settings`.

  - `mod_Schema_Manager` performs `ALTER TABLE` directly in BE file followed by `RefreshLink` in FE.
- **Phase 26 (Parameterized SQL Queries)** is completed.
  - 100% protection against SQL injections and quote encoding errors.
  - Standardized use of `DAO.QueryDef` with parameters for all DML operations (INSERT/UPDATE).
- **Phase 25 (Split Database Architecture)** is completed.
  - Database physically split into Front-End (`StaffState.accdb`) and Back-End (`StaffState_BE.accdb`).
  - `mod_Table_Relinker` implemented for automatic BE path restoration on startup.
  - `FactoryResetData` function added for rapid test data cleanup.
- **Phase 24 (100% English Codebase & Encoding Fix)** is completed.
  - Complete removal of Cyrillic from source files (.bas, .cls).
  - Transition to English system messages and comments to protect encoding with Git and AI agents.
- Phase 10 (Change Report) is implemented and working.
- Dashboard includes start/end date inputs for period reporting.
- Dashboard date inputs now normalize dot-separated dates and change report SQL uses forced MM/DD/YYYY literals.
- Report export logic generates change reports for selected dates.
- Phase 9 (Workflow Automation) is implemented with full sync pipeline.
- UI updates include a new dashboard button and a manual controls section.
- Phase 8 (Dynamic Search & Export) is fully operational and verified.
- Search and export work across all business fields via TableDefs.
- CurrentDb reference is stabilized in dynamic search (no invalid object errors).

## Implemented Modules
- mod_App_Init.bas
- mod_App_Logger.bas
- mod_Import_Logic.bas
- mod_Analysis_Logic.bas
- mod_Schema_Manager.bas
- mod_UI_Helpers.bas
- mod_Export_Logic.bas
- mod_Fix_Startup.bas
- mod_Table_Relinker.bas

## Key Tables
- tbl_Import_Buffer
- tbl_Personnel_Master
- tbl_History_Log
- tbl_Import_Meta

## UI Forms
- uf_Dashboard: All hardcoded Russian text (Captions, MsgBox, InputBox) replaced with `GetLoc` calls.
- uf_Search: All hardcoded Russian text replaced with `GetLoc` calls.
- uf_Settings: All hardcoded Russian text replaced with `GetLoc` calls.
  - Добавлен свободный комбобокс `cboProfile` для переключения активного профиля маппинга.
  - Добавлена кнопка `btnEditMapping` для быстрого редактирования заголовков Excel (изменение опечаток без пересоздания связи).
  - Подписи `lblExcel` и `lblDB` переведены из формата TextBox в Label и локализованы.
- uf_PersonCard: All hardcoded Russian text replaced with `GetLoc` calls. Dismissal status check now uses the localization key `STATUS_DISMISSED` instead of assembling the word via `ChrW()`.

## History
- **Phase 31 (2026-03-01)**:
  - Multi-Profile Mapping: Внедрена поддержка нескольких профилей маппинга (`tbl_Import_Profiles`). По умолчанию создаются 3 профиля: Основной, Снабжение, Финансы.
  - Auto-Detect: В `mod_Import_Logic` добавлен умный авто-детект (`DetectBestProfile`): система перед импортом сканирует заголовки Excel и сама выбирает подходящий профиль, где есть привязка `PersonUID` и максимальное совпадение колонок.
  - Interactive Wizard: Реализован «Интерактивный мастер импорта»: при нахождении неизвестной колонки система предлагает пользователю прямо на лету создать новое поле в базе (`ALTER TABLE`), назначить ему тип данных и добавить связь в маппинг.
  - Missed Columns: Добавлен вывод списка пропущенных (неразмеченных) колонок по итогам импорта.
  - Performance: Оптимизирована функция `GetFirstSheetName`: тяжелый запуск `Excel.Application` заменен на мгновенное чтение схемы через DAO.
- **Phase 30 (2026-03-01)**:
  - UI Localization: Created `tbl_Localization` for safe storage of Russian UI translations.
  - Caching Engine: Implemented `Scripting.Dictionary` caching in `mod_UI_Helpers` (`InitLocalization`, `GetLoc`) for rapid retrieval.
  - Seeder: Added `SeedLocalizationTable` in `mod_Schema_Manager`.
  - UI Refactoring: Replaced all Russian hardcoded strings in `uf_Dashboard`, `uf_Search`, `uf_Settings`, and `uf_PersonCard` with `GetLoc`.
  - Dismissal Status: `uf_PersonCard` now uses `STATUS_DISMISSED` localization key instead of `ChrW()` assembly.
- **Phase 29 (2026-03-01)**:
  - Batch Transactions: Implemented in `SyncBufferToMaster` with commit every 2000 records to prevent MS Access freezing on large sets (> 30k records).
  - UI Responsiveness: Integrated `DoEvents` to keep the Windows UI thread alive during long operations.
  - Performance Tuning: Set `dbMaxLocksPerFile` to 200,000 to support large transactional batches.
- **Phase 28 (2026-03-01)**:
  - Explicit UI Error Handling: Full removal of blind error suppression in `uf_PersonCard`.
  - Field Verification: Integrated `HasField` for safe Recordset access.
  - Logging: Full integration with `mod_App_Logger` to prevent data loss and ensure UI stability.
- **Phase 27 (2026-03-01)**:
  - Smart DDL Typing: UI for selecting data types (Text, Date, Number).
  - Dynamic Schema: On-the-fly type modification in `uf_Settings`.
  - BE Direct Modification: `mod_Schema_Manager` executes `ALTER TABLE` in BE and refreshes FE links.
- **Phase 26 (2026-03-01)**:
  - Parameterized SQL Queries: Full migration to `DAO.QueryDef` for DML.
  - Security: 100% protection against SQL injection and encoding issues.
- **Phase 25 (2026-03-01)**:
  - Split Database: Physical separation into FE and BE (`StaffState_BE.accdb`).
  - Auto-Relinking: Added `mod_Table_Relinker` for path restoration.
  - Maintenance: Added `FactoryResetData` for test data clearing.
- **Phase 24 (2026-02-28)**:
  - 100% English Codebase & Encoding Fix.
  - Complete removal of Cyrillic in source files (.bas, .cls).
  - Transition to English system messages and comments to protect encoding when working with Git and AI agents.
- Phase 10 (2026-02-01):
  - Implemented period change report export with date range filtering.
  - Added dashboard inputs for start/end dates and report action.
  - Export logic generates Excel report for the selected period.
- Phase 9 (2026-02-01):
  - Implemented workflow automation with a full sync pipeline for import, analysis, and history logging.
- Phase 8 (2026-02-01):
  - Universal Dynamic Search across all business fields (TableDefs-driven).
  - Dynamic Excel Export via Late Binding with auto-formatting and field mapping.