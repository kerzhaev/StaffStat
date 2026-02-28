# Completion Summary: bugfix/backup-and-search-stability

**Target branch:** `bugfix/backup-and-search-stability`  
**Date:** 2026-02-03

## 1. Cyrillic encoding (no more ????? in messages)

### mod_UI_Helpers
- **RuSearchCaption()** — "Поиск сотрудников" (ChrW)
- **RuBackupPathUndefined()** — message when path to DB is not defined
- **RuBackupSaved(strPath)** — "Резервная копия сохранена: ..."
- **RuBackupFailedLocked()** — Error 70: "Cannot backup while database is heavily used. Please close other Access windows and try again." (ChrW)
- **RuBackupFailedGeneric(strErrDesc)** — generic backup failure + hint to close Access or copy manually to Backups
- **RuValidationLogCleared()** — "Журнал очищен" (ChrW)

All user-facing Russian strings in backup/validation flow now built via ChrW in mod_UI_Helpers to avoid encoding issues on different Windows locales.

### mod_Maintenance_Logic
- CreateDatabaseBackup: all messages use mod_UI_Helpers.Ru* functions (no raw Cyrillic literals).
- ClearValidationLog: uses RuValidationLogCleared().

## 2. Search reset (uf_Search)

- **m_strMode** (Private): stores OpenArgs mode string (e.g. "MODE=DUPLICATES"); empty = normal search. Set in Form_Load and LoadDuplicatesModeAndNotify.
- **btnClear_Click** when form was opened in MODE=DUPLICATES (or m_strMode <> ""):
  - Sets **m_strMode** = "" and **m_fDuplicatesMode** = False.
  - Clears **lstResults.RowSource** and Requery.
  - Re-enables **txtFilter**, clears its value.
  - Sets form **Caption** to "Поиск сотрудников" via **RuSearchCaption()** (no lblHeader on form; caption used).
  - Enables **btnSearch**, **btnClear**, disables **btnExportExcel**; sets focus to txtFilter.
- On normal open, Form_Load sets Caption = RuSearchCaption() when not in duplicates mode.

## 3. Backup logic (mod_Maintenance_Logic)

- **CreateDatabaseBackup** uses **only** `FileSystemObject.CopyFile` (Late Binding: CreateObject("Scripting.FileSystemObject")). No Application.CompactRepair.
- Path to backup folder: **strBackupFolder = CurrentProject.Path & "\Backups"** (derived from CurrentProject.Path).
- **Error 70 (Permission Denied):** handled explicitly; shows RuBackupFailedLocked() — "Cannot backup while database is heavily used. Please close other Access windows and try again."
- Other errors: ShowMessage(RuBackupFailedGeneric(Err.Description)).
- Success: ShowMessage(RuBackupSaved(strTargetFile)).

## 4. Settings integration (mod_Import_Logic)

- **SelectExcelFile** (file picker): starting directory = GetSetting("ImportFolderPath"). If that setting is **empty**, default = **CurrentProject.Path**. Trailing backslash added for FileDialog.  
  (RunDynamicImport does not open the file picker; the picker is in SelectExcelFile, so the fix is in SelectExcelFile.)

## Technical requirements met

- Late Binding for FSO: CreateObject("Scripting.FileSystemObject").
- Option Explicit in all modified modules.
- No Application.CompactRepair for the current database.
- Russian UI strings in mod_UI_Helpers built with ChrW to avoid ????? on different locales.

## Files modified

- **mod_UI_Helpers.bas** — RuSearchCaption, RuBackupPathUndefined, RuBackupSaved, RuBackupFailedLocked, RuBackupFailedGeneric, RuValidationLogCleared (ChrW).
- **uf_Search.cls** — m_strMode; btnClear reset (m_strMode, caption RuSearchCaption, enable buttons); Form_Load caption when not duplicates.
- **mod_Maintenance_Logic.bas** — CreateDatabaseBackup: Ru* messages, Error 70 handling; ClearValidationLog: RuValidationLogCleared.
- **mod_Import_Logic.bas** — SelectExcelFile: empty ImportFolderPath -> CurrentProject.Path.
