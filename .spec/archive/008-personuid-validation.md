# Spec: Phase 7 — Task 2 — PersonUID Validation

**Status:** Implemented (2026-01-31)  
**Created:** 2026-01-31  
**Encoding:** UTF-8

---

## 1. Goal

Validate `PersonUID` (Личный номер) during synchronization so that invalid records are skipped and do not enter `tbl_Personnel_Master`. Invalid or malformed data must be logged for diagnostics; the user must be informed in Russian.

---

## 2. Validation Rule (Mask)

| Rule | Description |
|------|-------------|
| **Format** | 1 or 2 Cyrillic letters + hyphen + exactly 6 digits |
| **Examples (valid)** | `Ю-111111`, `АБ-123456`, `К-000001` |
| **Examples (invalid)** | `Ю-12`, `ABC-123456`, `Ю111111` (no hyphen), `Ю-1234567` (7 digits) |

**Regex pattern (VBScript.RegExp):**

- One letter: `[А-Яа-яЁё]-[0-9]{6}`
- Two letters: `[А-Яа-яЁё]{2}-[0-9]{6}`  
- Combined: `^[А-Яа-яЁё]{1,2}-[0-9]{6}$`

Cyrillic range in regex: `А-Яа-яЁё` (covers Russian letters; Win-1251 compatible).

---

## 3. Technical Approach

### 3.1. Validation implementation

- **Library:** `VBScript.RegExp` (Late Binding: `CreateObject("VBScript.RegExp")`).
- **Options:** `Global = False`, `IgnoreCase = True` (optional; pattern can be case-insensitive for letters).
- **Method:** `Test(strUID)` after `Pattern` and options are set.
- **Trim:** Validate `Trim(strUID)` so spaces do not cause false negatives.

### 3.2. Where to validate

- **Place:** Inside `mod_Analysis_Logic.SyncBufferToMaster`, in the main `Do While Not rsBuffer.EOF` loop.
- **When:** Immediately after `strUID = Nz(rsBuffer!PersonUID_Raw, "")`.
- **If empty:** Keep current behavior: skip (no insert, no log for empty).
- **If not empty and invalid:** Skip the record, log one warning per invalid UID, increment a "skipped" counter, then `rsBuffer.MoveNext` (continue loop).
- **If valid:** Proceed with existing logic (new employee or existing).

### 3.3. Logging

- **API:** `mod_App_Logger.LogInfo` (no `LogEvent` in current module; use `LogInfo` for warnings).
- **Call:** `LogInfo(englishMessage, "PersonUID_Validation")`.
- **Message (English):** e.g. `"Skipped invalid PersonUID: [value]"` or `"Invalid PersonUID format, skipped: [value]"`.
- **User message (Russian):** Add to the existing completion `MsgBox` at the end of `SyncBufferToMaster`, e.g. `"Пропущено (неверный формат личного номера): N"` when skipped count > 0.

---

## 4. Files to Create or Modify

| File | Action |
|------|--------|
| `StaffState.accdb.src/modules/mod_Validation_Logic.bas` | **Create.** Contains `IsValidPersonUID`. |
| `StaffState.accdb.src/modules/mod_Analysis_Logic.bas` | **Modify.** Call validation in loop; skip and log; add skipped count and Russian summary in MsgBox. |
| `.spec/PROJECT_CONTEXT.md` | **Update** after implementation (history, Phase 7 Task 2). |
| `.spec/008-personuid-validation.md` | **Move** to `.spec/archive/` after completion. |

---

## 5. Module: mod_Validation_Logic.bas

### 5.1. Requirements

- `Option Explicit`.
- Error handling: `On Error GoTo ErrorHandler` in the public function; release any RegExp object and exit safely.
- Comments and JSDoc: **English only**.
- File encoding: **Windows-1251** (VBA source).

### 5.2. Function signature

```vba
' =============================================
' @description Returns True if strUID matches mask: 1 or 2 Cyrillic letters,
'              hyphen, exactly 6 digits. Uses VBScript.RegExp (Late Binding).
' @param strUID [String] PersonUID to validate (e.g. from PersonUID_Raw).
' @return [Boolean] True if valid format, False otherwise.
' =============================================
Public Function IsValidPersonUID(ByVal strUID As String) As Boolean
```

### 5.3. Logic (pseudocode)

1. Trim `strUID`. If empty, return False.
2. CreateObject("VBScript.RegExp").
3. Set `.Pattern = "^[А-Яа-яЁё]{1,2}-[0-9]{6}$"` (string in Win-1251).
4. Set `.IgnoreCase = True` (optional).
5. `IsValidPersonUID = regExp.Test(strUID)`.
6. Set RegExp object to Nothing in exit path and in ErrorHandler.

---

## 6. Changes in mod_Analysis_Logic.SyncBufferToMaster

### 6.1. New variable

- `Dim iSkipped As Long` (initialized to 0 at start of procedure).

### 6.2. In the loop (after `strUID = Nz(rsBuffer!PersonUID_Raw, "")`)

- If `strUID = ""`: keep current behavior (no processing, no log).
- Else:  
  - If **Not** `IsValidPersonUID(strUID)` Then  
    - `LogInfo "Skipped invalid PersonUID: " & strUID, "PersonUID_Validation"`  
    - `iSkipped = iSkipped + 1`  
    - `rsBuffer.MoveNext`  
    - `Do While` continues (next record).  
  - Else: run existing logic (FindFirst, AddNew/Edit, CopyAllFields, LogChange, etc.).

### 6.3. ExitHandler / completion MsgBox

- Extend the existing "Синхронизация завершена!" message to include a line when `iSkipped > 0`, e.g.:
  - `"Пропущено (неверный формат личного номера): " & iSkipped`

---

## 7. Quality Checklist

- [x] `Option Explicit` in new module and unchanged in modified module.
- [x] Every procedure has `On Error GoTo ErrorHandler` and safe cleanup (Set … = Nothing).
- [x] All VBA comments and JSDoc in **English**.
- [x] VBA files saved in **Windows-1251**.
- [x] No Cyrillic in comments; Cyrillic only in UI strings (e.g. MsgBox) and in RegExp pattern string.

---

## 8. Acceptance Criteria

- [x] `IsValidPersonUID("Ю-111111")` = True, `IsValidPersonUID("АБ-123456")` = True.
- [x] `IsValidPersonUID("Ю-12")` = False, `IsValidPersonUID("ABC-123456")` = False, `IsValidPersonUID("")` = False.
- [x] Invalid records are skipped in `SyncBufferToMaster` and not inserted/updated in Master.
- [x] For each skipped record, `LogInfo` is called with English message and Source `"PersonUID_Validation"`.
- [x] Final MsgBox shows Russian text for skipped count when `iSkipped > 0`.
- [x] Loop continues to next record after skip (no crash, no duplicate processing).

---

## 9. References

- `.cursorrules` — Encoding (Win-1251 VBA, English comments), Late Binding, Option Explicit, Error Handling.
- `.spec/PROJECT_CONTEXT.md` — Current state, Phase 7 Task 2.
- `StaffState.accdb.src/.spec/007-performance-improvements.md` — Task 2 description (validation rule).
