# Spec 011: Truly Dynamic Search & Export (Phase 8)

## Goal
Implement search and export logic that automatically adapts to the number of columns in `tbl_Personnel_Master`. No hardcoded field names; schema is analyzed at runtime.

## 1. Universal Dynamic Search (uf_Search)

### PerformSearch
- **Source of fields:** Use DAO to loop through `CurrentDb.TableDefs("tbl_Personnel_Master").Fields`.
- **Technical blacklist (excluded from search):**  
  `ID`, `SourceID`, `IsActive`, `LastUpdated`, `LogID`, `PersonUID_Raw`.
- **Searchable data types only:** Text, Memo, Long, Date (DAO: dbText, dbMemo, dbLong, dbDate).
- **WHERE clause:** Build dynamically:  
  `WHERE ([FieldName1] LIKE '*text*') OR ([FieldName2] LIKE '*text*') ...`  
  For Long (e.g. SourceID): use `CStr([FieldName]) LIKE '*text*'` or exclude per blacklist.  
  For Date: use `Format([FieldName], "yyyy-mm-dd") LIKE '*text*'` or CStr.
- **SELECT clause:** Build dynamically: prefer first columns PersonUID, FullName, RankName, PosName for list display; then remaining non-blacklist fields. ListBox shows first 4 columns; export uses full recordset.

### Quality
- Windows-1251 for VBA source.
- English comments only.
- English MsgBox for errors (e.g. "No employees found.").

---

## 2. Universal Dynamic Export (mod_Export_Logic.bas)

### ExportSearchToExcel(rs As DAO.Recordset) As Boolean
- **Input:** Recordset from current search (same query as list, without TOP 50).
- **Field list:** Loop through `rs.Fields`. Skip technical blacklist:  
  `ID`, `SourceID`, `IsActive`, `LastUpdated`, `LogID`, `PersonUID_Raw`.
- **Headers:** Use `mod_UI_Helpers.GetFieldFriendlyName(f.Name)`. If result equals cleaned field name (no friendly mapping), use field name with underscores replaced by spaces.
- **Data transfer:** Use `CopyFromRecordset` for efficiency (Range.CopyFromRecordset).
- **Excel formatting:** Freeze top row, bold headers, AutoFit all columns.
- **Binding:** Late binding only: `CreateObject("Excel.Application")`.
- **Cleanup:** `Set xlApp = Nothing`, close workbook, quit Excel.
- **MsgBox:** English: "Search results exported: N records" or error message.

### Quality
- English comments only.
- Proper error handling and object release.

---

## 3. uf_Search UI Update

- Add button `btnExportExcel` with English caption (e.g. "Export to Excel").
- On Click: build same SQL as current search (no TOP 50), open DAO Recordset, call `ExportSearchToExcel(rs)`, close Recordset. If list is empty, show message "No data to export."

---

## 4. Definition of Done

- [x] Dynamic search uses TableDef.Fields; no hardcoded field names.
- [x] Export uses rs.Fields and GetFieldFriendlyName; blacklist applied.
- [x] Row-by-row write (skipping blacklist); Freeze Panes, Bold row 1, AutoFit.
- [x] All VBA in Windows-1251; comments in English; MsgBox in English.
- [x] PROJECT_CONTEXT.md updated; spec archived to `.spec/archive/`.
