# SPEC: Phase 7 - Performance & Usability Improvements

**Status:** ✅ APPROVED (Ready for Implementation)  
**Author:** Junior Developer (AI Assistant)  
**Tech Lead:** @kerga  
**Created:** 2026-01-23

---

## 1. OVERVIEW

This specification describes a set of incremental improvements to the StaffState system focused on:
- **Performance optimization** (indexes for large datasets)
- **Data quality** (validation and duplicate detection)
- **Reporting capabilities** (export and analytics)
- **Long-term maintenance** (archiving)

All improvements are **backward compatible** and do not break existing functionality.

---

## 2. BUSINESS JUSTIFICATION

Current system works well for <1000 records, but performance degrades with growth:
- Import of 5000 records takes 10-15 minutes (linear scan on each `FindFirst`)
- Search in `uf_Search` takes 2-3 seconds with 10,000+ employees
- No protection against importing duplicate files
- No analytical reports for management

**ROI:** These improvements will ensure system scalability to 50,000+ employees.

### 2.1. Database Size Projections (for Tech Lead's case: 20,000 employees, monthly imports)

**Assumptions:**
- 20,000 employees in Master table
- Monthly import (1x per month)
- ~10% of employees have changes per month (2,000 people)
- Average 2 fields change per person (4,000 history records/month)

**Size Calculation:**

| Component | Size (Year 1) | Size (Year 5) | Size (Year 10) |
|-----------|---------------|---------------|----------------|
| `tbl_Personnel_Master` | 10 MB | 10 MB | 10 MB |
| `tbl_History_Log` | 15 MB | 72 MB | 144 MB |
| `tbl_Import_Buffer` | 0 MB* | 0 MB* | 0 MB* |
| Indexes | 3 MB | 12 MB | 23 MB |
| **TOTAL** | **28 MB** | **94 MB** | **177 MB** |

*Buffer is cleared after each import

**Access 2010+ Limit:** 2 GB (2,000 MB)

**Conclusion:** 
- ✅ **Without archiving:** Database will reach ~180 MB after 10 years (only 9% of limit)
- ✅ **Archiving recommended after:** ~15-20 years or when History exceeds 500 MB
- ✅ **No size concerns for next 10+ years** of operation

**Worst-case scenario** (30% change rate, 3 fields per person):
- Year 5: ~350 MB (still 17% of limit)
- Year 10: ~700 MB (still 35% of limit)

---

## 3. PRIORITY LEVELS

### 🔴 PRIORITY A: CRITICAL (implement first)
- **Task 1:** Database Indexes (Performance)
- **Task 2:** PersonUID Validation (Data Quality)

### 🟡 PRIORITY B: IMPORTANT (user convenience)
- **Task 3:** Changes Report (Analytics)
- **Task 4:** Duplicate Import Detection (Safety)
- **Task 5:** Search Results Export (User Feature)

### 🟢 PRIORITY C: OPTIONAL (for large-scale deployment)
- **Task 6:** History Archiving (Database Optimization)
- **Task 7:** Search Column Configuration (Flexibility)

---

## 4. DETAILED TASK BREAKDOWN

---

### 🔴 TASK 1: Performance Indexes

#### 4.1.1. Problem Statement
- `mod_Analysis_Logic.SyncBufferToMaster()` calls `FindFirst` for each imported record
- Without index on `PersonUID`, Access performs full table scan (O(N) complexity)
- Example: 500 imports × 5000 Master records = 2,500,000 comparisons

#### 4.1.2. Solution
Create indexes on frequently queried fields:
1. `tbl_Personnel_Master.PersonUID` (UNIQUE)
2. `tbl_Personnel_Master.FullName`
3. `tbl_History_Log.PersonUID`
4. `tbl_History_Log.ChangeDate`

#### 4.1.3. Files to Modify
- `modules/mod_App_Init.bas`

#### 4.1.4. Implementation Plan
```
Step 1: Add function CreatePerformanceIndexes()
Step 2: Create 4 indexes using DAO.Index API
Step 3: Add error handling (On Error Resume Next if index exists)
Step 4: Call from InitDatabaseStructure() or create separate button
Step 5: Test with 10,000 records (measure speed improvement)
```

#### 4.1.5. Expected Impact
- Import time: 15 min → 30 sec (30x faster)
- Search time: 2 sec → 0.05 sec (40x faster)

#### 4.1.6. Risks
- Index creation takes ~30 seconds on first run (one-time cost)
- Indexes increase DB size by ~10-15%

---

### 🔴 TASK 2: PersonUID Validation

#### 4.2.1. Problem Statement
- No validation of `PersonUID` format during import
- Corrupted data (e.g., "Ю-12" or "ABC-123456") can enter Master table
- Breaks search and history tracking

#### 4.2.2. Solution
Implement format validation with flexible prefix:
- **Format 1:** `X-NNNNNN` (1 Cyrillic letter + dash + 6 digits) → length = 8
- **Format 2:** `XX-NNNNNN` (2 Cyrillic letters + dash + 6 digits) → length = 9
- Examples: `Ю-110111`, `ЮК-123456`, `АВ-999888`

#### 4.2.3. Files to Modify
- `modules/mod_Import_Logic.bas`

#### 4.2.4. Implementation Plan
```
Step 1: Create Private Function ValidatePersonUID(strUID As String) As Boolean
        - Check length IN (8, 9)
        - Find dash position (InStr)
        - Check dash at position 2 OR 3
        - Check characters before dash are Cyrillic letters (Asc range 192-255)
        - Check 6 characters after dash are digits (IsNumeric)
Step 2: Modify SQL in RunDynamicImport():
        - Add WHERE clause: "WHERE Len([Личный номер]) IN (8, 9)"
        - Additional filter: "[Личный номер] LIKE '*-######'"
Step 3: Log rejected records to Debug.Print (with reason)
Step 4: Test with corrupted Excel file + both formats
```

#### 4.2.5. Expected Impact
- Prevents garbage data from entering Master
- Supports both single and double-letter prefixes
- Easier debugging (clear rejection log with reasons)

#### 4.2.6. Risks
- If non-standard format exists (e.g., 3 letters or 5 digits), will be rejected
- Need to verify all existing data conforms to this pattern

---

### 🟡 TASK 3: Changes Report

#### 4.3.1. Problem Statement
- Management needs periodic reports: "What changed this week?"
- Current solution: manually filter `tbl_History_Log` (not user-friendly)

#### 4.3.2. Solution
Create Excel export function with date range filter.

#### 4.3.3. Files to Create/Modify
- **NEW:** `modules/mod_Reports.bas`
- `forms/uf_Dashboard.cls` (add button)

#### 4.3.4. Implementation Plan
```
Step 1: Create mod_Reports.bas
Step 2: Public Sub ExportChangesReport(dtStart As Date, dtEnd As Date)
        - SQL: JOIN History with Master (get FullName)
        - Late binding Excel: CreateObject("Excel.Application")
        - Export to Excel with formatted headers
Step 3: Add button "Отчет" to uf_Dashboard
Step 4: Add date picker dialog (or hardcode "last 7 days")
Step 5: Test with sample data
```

#### 4.3.5. Expected Impact
- Management gets weekly/monthly change summaries
- Excel format allows further analysis (pivot tables, etc.)

---

### 🟡 TASK 4: Duplicate Import Detection

#### 4.4.1. Problem Statement
- User can accidentally import same Excel file twice
- Creates duplicate history entries (noise in "Time Machine")

#### 4.4.2. Solution
Check file modification date against `tbl_Import_Meta` before import.

#### 4.4.3. Files to Modify
- `modules/mod_Import_Logic.bas`

#### 4.4.4. Implementation Plan
```
Step 1: Create Private Function CheckDuplicateImport(strFilePath) As Boolean
        - Get file date via GetFileModificationDate()
        - Query tbl_Import_Meta for same ExportFileDate
        - Return True if duplicate found
Step 2: In ImportExcelData():
        - Call CheckDuplicateImport() after file selection
        - Show MsgBox: "File already imported on YYYY-MM-DD. Continue?"
        - Allow user override (Yes/No)
Step 3: (Optional) Change tbl_Import_Meta to multi-row log instead of single-row
Step 4: Test with same file imported twice
```

#### 4.4.5. Expected Impact
- Prevents accidental duplicate imports
- Full import history log (if implemented as multi-row)

---

### 🟡 TASK 5: Search Results Export

#### 4.5.1. Problem Statement
- User searches for "all captains" (100+ results)
- No way to save/share the list (only on-screen view)

#### 4.5.2. Solution
Add "Export to Excel" button on `uf_Search` form.

#### 4.5.3. Files to Modify
- `forms/uf_Search.bas` (add button)
- `forms/uf_Search.cls` (add event handler)

#### 4.5.4. Implementation Plan
```
Step 1: Add button "btnExport" to uf_Search form
Step 2: Create Private Sub btnExport_Click()
        - Check if lstResults has data
        - Create Excel via Late Binding
        - Export headers: "Личный номер | ФИО | Должность"
        - Loop through ListBox.Column(i, row)
        - Format (autofit columns, bold headers)
Step 3: Test with 500+ search results
```

#### 4.5.5. Expected Impact
- Users can create reports from search results
- Useful for ad-hoc queries

---

### 🟢 TASK 6: History Archiving

#### 4.6.1. Problem Statement
- After 5 years, `tbl_History_Log` may contain 500,000+ records
- Slows down "Time Machine" queries

#### 4.6.2. Solution
Move records older than N months to archive table.

#### 4.6.3. Files to Modify
- `modules/mod_App_Init.bas` (create archive table)
- `modules/mod_App_Init.bas` (add ArchiveOldHistory function)
- `forms/uf_Dashboard.cls` (add button)

#### 4.6.4. Implementation Plan
```
Step 1: In InitDatabaseStructure():
        - Create tbl_History_Archive (same structure as tbl_History_Log)
Step 2: Create Public Sub ArchiveOldHistory(iMonths As Long)
        - Calculate cutoff date: DateAdd("m", -iMonths, Date)
        - INSERT INTO tbl_History_Archive SELECT FROM tbl_History_Log WHERE ...
        - DELETE FROM tbl_History_Log WHERE ChangeDate < cutoff
Step 3: Add button "Архивация" to Dashboard (or system settings)
Step 4: Test with 50,000 records
```

#### 4.6.5. Expected Impact
- Keeps active history table small (<10,000 records)
- Preserves old data for audits

---

### 🟢 TASK 7: Search Column Configuration

#### 4.7.1. Problem Statement
- Different users need different fields in search results
- HR needs: FullName, Rank, Status
- Commanders need: FullName, Position, Unit

#### 4.7.2. Solution
Add checkboxes to toggle displayed columns.

#### 4.7.3. Files to Modify
- `forms/uf_Search.bas` (add checkboxes)
- `forms/uf_Search.cls` (dynamic SQL generation)

#### 4.7.4. Implementation Plan
```
Step 1: Add checkboxes: chkShowRank, chkShowStatus, chkShowPosition
Step 2: Create Private Sub BuildSearchSQL(strFilter As String) As String
        - Base: SELECT PersonUID, FullName
        - If chkShowRank: Add ", RankName"
        - If chkShowStatus: Add ", WorkStatus"
        - Dynamically set lstResults.ColumnCount
Step 3: Call BuildSearchSQL() from PerformSearch()
Step 4: (Optional) Save preferences to local table
Step 5: Test with different checkbox combinations
```

#### 4.7.5. Expected Impact
- Flexible UI for different user roles
- Reduces visual clutter

---

## 5. IMPLEMENTATION SEQUENCE

### Session 1: Critical Performance (30 min)
```
├─ Task 1: Create Indexes (15 min)
│  ├─ Modify mod_App_Init.bas
│  ├─ Test with 10,000 records
│  └─ Measure speed improvement
└─ Task 2: Validation (15 min)
   ├─ Modify mod_Import_Logic.bas
   └─ Test with corrupted Excel
```

### Session 2: User Safety (30 min)
```
├─ Task 4: Duplicate Detection (15 min)
│  ├─ Modify mod_Import_Logic.bas
│  └─ Test with same file twice
└─ Task 5: Search Export (15 min)
   ├─ Modify uf_Search form
   └─ Test with 500+ results
```

### Session 3: Analytics (30 min)
```
└─ Task 3: Changes Report (30 min)
   ├─ Create mod_Reports.bas
   ├─ Modify uf_Dashboard
   └─ Test with 1-month date range
```

### Session 4: Optional Features (45 min)
```
├─ Task 6: Archiving (25 min)
│  ├─ Modify mod_App_Init.bas
│  └─ Test with 50,000 records
└─ Task 7: Column Config (20 min)
   ├─ Modify uf_Search form
   └─ Test with different roles
```

---

## 6. TESTING STRATEGY

### 6.1. Performance Testing
- Generate fake dataset: 10,000 employees, 50,000 history records
- Measure `SyncBufferToMaster()` time before/after indexes
- Measure search time in `uf_Search` before/after indexes

### 6.2. Data Quality Testing
- Create corrupted Excel with invalid PersonUIDs
- Verify rejection and logging

### 6.3. Edge Cases
- Import file with 0 records
- Import file with duplicate PersonUIDs within same file
- Search with Cyrillic and Latin characters

---

## 7. ROLLBACK PLAN

All changes are additive (new functions, new indexes). If issues arise:
1. Indexes can be dropped via DAO without data loss
2. New modules can be excluded from import
3. Forms retain old versions in git history

---

## 8. TECH LEAD DECISIONS ✅

**Approved on 2026-01-23**

1. **Task 4 (Duplicates):**
   - ✅ **DECISION:** Keep `tbl_Import_Meta` as single-row (current design)
   - **Reason:** Simplicity. Only last import metadata is needed for change tracking.

2. **Task 6 (Archiving):**
   - ✅ **DECISION:** Implement Option A (archive to separate table in same DB)
   - **Reason:** Based on Tech Lead's data (20k employees, monthly imports), archiving won't be needed for 10+ years.
   - **Priority:** LOW (can be postponed to Phase 8+)
   
   **Option A Selected: Archive to separate table in same DB**
   - **Pros:**
     - Simple implementation (INSERT INTO ... SELECT)
     - Single database file (easier backup)
     - Can query archived data via JOIN if needed
   - **Cons:**
     - Database size grows indefinitely (but acceptable for 10+ years)
   
   **When to implement:**
   - When `tbl_History_Log` reaches 500,000+ records
   - Or when "Time Machine" queries take >2 seconds
   - Based on current data: ~Year 15-20 of operation

3. **Task 7 (Column Config):**
   - ✅ **DECISION:** Search fields priority:
     1. PersonUID (Личный номер)
     2. SourceID (Табельный номер - "Лицо")
     3. FullName (ФИО)
     4. BirthDate (Дата рождения)
     5. (Future) Military Unit (Воинская часть - not yet implemented)
   - **Implementation:** Add checkboxes for optional fields (BirthDate, Unit when added)

4. **General:**
   - ✅ **DECISION:** Create separate button "Создать индексы" (manual trigger)
   - **Reason:** Educational project, user wants control over DB structure changes
   - **Location:** Add button to uf_Dashboard or create new system settings form

---

## 9. ACCEPTANCE CRITERIA

### Task 1 (Indexes) - DONE when:
- [ ] 4 indexes created successfully
- [ ] `SyncBufferToMaster()` runs 10x+ faster on 5000+ records
- [ ] Search in `uf_Search` is instant (<0.1 sec)

### Task 2 (Validation) - DONE when:
- [ ] Invalid PersonUIDs are rejected during import
- [ ] Rejection log appears in Debug window
- [ ] Valid data imports without errors

### Task 3 (Report) - DONE when:
- [ ] Button "Отчет" exists on Dashboard
- [ ] Excel file exports with date range filter
- [ ] Headers formatted, data readable

### Task 4 (Duplicates) - DONE when:
- [ ] Warning appears when importing same file twice
- [ ] User can override with "Yes"
- [ ] Import metadata updated correctly

### Task 5 (Export) - DONE when:
- [ ] Button "Экспорт" exists on uf_Search
- [ ] Excel file contains all search results
- [ ] Format matches search display

### Task 6 (Archiving) - DONE when:
- [ ] Archive table created
- [ ] Old records moved successfully
- [ ] Active history table <10,000 records

### Task 7 (Config) - DONE when:
- [ ] Checkboxes control column visibility
- [ ] ListBox adjusts dynamically
- [ ] No errors with different combinations

---

## 10. POST-IMPLEMENTATION

After each session:
1. Update `.spec/PROJECT_CONTEXT.md` (Section 7: History)
2. Test with production-like data
3. Git commit with descriptive message
4. Demo to Tech Lead for approval

---

## 11. REFERENCES

- Constitution: `.cursorrules` (Section 5: WORKFLOW)
- Context: `.spec/PROJECT_CONTEXT.md`
- Current Modules:
  - `mod_App_Init.bas`
  - `mod_Import_Logic.bas`
  - `mod_Analysis_Logic.bas`
- Current Forms:
  - `uf_Dashboard`
  - `uf_Search`
  - `uf_PersonCard`

---

**AWAITING TECH LEAD APPROVAL TO PROCEED** ✅
