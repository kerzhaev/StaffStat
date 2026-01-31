# Specification: 012 Period Change Report

## 1. Context & Goal
Provide a period-based change report by filtering `tbl_History_Log` by date range and exporting the result to Excel using late binding.

## 2. User Stories
*   [ ] As a manager, I want to filter changes by date range so I can review a specific period.
*   [ ] As an analyst, I want to export the filtered report to Excel so I can share it.
*   [ ] As a user, I want clear validation errors if my date range is invalid.

## 3. Functional Requirements
*   **Input:** Start date, end date, and export file path (from SaveAs dialog).
*   **Process:**
    * Validate that both dates are present and `StartDate <= EndDate`.
    * Build a DAO query/recordset from `tbl_History_Log` where `ChangeDate` is between the dates (inclusive).
    * Order by `ChangeDate`, then `PersonUID`.
    * Create Excel via `CreateObject("Excel.Application")` (late binding).
    * Create a workbook, write headers, write rows, autofit columns, save the file, and close Excel.
*   **Output:** An `.xlsx` file containing the filtered change report.
*   **Notes:** No data modifications; read-only export. SQL must be UPPERCASE.

## 4. UI/UX
*   Date range inputs (Start/End) near the history report action.
*   "Export Change Report" button.
*   Messages: validation errors, no data found, export success, export failure.

## 5. Constraints
*   Compatibility: Office 2010+.
*   Binding: Late binding for Excel automation.
*   Columns (in this order): `LogID`, `PersonUID`, `ChangeDate`, `FieldName`, `OldValue`, `NewValue`.
*   Error handling:
    * Invalid or missing dates.
    * Start date after end date.
    * No records found in the period.
    * Excel not available or cannot create instance.
    * File path invalid or save failed.
    * DAO query/recordset failure.
