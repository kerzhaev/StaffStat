Attribute VB_Name = "mod_Export_Logic"
Option Compare Database
Option Explicit

' =============================================
' @module mod_Export_Logic
' @description Dynamic export of search results to Excel (late binding).
' Comments: English only. Encoding: Windows-1251.
' =============================================

' DAO DataTypeEnum (for searchable check): dbText=10, dbMemo=12, dbLong=4, dbDate=8
Private Const c_dbText  As Integer = 10
Private Const c_dbMemo  As Integer = 12
Private Const c_dbLong  As Integer = 4
Private Const c_dbDate  As Integer = 8

' Technical fields excluded from search and export
Private Const c_Blacklist As String = "ID|SourceID|IsActive|LastUpdated|LogID|PersonUID_Raw"

' =============================================
' @description Returns True if field name is in technical blacklist.
' @param strName [String] Field name
' =============================================
Public Function IsTechnicalField(ByVal strName As String) As Boolean
    Dim s As String
    s = Trim$(Nz(strName, ""))
    If s = "" Then
        IsTechnicalField = True
        Exit Function
    End If
    IsTechnicalField = ("|" & c_Blacklist & "|") Like "*|" & s & "|*"
End Function

' =============================================
' @description Exports recordset to new Excel workbook.
' Caller must pass recordset built with non-blacklist columns only (e.g. from BuildDynamicSearchSQL).
' Uses GetFieldFriendlyName for headers; CopyFromRecordset; Freeze, Bold, AutoFit.
' @param rs [DAO.Recordset] Recordset from search (caller must not close it)
' @return True if success
' =============================================
Public Function ExportSearchToExcel(ByVal rs As DAO.Recordset) As Boolean
    On Error GoTo ErrorHandler

    Dim xlApp     As Object
    Dim xlWb      As Object
    Dim xlWs      As Object
    Dim colCount  As Long
    Dim i         As Long
    Dim cap       As String
    Dim recCount  As Long

    If rs Is Nothing Then
        MsgBox "No recordset to export.", vbExclamation
        Exit Function
    End If

    colCount = rs.Fields.count
    If colCount = 0 Then
        MsgBox "No exportable columns.", vbExclamation
        Exit Function
    End If

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    xlApp.DisplayAlerts = False
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Worksheets(1)

    ' Headers row (friendly names)
    For i = 0 To colCount - 1
        cap = mod_UI_Helpers.GetFieldFriendlyName(rs.Fields(i).Name)
        If Trim$(cap) = "" Then cap = Replace(rs.Fields(i).Name, "_", " ")
        xlWs.Cells(1, i + 1).Value = cap
    Next i

    ' Data via CopyFromRecordset (recordset has only exportable columns from search SQL)
    If Not rs.EOF Then
        xlWs.Range("A2").CopyFromRecordset rs
    End If
    recCount = xlWs.UsedRange.Rows.count - 1
    If recCount < 0 Then recCount = 0

    ' Format: Freeze top row, Bold headers, AutoFit
    With xlWs.Range(xlWs.Cells(1, 1), xlWs.Cells(1, colCount))
        .Font.Bold = True
    End With
    xlWs.Rows(2).Select
    xlApp.ActiveWindow.FreezePanes = True
    xlWs.Range(xlWs.Cells(1, 1), xlWs.Cells(1, colCount)).Select
    xlWs.UsedRange.Columns.AutoFit

    MsgBox "Search results exported: " & recCount & " records.", vbInformation
    ExportSearchToExcel = True

Cleanup:
    On Error Resume Next
    If Not xlWs Is Nothing Then Set xlWs = Nothing
    If Not xlWb Is Nothing Then Set xlWb = Nothing
    If Not xlApp Is Nothing Then
        xlApp.DisplayAlerts = True
        Set xlApp = Nothing
    End If
    Exit Function

ErrorHandler:
    MsgBox "Export error: " & Err.Description, vbCritical
    ExportSearchToExcel = False
    GoTo Cleanup
End Function

' =============================================
' @description Exports change report by date range to Excel.
' @param dtStart [Date] Start date (inclusive)
' @param dtEnd [Date] End date (inclusive)
' @return True if success
' =============================================
Public Function ExportChangeReport(ByVal dtStart As Date, ByVal dtEnd As Date) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
    Dim recCount As Long

    Set db = CurrentDb

    strSQL = "SELECT " & _
             "H.ChangeDate AS [Date], " & _
             "M.FullName AS [FullName], " & _
             "H.PersonUID AS [PersonUID], " & _
             "H.FieldName AS [FieldName], " & _
             "H.OldValue AS [OldValue], " & _
             "H.NewValue AS [NewValue] " & _
             "FROM tbl_History_Log AS H " & _
             "LEFT JOIN tbl_Personnel_Master AS M ON H.PersonUID = M.PersonUID " & _
             "WHERE H.ChangeDate BETWEEN " & FormatDateLiteral(dtStart) & " AND " & FormatDateLiteral(dtEnd) & " " & _
             "ORDER BY M.FullName, H.PersonUID, H.ChangeDate;"

    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)

    If rs.EOF Then
        MsgBox "No changes found for selected period.", vbInformation
        GoTo Cleanup
    End If

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    xlApp.DisplayAlerts = False
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Worksheets(1)

    ' Headers (friendly names in Excel)
    xlWs.Cells(1, 1).Value = "Date"
    xlWs.Cells(1, 2).Value = "Full Name"
    xlWs.Cells(1, 3).Value = "ID"
    xlWs.Cells(1, 4).Value = "Field Changed"
    xlWs.Cells(1, 5).Value = "Old Value"
    xlWs.Cells(1, 6).Value = "New Value"

    ' Data via CopyFromRecordset
    xlWs.Range("A2").CopyFromRecordset rs

    rs.MoveLast
    recCount = rs.RecordCount
    rs.MoveFirst

    ' Format: Freeze top row, Bold headers, AutoFit
    With xlWs.Range(xlWs.Cells(1, 1), xlWs.Cells(1, 6))
        .Font.Bold = True
    End With
    xlWs.Rows(2).Select
    xlApp.ActiveWindow.FreezePanes = True
    xlWs.UsedRange.Columns.AutoFit

    ExportChangeReport = True

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not xlWs Is Nothing Then Set xlWs = Nothing
    If Not xlWb Is Nothing Then Set xlWb = Nothing
    If Not xlApp Is Nothing Then
        xlApp.DisplayAlerts = True
        Set xlApp = Nothing
    End If
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    MsgBox "Export error: " & Err.Description, vbCritical
    ExportChangeReport = False
    GoTo Cleanup
End Function

' =============================================
' @description Formats date for Access SQL literal.
' =============================================
Private Function FormatDateLiteral(ByVal dtValue As Date) As String
    FormatDateLiteral = "#" & Format(dtValue, "mm\/dd\/yyyy") & "#"
End Function
