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
