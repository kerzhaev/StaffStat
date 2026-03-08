Attribute VB_Name = "mod_Export_Logic"
Option Compare Database
Option Explicit

' =============================================
' @module mod_Export_Logic
' @description Dynamic export of search results to Excel (Late Binding).
'              Excel: CreateObject("Excel.Application"); no reference to Excel Object Library.
' @note 100% English version. Safe for modern IDEs, Git and AI tools.
' =============================================

' DAO DataTypeEnum (for searchable check): dbText=10, dbMemo=12, dbLong=4, dbDate=8
Private Const c_dbText  As Integer = 10
Private Const c_dbMemo  As Integer = 12
Private Const c_dbLong  As Integer = 4
Private Const c_dbDate  As Integer = 8

' Technical fields excluded from search and export
Private Const c_Blacklist As String = "ID|SourceID|IsActive|LastUpdated|LogID"

' Excel constants (Late Binding)
Private Const xlCenter As Long = -4108
Private Const xlLeft As Long = -4131
Private Const xlTop As Long = -4160
Private Const xlContinuous As Long = 1
Private Const xlUp As Long = -4162
Private Const xlOpenXMLWorkbook As Long = 51

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
' @description Returns header for field from tbl_Import_Mapping.
'              Uses pre-loaded dictionary (no Static cache - fresh data each export).
' @param strFieldName [String] English field name
' @param dictMapping [Object] Scripting.Dictionary: TargetField -> ExcelHeader (caller loads from DB)
' @return [String] Header or empty string if not found
' =============================================
Private Function GetHeaderFromMapping(ByVal strFieldName As String, ByVal dictMapping As Object) As String
    On Error GoTo ErrorHandler

    If dictMapping Is Nothing Then
        GetHeaderFromMapping = ""
        Exit Function
    End If
    If dictMapping.Exists(strFieldName) Then
        GetHeaderFromMapping = dictMapping(strFieldName)
    Else
        GetHeaderFromMapping = ""
    End If

    Exit Function
ErrorHandler:
    Debug.Print "GetHeaderFromMapping error: " & Err.Description
    GetHeaderFromMapping = ""
End Function

' =============================================
' @description Loads tbl_Import_Mapping (ProfileID=1) into a Scripting.Dictionary.
'              Called once per export so manual changes to tbl_Import_Mapping apply immediately.
' @return [Object] Dictionary: TargetField -> ExcelHeader, or Nothing on error
' =============================================
Private Function LoadMappingDictionary() As Object
    On Error GoTo ErrorHandler

    Dim dict As Object
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String

    Set dict = CreateObject("Scripting.Dictionary")
    Set db = CurrentDb
    strSQL = "SELECT TargetField, ExcelHeader FROM tbl_Import_Mapping WHERE ProfileID = 1;"
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)

    Do While Not rs.EOF
        dict(rs!targetField.value) = Nz(rs!ExcelHeader.value, "")
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set db = Nothing

    Debug.Print "Mapping Dictionary loaded with " & dict.count & " keys"
    Set LoadMappingDictionary = dict
    Exit Function

ErrorHandler:
    Debug.Print "LoadMappingDictionary error: " & Err.Description
    Set LoadMappingDictionary = Nothing
End Function

' =============================================
' @description Exports full search results to Excel (all records, no TOP 50).
'              Uses uf_Search.GetSearchSQL(False) for query.
'              Headers from tbl_Import_Mapping, fallback to GetFieldFriendlyName.
'              Late Binding Excel with formatting.
' @return True if success
' =============================================
Public Function ExportFullSearchToExcel() As Boolean
    Dim result As Object

    Set result = ExportFullSearchToExcelResult()
    ExportFullSearchToExcel = CBool(result("Success")) And CStr(Nz(result("Status"), "")) = "SUCCESS"
    Set result = Nothing
End Function

Public Function ExportFullSearchToExcelResult() As Object
    On Error GoTo ErrorHandler

    Dim result As Object
    Dim ufSearch As Object
    Dim strSQL As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
    Dim colCount As Long
    Dim i As Long
    Dim strHeader As String
    Dim recCount As Long
    Dim lastRow As Long
    Dim rngHeaders As Object
    Dim rngData As Object
    Dim strFieldName As String
    Dim dictMap As Object
    Dim exportDir As String
    Dim fullPath As String

    Set result = CreateExportResult()

    ' Get SQL from uf_Search (without TOP 50)
    strSQL = Forms!uf_Search.GetSearchSQL(bAddTop50:=False)
    If strSQL = "" Then
        result("Success") = True
        result("Status") = "NO_DATA"
        result("Message") = "No data to export."
        GoTo Cleanup
    End If

    ' Resolve export folder: CurrentProject.Path\Exports (create if missing)
    exportDir = CurrentProject.Path & "\Exports"
    If Len(CurrentProject.Path) = 0 Then
        result("Status") = "VALIDATION_ERROR"
        result("ErrorMessage") = "Database path not found. Please open the database from a trusted folder."
        result("Message") = CStr(result("ErrorMessage"))
        GoTo Cleanup
    End If
    If Dir(exportDir, vbDirectory) = "" Then
        MkDir exportDir
        Debug.Print "Export: created folder " & exportDir
    End If

    ' Open recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If rs.EOF Then
        result("Success") = True
        result("Status") = "NO_DATA"
        result("Message") = "No data to export."
        GoTo Cleanup
    End If

    ' Excel: Late Binding
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    xlApp.DisplayAlerts = False
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Worksheets(1)

    colCount = rs.Fields.count

    ' Load mapping once per export (fresh from tbl_Import_Mapping)
    Set dictMap = LoadMappingDictionary()

    ' Headers row: headers from mapping
    For i = 0 To colCount - 1
        strFieldName = rs.Fields(i).Name
        strHeader = GetHeaderFromMapping(strFieldName, dictMap)
        If Trim$(strHeader) = "" Then
            ' Fallback 1: GetFieldFriendlyName from mod_UI_Helpers
            strHeader = mod_UI_Helpers.GetFieldFriendlyName(strFieldName)
            If Trim$(strHeader) = "" Or strFieldName = strHeader Then
                ' Fallback 2: Replace underscores with spaces
                strHeader = Replace(strFieldName, "_", " ")
            End If
        End If
        xlWs.Cells(1, i + 1).value = strHeader
    Next i

    ' Data via CopyFromRecordset
    xlWs.Range("A2").CopyFromRecordset rs
    lastRow = xlWs.Cells(xlWs.Rows.count, 1).End(xlUp).Row
    If lastRow < 2 Then lastRow = 2
    recCount = lastRow - 1

    ' Format: Bold headers, AutoFit, FreezePanes, AutoFilter, Borders
    Set rngHeaders = xlWs.Range(xlWs.Cells(1, 1), xlWs.Cells(1, colCount))
    With rngHeaders.Font
        .Bold = True
    End With
    xlWs.Range("A2").Select
    xlApp.ActiveWindow.FreezePanes = True

    Set rngData = xlWs.Range(xlWs.Cells(1, 1), xlWs.Cells(lastRow, colCount))
    With rngData
        .AutoFilter
        .Borders.LineStyle = xlContinuous
    End With

    xlWs.UsedRange.Columns.AutoFit

    ' Format date columns to dd.mm.yyyy
    For i = 0 To colCount - 1
        If rs.Fields(i).Type = c_dbDate Then
            xlWs.Range(xlWs.Cells(2, i + 1), xlWs.Cells(lastRow, i + 1)).NumberFormat = "dd.mm.yyyy"
        End If
    Next i

    ' Save to Exports folder
    fullPath = exportDir & "\SearchExport_" & Format(Now(), "yyyymmdd_hhnnss") & ".xlsx"
    xlWb.SaveAs fullPath, xlOpenXMLWorkbook
    Debug.Print "Export saved: " & fullPath

    result("Success") = True
    result("Status") = "SUCCESS"
    result("RecordCount") = recCount
    result("ExportPath") = fullPath
    result("Message") = mod_UI_Helpers.GetMsgExportCompleted(recCount)

Cleanup:
    On Error Resume Next
    If Not dictMap Is Nothing Then Set dictMap = Nothing
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    If Not rngData Is Nothing Then Set rngData = Nothing
    If Not rngHeaders Is Nothing Then Set rngHeaders = Nothing
    If Not xlWs Is Nothing Then Set xlWs = Nothing
    If Not xlWb Is Nothing Then Set xlWb = Nothing
    If Not xlApp Is Nothing Then
        xlApp.DisplayAlerts = True
        Set xlApp = Nothing
    End If
    Set ExportFullSearchToExcelResult = result
    Exit Function

ErrorHandler:
    Debug.Print "ExportFullSearchToExcel error: " & Err.Description & " (" & Err.Number & ")"
    result("Success") = False
    result("Status") = "ERROR"
    result("ErrorNumber") = Err.Number
    result("ErrorMessage") = "Export error: " & Err.Description
    result("Message") = CStr(result("ErrorMessage"))
    GoTo Cleanup
End Function

' =============================================
' @description Exports recordset to new Excel workbook (Legacy).
' @param rs [DAO.Recordset] Recordset from search
' @return True if success
' =============================================
Public Function ExportSearchToExcel(ByVal rs As DAO.Recordset) As Boolean
    Dim result As Object

    Set result = ExportSearchToExcelResult(rs)
    ExportSearchToExcel = CBool(result("Success")) And CStr(Nz(result("Status"), "")) = "SUCCESS"
    Set result = Nothing
End Function

Public Function ExportSearchToExcelResult(ByVal rs As DAO.Recordset) As Object
    On Error GoTo ErrorHandler

    Dim result As Object
    Dim xlApp     As Object
    Dim xlWb      As Object
    Dim xlWs      As Object
    Dim colCount  As Long
    Dim i         As Long
    Dim cap       As String
    Dim recCount  As Long

    Set result = CreateExportResult()

    If rs Is Nothing Then
        result("Status") = "VALIDATION_ERROR"
        result("ErrorMessage") = "No recordset to export."
        result("Message") = CStr(result("ErrorMessage"))
        GoTo Cleanup
    End If

    colCount = rs.Fields.count

    If colCount = 0 Then
        result("Status") = "VALIDATION_ERROR"
        result("ErrorMessage") = "No exportable columns."
        result("Message") = CStr(result("ErrorMessage"))
        GoTo Cleanup
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
        xlWs.Cells(1, i + 1).value = cap
    Next i

    ' Data via CopyFromRecordset
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

    result("Success") = True
    result("Status") = "SUCCESS"
    result("RecordCount") = recCount
    result("Message") = "Search results exported: " & recCount & " records."

Cleanup:
    On Error Resume Next
    If Not xlWs Is Nothing Then Set xlWs = Nothing
    If Not xlWb Is Nothing Then Set xlWb = Nothing
    If Not xlApp Is Nothing Then
        xlApp.DisplayAlerts = True
        Set xlApp = Nothing
    End If
    Set ExportSearchToExcelResult = result
    Exit Function

ErrorHandler:
    Debug.Print "ExportSearchToExcel error: " & Err.Description & " (" & Err.Number & ")"
    result("Success") = False
    result("Status") = "ERROR"
    result("ErrorNumber") = Err.Number
    result("ErrorMessage") = "Export error: " & Err.Description
    result("Message") = CStr(result("ErrorMessage"))
    GoTo Cleanup
End Function

' =============================================
' @description Exports change report by date range to Excel.
' @param dtStart [Date] Start date (inclusive)
' @param dtEnd [Date] End date (inclusive)
' @return True if success
' =============================================
Public Function ExportChangeReport(ByVal dtStart As Date, ByVal dtEnd As Date) As Boolean
    Dim result As Object

    Set result = ExportChangeReportResult(dtStart, dtEnd)
    ExportChangeReport = CBool(result("Success")) And CStr(Nz(result("Status"), "")) = "SUCCESS"
    Set result = Nothing
End Function

Public Function ExportChangeReportResult(ByVal dtStart As Date, ByVal dtEnd As Date) As Object
    On Error GoTo ErrorHandler

    Dim result As Object
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
    Dim recCount As Long
    Dim strOrgName As String
    Dim lngHeaderRow As Long

    Set result = CreateExportResult()
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
        result("Success") = True
        result("Status") = "NO_DATA"
        result("Message") = "No changes found for selected period."
        GoTo Cleanup
    End If

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    xlApp.DisplayAlerts = False
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Worksheets(1)

    ' Report header from OrganizationName setting
    strOrgName = Trim$(Nz(mod_Maintenance_Logic.GetSetting("OrganizationName", ""), ""))
    If Len(strOrgName) > 0 Then
        xlWs.Cells(1, 1).value = strOrgName
        xlWs.Range("A1:F1").Merge
        xlWs.Range("A1").Font.Bold = True
    End If

    ' Headers (friendly names in Excel)
    lngHeaderRow = IIf(Len(strOrgName) > 0, 2, 1)
    xlWs.Cells(lngHeaderRow, 1).value = "Date"
    xlWs.Cells(lngHeaderRow, 2).value = "Full Name"
    xlWs.Cells(lngHeaderRow, 3).value = "ID"
    xlWs.Cells(lngHeaderRow, 4).value = "Field Changed"
    xlWs.Cells(lngHeaderRow, 5).value = "Old Value"
    xlWs.Cells(lngHeaderRow, 6).value = "New Value"

    ' Data via CopyFromRecordset
    xlWs.Range("A" & (lngHeaderRow + 1)).CopyFromRecordset rs

    rs.MoveLast
    recCount = rs.RecordCount
    rs.MoveFirst

    ' Format: Bold headers, FreezePanes, AutoFit
    With xlWs.Range(xlWs.Cells(lngHeaderRow, 1), xlWs.Cells(lngHeaderRow, 6))
        .Font.Bold = True
    End With
    xlWs.Rows(lngHeaderRow + 1).Select
    xlApp.ActiveWindow.FreezePanes = True
    xlWs.UsedRange.Columns.AutoFit

    result("Success") = True
    result("Status") = "SUCCESS"
    result("RecordCount") = recCount
    result("Message") = "Change report exported: " & recCount & " record(s)."

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
    Set ExportChangeReportResult = result
    Exit Function

ErrorHandler:
    Debug.Print "ExportChangeReport error: " & Err.Description & " (" & Err.Number & ")"
    result("Success") = False
    result("Status") = "ERROR"
    result("ErrorNumber") = Err.Number
    result("ErrorMessage") = "Export error: " & Err.Description
    result("Message") = CStr(result("ErrorMessage"))
    GoTo Cleanup
End Function

' =============================================
' @description Formats date for Access SQL literal.
' =============================================
Private Function FormatDateLiteral(ByVal dtValue As Date) As String
    FormatDateLiteral = "#" & Format(dtValue, "mm\/dd\/yyyy") & "#"
End Function

Private Function CreateExportResult() As Object
    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d("Success") = False
    d("Status") = "PENDING"
    d("Message") = ""
    d("ErrorMessage") = ""
    d("ErrorNumber") = 0
    d("RecordCount") = 0
    d("ExportPath") = ""

    Set CreateExportResult = d
End Function
