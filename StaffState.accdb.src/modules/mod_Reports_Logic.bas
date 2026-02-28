Attribute VB_Name = "mod_Reports_Logic"
Option Compare Database
Option Explicit

' =============================================
' @module mod_Reports_Logic
' @description Audit report: export personnel change history to Excel (Phase 15).
'              Excel: Late Binding only (CreateObject("Excel.Application")); no reference to Excel Object Library.
'              Excel constants used as literals: xlUp=-4162, xlLeft=-4131, xlTop=-4160, xlContinuous=1.
' Comments: English only. Keep source ASCII-only; Russian UI via ChrW().
' =============================================

' --- Russian UI strings (ChrW = Unicode, display correctly in Access regardless of file encoding) ---
Private Function RuPromptStartDate() As String
    RuPromptStartDate = ChrW(1044) & ChrW(1072) & ChrW(1090) & ChrW(1072) & " " & ChrW(1085) & ChrW(1072) & ChrW(1095) & ChrW(1072) & ChrW(1083) & ChrW(1072) & " " & ChrW(1087) & ChrW(1077) & ChrW(1088) & ChrW(1080) & ChrW(1086) & ChrW(1076) & ChrW(1072) & " (" & ChrW(1044) & ChrW(1044) & "." & ChrW(1052) & ChrW(1052) & "." & ChrW(1043) & ChrW(1043) & ChrW(1043) & " " & ChrW(1080) & ChrW(1083) & ChrW(1080) & " " & ChrW(1044) & ChrW(1044) & "/" & ChrW(1052) & ChrW(1052) & "/" & ChrW(1043) & ChrW(1043) & ChrW(1043) & "):"
End Function
Private Function RuPromptEndDate() As String
    RuPromptEndDate = ChrW(1044) & ChrW(1072) & ChrW(1090) & ChrW(1072) & " " & ChrW(1086) & ChrW(1082) & ChrW(1086) & ChrW(1085) & ChrW(1095) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1103) & " " & ChrW(1087) & ChrW(1077) & ChrW(1088) & ChrW(1080) & ChrW(1086) & ChrW(1076) & ChrW(1072) & " (" & ChrW(1044) & ChrW(1044) & "." & ChrW(1052) & ChrW(1052) & "." & ChrW(1043) & ChrW(1043) & ChrW(1043) & " " & ChrW(1080) & ChrW(1083) & ChrW(1080) & " " & ChrW(1044) & ChrW(1044) & "/" & ChrW(1052) & ChrW(1052) & "/" & ChrW(1043) & ChrW(1043) & ChrW(1043) & "):"
End Function
Private Function RuTitleAuditPeriod() As String
    RuTitleAuditPeriod = ChrW(1040) & ChrW(1091) & ChrW(1076) & ChrW(1080) & ChrW(1090) & " " & ChrW(8212) & " " & ChrW(1087) & ChrW(1077) & ChrW(1088) & ChrW(1080) & ChrW(1086) & ChrW(1076)
End Function
Private Function RuInvalidDateFormat() As String
    RuInvalidDateFormat = ChrW(1053) & ChrW(1077) & ChrW(1074) & ChrW(1077) & ChrW(1088) & ChrW(1085) & ChrW(1099) & ChrW(1081) & " " & ChrW(1092) & ChrW(1086) & ChrW(1088) & ChrW(1084) & ChrW(1072) & ChrW(1090) & " " & ChrW(1076) & ChrW(1072) & ChrW(1090) & ChrW(1099) & ". " & ChrW(1059) & ChrW(1082) & ChrW(1072) & ChrW(1078) & ChrW(1080) & ChrW(1090) & ChrW(1077) & " " & ChrW(1076) & ChrW(1072) & ChrW(1090) & ChrW(1099) & " " & ChrW(1074) & " " & ChrW(1092) & ChrW(1086) & ChrW(1088) & ChrW(1084) & ChrW(1072) & ChrW(1090) & ChrW(1077) & " " & ChrW(1044) & ChrW(1044) & "." & ChrW(1052) & ChrW(1052) & "." & ChrW(1043) & ChrW(1043) & ChrW(1043) & "."
End Function
Private Function RuStartAfterEnd() As String
    RuStartAfterEnd = ChrW(1044) & ChrW(1072) & ChrW(1090) & ChrW(1072) & " " & ChrW(1085) & ChrW(1072) & ChrW(1095) & ChrW(1072) & ChrW(1083) & ChrW(1072) & " " & ChrW(1085) & ChrW(1077) & " " & ChrW(1084) & ChrW(1086) & ChrW(1078) & ChrW(1077) & ChrW(1090) & " " & ChrW(1073) & ChrW(1099) & ChrW(1090) & ChrW(1100) & " " & ChrW(1087) & ChrW(1086) & ChrW(1079) & ChrW(1078) & ChrW(1077) & ChrW(1095) & ChrW(1077) & " " & ChrW(1076) & ChrW(1072) & ChrW(1090) & ChrW(1099) & " " & ChrW(1086) & ChrW(1082) & ChrW(1086) & ChrW(1085) & ChrW(1095) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1103) & "."
End Function
Private Function RuNoChangesFound() As String
    RuNoChangesFound = ChrW(1047) & ChrW(1072) & " " & ChrW(1091) & ChrW(1082) & ChrW(1072) & ChrW(1079) & ChrW(1072) & ChrW(1085) & ChrW(1085) & ChrW(1099) & ChrW(1081) & " " & ChrW(1087) & ChrW(1077) & ChrW(1088) & ChrW(1080) & ChrW(1086) & ChrW(1076) & " " & ChrW(1080) & ChrW(1079) & ChrW(1084) & ChrW(1077) & ChrW(1085) & ChrW(1077) & ChrW(1085) & ChrW(1080) & ChrW(1081) & " " & ChrW(1085) & ChrW(1077) & " " & ChrW(1085) & ChrW(1072) & ChrW(1081) & ChrW(1076) & ChrW(1077) & ChrW(1085) & ChrW(1086)
End Function
Private Function RuReportTitle() As String
    RuReportTitle = ChrW(1054) & ChrW(1090) & ChrW(1095) & ChrW(1077) & ChrW(1090) & " " & ChrW(1087) & ChrW(1086) & " " & ChrW(1080) & ChrW(1079) & ChrW(1084) & ChrW(1077) & ChrW(1085) & ChrW(1077) & ChrW(1085) & ChrW(1080) & ChrW(1103) & ChrW(1084) & " " & ChrW(1076) & ChrW(1072) & ChrW(1085) & ChrW(1085) & ChrW(1099) & ChrW(1093) & " (" & ChrW(1040) & ChrW(1091) & ChrW(1076) & ChrW(1080) & ChrW(1090) & ")"
End Function
Private Function RuPeriodFrom() As String
    RuPeriodFrom = ChrW(1055) & ChrW(1077) & ChrW(1088) & ChrW(1080) & ChrW(1086) & ChrW(1076) & ": " & ChrW(1089) & " "
End Function
Private Function RuPeriodTo() As String
    RuPeriodTo = " " & ChrW(1087) & ChrW(1086) & " "
End Function
Private Function RuColEmployee() As String
    RuColEmployee = ChrW(1057) & ChrW(1086) & ChrW(1090) & ChrW(1088) & ChrW(1091) & ChrW(1076) & ChrW(1085) & ChrW(1080) & ChrW(1082)
End Function
Private Function RuColPersonUID() As String
    RuColPersonUID = ChrW(1051) & ChrW(1080) & ChrW(1095) & ChrW(1085) & ChrW(1099) & ChrW(1081) & " " & ChrW(1085) & ChrW(1086) & ChrW(1084) & ChrW(1077) & ChrW(1088)
End Function
Private Function RuColField() As String
    RuColField = ChrW(1055) & ChrW(1086) & ChrW(1083) & ChrW(1077)
End Function
Private Function RuColChangeDate() As String
    RuColChangeDate = ChrW(1044) & ChrW(1072) & ChrW(1090) & ChrW(1072) & " " & ChrW(1080) & ChrW(1079) & ChrW(1084) & ChrW(1077) & ChrW(1085) & ChrW(1077) & ChrW(1085) & ChrW(1080) & ChrW(1103)
End Function
Private Function RuColOldValue() As String
    RuColOldValue = ChrW(1057) & ChrW(1090) & ChrW(1072) & ChrW(1088) & ChrW(1086) & ChrW(1077) & " " & ChrW(1079) & ChrW(1085) & ChrW(1072) & ChrW(1095) & ChrW(1077) & ChrW(1085) & ChrW(1080) & ChrW(1077)
End Function
Private Function RuColNewValue() As String
    RuColNewValue = ChrW(1053) & ChrW(1086) & ChrW(1074) & ChrW(1086) & ChrW(1077) & " " & ChrW(1079) & ChrW(1085) & ChrW(1072) & ChrW(1095) & ChrW(1077) & ChrW(1085) & ChrW(1080) & ChrW(1077)
End Function
Private Function RuErrorReport() As String
    RuErrorReport = ChrW(1054) & ChrW(1096) & ChrW(1080) & ChrW(1073) & ChrW(1082) & ChrW(1072) & " " & ChrW(1092) & ChrW(1086) & ChrW(1088) & ChrW(1084) & ChrW(1080) & ChrW(1088) & ChrW(1086) & ChrW(1074) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1103) & " " & ChrW(1086) & ChrW(1090) & ChrW(1095) & ChrW(1105) & ChrW(1090) & ChrW(1072) & ": "
End Function
Private Function RuTitleSnapshot() As String
    RuTitleSnapshot = ChrW(1064) & ChrW(1090) & ChrW(1072) & ChrW(1090) & ChrW(1085) & ChrW(1099) & ChrW(1081) & " " & ChrW(1089) & ChrW(1088) & ChrW(1077) & ChrW(1079)
End Function
Private Function RuColWorkStatus() As String
    RuColWorkStatus = ChrW(1057) & ChrW(1090) & ChrW(1072) & ChrW(1090) & ChrW(1091) & ChrW(1089)
End Function
Private Function RuColIsActive() As String
    RuColIsActive = ChrW(1040) & ChrW(1082) & ChrW(1090) & ChrW(1080) & ChrW(1074) & ChrW(1077) & ChrW(1085)
End Function
Private Function RuColFullName() As String
    RuColFullName = ChrW(1060) & ChrW(1048) & ChrW(1054)
End Function
Private Function RuColRankName() As String
    RuColRankName = ChrW(1047) & ChrW(1074) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1077)
End Function
Private Function RuColPosName() As String
    RuColPosName = ChrW(1044) & ChrW(1086) & ChrW(1083) & ChrW(1078) & ChrW(1085) & ChrW(1086) & ChrW(1089) & ChrW(1090) & ChrW(1100)
End Function
Private Function RuNoDataSnapshot() As String
    RuNoDataSnapshot = ChrW(1053) & ChrW(1077) & ChrW(1090) & " " & ChrW(1076) & ChrW(1072) & ChrW(1085) & ChrW(1085) & ChrW(1099) & ChrW(1093) & " " & ChrW(1076) & ChrW(1083) & ChrW(1103) & " " & ChrW(1089) & ChrW(1088) & ChrW(1077) & ChrW(1079) & ChrW(1072) & "."
End Function

' =============================================
' @description Exports personnel change history to Excel.
'              If dtStart/dtEnd provided and valid, use them; else ask via InputBox (e.g. when called without args).
'              Uses DAO JOIN (tbl_History_Log + tbl_Personnel_Master). Late-binding Excel.
' @param dtStart [Variant] Optional. Start date (e.g. from form); if both omitted, uses InputBox.
' @param dtEnd [Variant] Optional. End date (e.g. from form).
' =============================================
Public Sub GenerateAuditReport(Optional ByVal dtStart As Variant, Optional ByVal dtEnd As Variant)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim strStart As String
    Dim strEnd As String
    Dim strStartUS As String
    Dim strEndUS As String
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
    Dim lastRow As Long
    Dim rngData As Object
    Dim rngHeaders As Object
    Dim bUseFormDates As Boolean

    bUseFormDates = (Not IsMissing(dtStart)) And (Not IsMissing(dtEnd)) And IsDate(dtStart) And IsDate(dtEnd) And (CDate(dtStart) <= CDate(dtEnd))

    If bUseFormDates Then
        dtStart = CDate(dtStart)
        dtEnd = CDate(dtEnd)
    Else
        ' Ask user for dates via InputBox (when called without valid params)
        strStart = Trim$(InputBox(RuPromptStartDate(), RuTitleAuditPeriod(), Format(Date - 30, "dd.mm.yyyy")))
        If strStart = "" Then Exit Sub
        strEnd = Trim$(InputBox(RuPromptEndDate(), RuTitleAuditPeriod(), Format(Date, "dd.mm.yyyy")))
        If strEnd = "" Then Exit Sub
        strStart = Replace(strStart, ".", "/")
        strEnd = Replace(strEnd, ".", "/")
        If Not IsDate(strStart) Or Not IsDate(strEnd) Then
            mod_UI_Helpers.ShowMessage RuInvalidDateFormat(), vbExclamation
            Exit Sub
        End If
        dtStart = CDate(strStart)
        dtEnd = CDate(strEnd)
        If dtStart > dtEnd Then
            mod_UI_Helpers.ShowMessage RuStartAfterEnd(), vbExclamation
            Exit Sub
        End If
    End If

    ' Build SQL with US date literals (Access/Jet)
    strStartUS = "#" & Format(dtStart, "mm\/dd\/yyyy") & "#"
    strEndUS = "#" & Format(dtEnd, "mm\/dd\/yyyy") & " 23:59:59#"

    Set db = CurrentDb
    strSQL = "SELECT M.FullName, H.PersonUID, H.FieldName, H.ChangeDate, H.OldValue, H.NewValue " & _
             "FROM tbl_History_Log AS H " & _
             "INNER JOIN tbl_Personnel_Master AS M ON H.PersonUID = M.PersonUID " & _
             "WHERE H.ChangeDate BETWEEN " & strStartUS & " AND " & strEndUS & " " & _
             "ORDER BY H.ChangeDate DESC, M.FullName ASC;"
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)

    If rs.EOF Then
        mod_UI_Helpers.ShowMessage RuNoChangesFound(), vbInformation
        GoTo Cleanup
    End If

    ' Excel: Late Binding
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    xlApp.DisplayAlerts = False
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Worksheets(1)

    ' Header row 1: title
    xlWs.Cells(1, 1).value = RuReportTitle()
    xlWs.Range("A1:F1").Merge
    With xlWs.Range("A1").Font
        .Bold = True
        .Size = 14
    End With

    ' Sub-header row 2: period
    xlWs.Cells(2, 1).value = RuPeriodFrom() & Format(dtStart, "dd.mm.yyyy") & RuPeriodTo() & Format(dtEnd, "dd.mm.yyyy")
    xlWs.Range("A2:F2").Merge
    xlWs.Range("A2").Font.Italic = True

    ' Table headers row 4
    xlWs.Cells(4, 1).value = RuColEmployee()
    xlWs.Cells(4, 2).value = RuColPersonUID()
    xlWs.Cells(4, 3).value = RuColField()
    xlWs.Cells(4, 4).value = RuColChangeDate()
    xlWs.Cells(4, 5).value = RuColOldValue()
    xlWs.Cells(4, 6).value = RuColNewValue()

    ' Data from row 5
    xlWs.Range("A5").CopyFromRecordset rs

    lastRow = xlWs.Cells(xlWs.Rows.count, 1).End(-4162).Row
    If lastRow < 5 Then lastRow = 5

    ' Format: freeze top 3 rows (select A4 so rows 1-3 freeze)
    xlWs.Range("A4").Select
    xlApp.ActiveWindow.FreezePanes = True

    ' Bold headers row 4
    Set rngHeaders = xlWs.Range(xlWs.Cells(4, 1), xlWs.Cells(4, 6))
    rngHeaders.Font.Bold = True

    ' AutoFilter on row 4
    Set rngData = xlWs.Range(xlWs.Cells(4, 1), xlWs.Cells(lastRow, 6))
    rngData.AutoFilter

    ' Date column (D) format: dd.mm.yyyy hh:mm
    xlWs.Range(xlWs.Cells(5, 4), xlWs.Cells(lastRow, 4)).NumberFormat = "dd.mm.yyyy hh:mm"

    ' Borders and alignment for table (rows 4 to lastRow)
    With xlWs.Range(xlWs.Cells(4, 1), xlWs.Cells(lastRow, 6))
        .Borders.LineStyle = 1
        .HorizontalAlignment = -4131
        .VerticalAlignment = -4160
    End With

    xlWs.UsedRange.Columns.AutoFit

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    If Not rngData Is Nothing Then Set rngData = Nothing
    If Not rngHeaders Is Nothing Then Set rngHeaders = Nothing
    If Not xlWs Is Nothing Then Set xlWs = Nothing
    If Not xlWb Is Nothing Then Set xlWb = Nothing
    If Not xlApp Is Nothing Then
        xlApp.DisplayAlerts = True
        Set xlApp = Nothing
    End If
    Exit Sub

ErrorHandler:
    Debug.Print "GenerateAuditReport error: " & Err.Description & " (" & Err.Number & ")"
    mod_UI_Helpers.ShowMessage RuErrorReport() & Err.Description, vbCritical
    GoTo Cleanup
End Sub

' =============================================
' @description Exports current personnel (tbl_Personnel_Master) to Excel. Order by FullName.
'              Columns: PersonUID, FullName, RankName, PosName, WorkStatus, IsActive.
'              Style: Bold headers, AutoFilter, FreezePanes, Borders (same as GenerateAuditReport).
' =============================================
Public Sub GenerateCurrentStaffReport()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
    Dim lastRow As Long
    Dim rngData As Object
    Dim rngHeaders As Object

    Set db = CurrentDb
    strSQL = "SELECT PersonUID, FullName, RankName, PosName, WorkStatus, IsActive FROM tbl_Personnel_Master ORDER BY FullName ASC;"
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)

    If rs.EOF Then
        mod_UI_Helpers.ShowMessage RuNoDataSnapshot(), vbInformation
        GoTo Cleanup
    End If

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    xlApp.DisplayAlerts = False
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Worksheets(1)

    ' Title row 1
    xlWs.Cells(1, 1).value = RuTitleSnapshot()
    xlWs.Range("A1:F1").Merge
    With xlWs.Range("A1").Font
        .Bold = True
        .Size = 14
    End With

    ' Headers row 3
    xlWs.Cells(3, 1).value = RuColPersonUID()
    xlWs.Cells(3, 2).value = RuColFullName()
    xlWs.Cells(3, 3).value = RuColRankName()
    xlWs.Cells(3, 4).value = RuColPosName()
    xlWs.Cells(3, 5).value = RuColWorkStatus()
    xlWs.Cells(3, 6).value = RuColIsActive()

    ' Data from row 4
    xlWs.Range("A4").CopyFromRecordset rs

    lastRow = xlWs.Cells(xlWs.Rows.count, 1).End(-4162).Row
    If lastRow < 4 Then lastRow = 4

    ' Freeze first 3 rows
    xlWs.Range("A4").Select
    xlApp.ActiveWindow.FreezePanes = True

    ' Bold headers
    Set rngHeaders = xlWs.Range(xlWs.Cells(3, 1), xlWs.Cells(3, 6))
    rngHeaders.Font.Bold = True

    ' AutoFilter
    Set rngData = xlWs.Range(xlWs.Cells(3, 1), xlWs.Cells(lastRow, 6))
    rngData.AutoFilter

    ' Borders and alignment
    With xlWs.Range(xlWs.Cells(3, 1), xlWs.Cells(lastRow, 6))
        .Borders.LineStyle = 1
        .HorizontalAlignment = -4131
        .VerticalAlignment = -4160
    End With

    xlWs.UsedRange.Columns.AutoFit

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    If Not rngData Is Nothing Then Set rngData = Nothing
    If Not rngHeaders Is Nothing Then Set rngHeaders = Nothing
    If Not xlWs Is Nothing Then Set xlWs = Nothing
    If Not xlWb Is Nothing Then Set xlWb = Nothing
    If Not xlApp Is Nothing Then
        xlApp.DisplayAlerts = True
        Set xlApp = Nothing
    End If
    Exit Sub

ErrorHandler:
    Debug.Print "GenerateCurrentStaffReport error: " & Err.Description & " (" & Err.Number & ")"
    mod_UI_Helpers.ShowMessage RuErrorReport() & Err.Description, vbCritical
    GoTo Cleanup
End Sub
