Attribute VB_Name = "mod_Maintenance_Logic"
Option Compare Database
Option Explicit

' =============================================
' @module mod_Maintenance_Logic
' @description Settings storage and data health checks.
' @note 100% English version. Safe for modern IDEs, Git and AI tools.
' =============================================

' =============================================
' @description Retrieves a value from tbl_Settings by key. Uses DAO Recordset.
' @param key [String] Setting key
' @param defaultValue [Variant] Optional value if key not found
' @return [Variant] SettingValue or defaultValue
' =============================================
Public Function GetSetting(ByVal key As String, Optional ByVal defaultValue As Variant) As Variant
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSafeKey As String

    Debug.Print "GetSetting: entry, key=[" & key & "]"
    mod_Schema_Manager.CreateSettingsTable
    Set db = CurrentDb
    strSafeKey = Replace(Trim$(key), "'", "''")
    Debug.Print "GetSetting: opening tbl_Settings as snapshot"
    Set rs = db.OpenRecordset("tbl_Settings", dbOpenSnapshot, dbReadOnly)

    Debug.Print "GetSetting: FindFirst SettingKey = '" & strSafeKey & "'"
    rs.FindFirst "SettingKey = '" & strSafeKey & "'"
    If rs.NoMatch Then
        Debug.Print "GetSetting: key not found, returning default"
        GetSetting = defaultValue
    Else
        GetSetting = Nz(rs!SettingValue, defaultValue)
        Debug.Print "GetSetting: found, value length=" & Len(CStr(Nz(rs!SettingValue, "")))
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Debug.Print "GetSetting: done"
    Exit Function

ErrorHandler:
    Debug.Print "GetSetting error: " & Err.Description & " (" & Err.Number & ")"
    GetSetting = defaultValue
    If Not rs Is Nothing Then
        On Error Resume Next
        If rs.EditMode <> dbEditNone Then rs.Close
        rs.Close
        Set rs = Nothing
    End If
    If Not db Is Nothing Then Set db = Nothing
End Function

' =============================================
' @description Inserts or updates a setting in tbl_Settings.
' @param key [String] Setting key
' @param value [Variant] Value to store (converted to string)
' @param group [String] Optional SettingGroup, default "General"
' =============================================
Public Sub SetSetting(ByVal key As String, ByVal value As Variant, Optional ByVal group As String = "General")
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSafeKey As String
    Dim strValue As String
    Dim strGroup As String

    Debug.Print "SetSetting: entry, key=[" & key & "]"
    mod_Schema_Manager.CreateSettingsTable
    Set db = CurrentDb
    strSafeKey = Replace(Trim$(key), "'", "''")
    strValue = Trim$(CStr(Nz(value, "")))
    If Len(strValue) > 255 Then strValue = Left(strValue, 255)
    strGroup = Trim$(group)
    If Len(strGroup) > 50 Then strGroup = Left(strGroup, 50)

    Debug.Print "SetSetting: opening tbl_Settings as dynaset"
    Set rs = db.OpenRecordset("tbl_Settings", dbOpenDynaset)
    rs.FindFirst "SettingKey = '" & strSafeKey & "'"

    If Not rs.NoMatch Then
        Debug.Print "SetSetting: key found, editing existing record"
        rs.Edit
        rs!SettingValue = strValue
        rs!SettingGroup = strGroup
        rs.Update
        Debug.Print "SetSetting: record updated"
    Else
        Debug.Print "SetSetting: key not found, adding new record"
        rs.AddNew
        rs!SettingKey = Left(Trim$(key), 50)
        rs!SettingValue = strValue
        rs!SettingGroup = strGroup
        rs.Update
        Debug.Print "SetSetting: new record added"
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Debug.Print "SetSetting: done"
    Exit Sub

ErrorHandler:
    Debug.Print "SetSetting error: " & Err.Description & " (" & Err.Number & ")"
    If Not rs Is Nothing Then
        On Error Resume Next
        If rs.EditMode <> dbEditNone Then rs.CancelUpdate
        rs.Close
        Set rs = Nothing
    End If
    If Not db Is Nothing Then Set db = Nothing
End Sub

' =============================================
' @description Validates PersonUID format: 1 or 2 Cyrillic letters + hyphen + exactly 6 digits (e.g. A-111114, AA-123114).
'              Case-insensitive.
' @param strUID [String] PersonUID to validate
' @return [Boolean] True if valid
' =============================================
Public Function IsValidPersonUID(ByVal strUID As String) As Boolean
    On Error GoTo ErrorHandler

    Dim s As String
    Dim i As Long
    Dim c As Long

    IsValidPersonUID = False
    s = Trim$(Nz(strUID, ""))
    If Len(s) = 0 Then Exit Function

    ' Length: 1 letter + "-" + 6 digits = 8; 2 letters + "-" + 6 digits = 9
    If Len(s) <> 8 And Len(s) <> 9 Then Exit Function
    If Len(s) = 8 Then
        If Mid(s, 2, 1) <> "-" Then Exit Function
        For i = 3 To 8
            c = AscW(Mid(s, i, 1))
            If c < 48 Or c > 57 Then Exit Function
        Next i
        c = AscW(Mid(s, 1, 1))
        If Not IsRussianLetter(c) Then Exit Function
    Else
        If Mid(s, 3, 1) <> "-" Then Exit Function
        For i = 4 To 9
            c = AscW(Mid(s, i, 1))
            If c < 48 Or c > 57 Then Exit Function
        Next i
        For i = 1 To 2
            c = AscW(Mid(s, i, 1))
            If Not IsRussianLetter(c) Then Exit Function
        Next i
    End If

    IsValidPersonUID = True
    Exit Function

ErrorHandler:
    Debug.Print "IsValidPersonUID error: " & Err.Description
    IsValidPersonUID = False
End Function

' =============================================
' @description Helper: True if Unicode code point is a Cyrillic letter.
' =============================================
Private Function IsRussianLetter(ByVal codePoint As Long) As Boolean
    IsRussianLetter = (codePoint >= 1040 And codePoint <= 1071) Or codePoint = 1025 Or _
                      (codePoint >= 1072 And codePoint <= 1103) Or codePoint = 1105
End Function

' =============================================
' @description Writes one row to tbl_Validation_Log.
' @param recordId [Long] Record ID (0 if N/A)
' @param tableName [String] Table name
' @param errorType [String] Error type (e.g. Duplicate, Orphan, FutureDate, InvalidPersonUID)
' @param errorMessage [String] Message (max 255 chars)
' =============================================
Public Sub LogValidationError(ByVal recordId As Long, ByVal tableName As String, ByVal errorType As String, ByVal errorMessage As String)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim strSQL As String
    Dim strSafeTable As String
    Dim strSafeType As String
    Dim strSafeMsg As String

    Set db = CurrentDb
    strSafeTable = Replace(Left(Trim$(tableName), 50), "'", "''")
    strSafeType = Replace(Left(Trim$(errorType), 50), "'", "''")
    strSafeMsg = Replace(Left(Trim$(errorMessage), 255), "'", "''")

    strSQL = "INSERT INTO tbl_Validation_Log (RecordID, TableName, ErrorType, ErrorMessage, CheckDate) VALUES (" & recordId & ", '" & strSafeTable & "', '" & strSafeType & "', '" & strSafeMsg & "', Now());"
    db.Execute strSQL, dbFailOnError

    Set db = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "LogValidationError: " & Err.Description
    Set db = Nothing
End Sub

' =============================================
' @description Returns dashboard statistics.
' @return [Object] Dictionary (Late Binding)
' =============================================
Public Function GetDashboardStats() As Object
    On Error GoTo ErrorHandler

    Dim d As Object
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim lngTotal As Long
    Dim lngActive As Long
    Dim lngErrors As Long

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    lngTotal = 0
    lngActive = 0
    lngErrors = 0

    Set db = CurrentDb

    strSQL = "SELECT COUNT(*) AS Cnt FROM tbl_Personnel_Master;"
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If Not rs.EOF Then lngTotal = Nz(rs!Cnt, 0)
    rs.Close
    Set rs = Nothing

    strSQL = "SELECT COUNT(*) AS Cnt FROM tbl_Personnel_Master WHERE IsActive = True;"
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If Not rs.EOF Then lngActive = Nz(rs!Cnt, 0)
    rs.Close
    Set rs = Nothing

    mod_Schema_Manager.CreateValidationLogTable
    strSQL = "SELECT COUNT(*) AS Cnt FROM tbl_Validation_Log;"
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If Not rs.EOF Then lngErrors = Nz(rs!Cnt, 0)
    rs.Close
    Set rs = Nothing

    d("TotalCount") = lngTotal
    d("ActiveCount") = lngActive
    d("ErrorCount") = lngErrors
    Set GetDashboardStats = d

    Set db = Nothing
    Exit Function

ErrorHandler:
    Debug.Print "GetDashboardStats error: " & Err.Description & " (" & Err.Number & ")"
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    If Not d Is Nothing Then
        d("TotalCount") = 0
        d("ActiveCount") = 0
        d("ErrorCount") = 0
        Set GetDashboardStats = d
    Else
        Set d = CreateObject("Scripting.Dictionary")
        d("TotalCount") = 0
        d("ActiveCount") = 0
        d("ErrorCount") = 0
        Set GetDashboardStats = d
    End If
End Function

' =============================================
' @description Runs data health check. Writes all findings to tbl_Validation_Log.
' @param bSilentIfNoErrors [Boolean] If True, do not show message when 0 errors.
' @return [Long] Number of errors found.
' =============================================
Public Function RunDataHealthCheck(Optional ByVal bSilentIfNoErrors As Boolean = False) As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim lngTotal As Long
    Dim lngDup As Long
    Dim lngOrphan As Long
    Dim lngFuture As Long

    RunDataHealthCheck = 0
    Set db = CurrentDb
    lngTotal = 0
    lngDup = 0
    lngOrphan = 0
    lngFuture = 0

    mod_Schema_Manager.CreateValidationLogTable
    db.Execute "DELETE FROM tbl_Validation_Log;", dbFailOnError

    ' DUPLICATES: PersonUID
    strSQL = "SELECT PersonUID, COUNT(*) AS Cnt FROM tbl_Personnel_Master GROUP BY PersonUID HAVING COUNT(*)>1;"
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    Do While Not rs.EOF
        LogValidationError 0, "tbl_Personnel_Master", "Duplicate", "Duplicate PersonUID: " & Nz(rs!PersonUID, "")
        lngDup = lngDup + 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    ' DUPLICATES: FullName + BirthDate
    strSQL = "SELECT FullName, BirthDate, COUNT(*) AS Cnt FROM tbl_Personnel_Master GROUP BY FullName, BirthDate HAVING COUNT(*)>1;"
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    Do While Not rs.EOF
        LogValidationError 0, "tbl_Personnel_Master", "Duplicate", "Duplicate FullName+BirthDate: " & Nz(rs!FullName, "") & " | " & Format(Nz(rs!BirthDate, ""), "yyyy-mm-dd")
        lngDup = lngDup + 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    ' ORPHANS: PersonUID in tbl_History_Log not in tbl_Personnel_Master
    strSQL = "SELECT H.LogID, H.PersonUID FROM tbl_History_Log H LEFT JOIN tbl_Personnel_Master P ON H.PersonUID = P.PersonUID WHERE P.PersonUID IS NULL;"
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    Do While Not rs.EOF
        LogValidationError CLng(rs!logId), "tbl_History_Log", "Orphan", "PersonUID not in Personnel_Master: " & Nz(rs!PersonUID, "")
        lngOrphan = lngOrphan + 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    ' FUTURE DATES: ChangeDate > Now()
    strSQL = "SELECT LogID, PersonUID, ChangeDate FROM tbl_History_Log WHERE ChangeDate > Now();"
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    Do While Not rs.EOF
        LogValidationError CLng(rs!logId), "tbl_History_Log", "FutureDate", "ChangeDate in future: " & Format(Nz(rs!ChangeDate, ""), "yyyy-mm-dd hh:nn")
        lngFuture = lngFuture + 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    lngTotal = lngDup + lngOrphan + lngFuture
    RunDataHealthCheck = lngTotal

    If Not (bSilentIfNoErrors And lngTotal = 0) Then
        Dim strSummary As String
        strSummary = "Health check complete. " & lngTotal & " error(s) found."
        If lngTotal > 0 Then
            strSummary = strSummary & vbCrLf & "Duplicates: " & lngDup & ", Orphans: " & lngOrphan & ", Future dates: " & lngFuture
        End If
        mod_UI_Helpers.ShowMessage strSummary, vbInformation
    End If

    Set db = Nothing
    Exit Function

ErrorHandler:
    Debug.Print "RunDataHealthCheck error: " & Err.Description & " (" & Err.Number & ")"
    If Not rs Is Nothing Then
        If rs.EditMode <> dbEditNone Then rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    mod_UI_Helpers.ShowMessage "Health check failed: " & Err.Description, vbExclamation
    RunDataHealthCheck = -1
End Function

' =============================================
' @description Exports tbl_Validation_Log to a new Excel workbook (Late Binding).
' @return [Boolean] True if success
' =============================================
Public Function ExportValidationLogToExcel() As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
    Dim colCount As Long
    Dim recCount As Long

    ExportValidationLogToExcel = False
    mod_Schema_Manager.CreateValidationLogTable
    Set db = CurrentDb

    strSQL = "SELECT LogID AS [ID], RecordID, TableName AS [Table], ErrorType AS [Error Type], ErrorMessage AS [Message], CheckDate AS [Date] FROM tbl_Validation_Log ORDER BY LogID;"
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)

    If rs.EOF Then
        mod_UI_Helpers.ShowMessage "No validation log records to export.", vbInformation
        rs.Close
        Set rs = Nothing
        Set db = Nothing
        Exit Function
    End If

    colCount = 6
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    xlApp.DisplayAlerts = False
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Worksheets(1)

    xlWs.Cells(1, 1).value = "ID"
    xlWs.Cells(1, 2).value = "RecordID"
    xlWs.Cells(1, 3).value = "Table"
    xlWs.Cells(1, 4).value = "Error Type"
    xlWs.Cells(1, 5).value = "Message"
    xlWs.Cells(1, 6).value = "Date"

    xlWs.Range("A2").CopyFromRecordset rs
    recCount = xlWs.UsedRange.Rows.count - 1
    If recCount < 0 Then recCount = 0

    With xlWs.Range(xlWs.Cells(1, 1), xlWs.Cells(1, colCount))
        .Font.Bold = True
    End With
    xlWs.Rows(2).Select
    xlApp.ActiveWindow.FreezePanes = True
    xlWs.Range(xlWs.Cells(1, 1), xlWs.Cells(1, colCount)).Select
    xlWs.UsedRange.Columns.AutoFit

    mod_UI_Helpers.ShowMessage "Validation log exported: " & recCount & " record(s).", vbInformation
    ExportValidationLogToExcel = True

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
    Debug.Print "ExportValidationLogToExcel error: " & Err.Description & " (" & Err.Number & ")"
    mod_UI_Helpers.ShowMessage "Export failed: " & Err.Description, vbExclamation
    ExportValidationLogToExcel = False
    GoTo Cleanup
End Function

' =============================================
' @description Creates a backup copy of the current database in \Backups subfolder.
' @return [Boolean] True on success
' =============================================
Public Function CreateDatabaseBackup() As Boolean
    On Error GoTo ErrorHandler

    Dim strBackupDir As String
    Dim strDestPath As String
    Dim fso As Object

    CreateDatabaseBackup = False
    strBackupDir = CurrentProject.Path & "\Backups"
    If Len(CurrentProject.Path) = 0 Or Len(CurrentProject.FullName) = 0 Then
        mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetMsgBackupPathUndefined(), vbExclamation
        Exit Function
    End If

    If Dir(strBackupDir, vbDirectory) = "" Then
        MkDir strBackupDir
        Debug.Print "CreateDatabaseBackup: created folder " & strBackupDir
    End If

    strDestPath = strBackupDir & "\StaffState_Backup_" & Format(Now(), "yyyymmdd_hhnnss") & ".accdb"
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile CurrentProject.FullName, strDestPath, True
    Set fso = Nothing

    mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetMsgBackupSaved(strDestPath), vbInformation
    CreateDatabaseBackup = True
    Exit Function

ErrorHandler:
    Debug.Print "CreateDatabaseBackup error: " & Err.Description & " (" & Err.Number & ")"
    If Err.Number = 70 Then
        mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetMsgBackupError70(), vbExclamation
    Else
        mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetMsgBackupFailedGeneric(Err.Description), vbExclamation
    End If
    If Not fso Is Nothing Then Set fso = Nothing
End Function

' =============================================
' @description Clears tbl_Validation_Log after user confirmation.
' =============================================
Public Sub ClearValidationLog()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database

    If Not mod_UI_Helpers.AskUserYesNo("Are you sure you want to clear the validation log?", "StaffState") Then Exit Sub

    Set db = CurrentDb
    db.Execute "DELETE * FROM tbl_Validation_Log", dbFailOnError
    Set db = Nothing
    mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetMsgValidationLogCleared(), vbInformation
    Exit Sub

ErrorHandler:
    Debug.Print "ClearValidationLog error: " & Err.Description & " (" & Err.Number & ")"
    If Not db Is Nothing Then Set db = Nothing
    mod_UI_Helpers.ShowMessage "Clear failed: " & Err.Description, vbExclamation
End Sub
