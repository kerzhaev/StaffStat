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
' =============================================
' @description Writes one row to tbl_Validation_Log using Parameterized QueryDef.
' =============================================
Public Sub LogValidationError(ByVal recordId As Long, ByVal tableName As String, ByVal errorType As String, ByVal errorMessage As String)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim strSQL As String

    Set db = CurrentDb
    strSQL = "PARAMETERS prmRecID Long, prmTable Text (50), prmType Text (50), prmMsg Text (255); " & _
             "INSERT INTO tbl_Validation_Log (RecordID, TableName, ErrorType, ErrorMessage, CheckDate) " & _
             "VALUES ([prmRecID], [prmTable], [prmType], [prmMsg], Now());"

    Set qdf = db.CreateQueryDef("", strSQL)
    qdf.Parameters("prmRecID").value = recordId
    qdf.Parameters("prmTable").value = Left$(Trim$(tableName), 50)
    qdf.Parameters("prmType").value = Left$(Trim$(errorType), 50)
    qdf.Parameters("prmMsg").value = Left$(Trim$(errorMessage), 255)

    qdf.Execute dbFailOnError

    Set qdf = Nothing
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
    Dim result As Object

    Set result = RunDataHealthCheckResult(bSilentIfNoErrors)
    If CBool(result("Success")) Then
        RunDataHealthCheck = CLng(result("TotalErrors"))
    Else
        RunDataHealthCheck = -1
    End If
    Set result = Nothing
End Function

Public Function RunDataHealthCheckResult(Optional ByVal bSilentIfNoErrors As Boolean = False) As Object
    On Error GoTo ErrorHandler

    Dim result As Object
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim lngTotal As Long
    Dim lngDup As Long
    Dim lngOrphan As Long
    Dim lngFuture As Long
    Dim lngEmpty As Long

    Set result = CreateOperationResult()
    result("TotalErrors") = 0
    result("DuplicateCount") = 0
    result("OrphanCount") = 0
    result("FutureDateCount") = 0
    result("EmptyFieldCount") = 0
    result("ShouldNotifyUser") = False

    Set db = CurrentDb
    lngTotal = 0
    lngDup = 0
    lngOrphan = 0
    lngFuture = 0
    lngEmpty = 0

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

    ' EMPTY FIELDS: FullName, BirthDate_Text, PersonUID in tbl_Personnel_Master
    strSQL = "SELECT PersonUID, FullName, BirthDate_Text FROM tbl_Personnel_Master " & _
             "WHERE Nz(FullName,'')='' OR Nz(BirthDate_Text,'')='' OR Nz(PersonUID,'')='';"
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    Do While Not rs.EOF
        Dim strEmptyDetails As String
        strEmptyDetails = ""
        If Nz(rs!FullName, "") = "" Then strEmptyDetails = strEmptyDetails & "FullName "
        If Nz(rs!BirthDate_Text, "") = "" Then strEmptyDetails = strEmptyDetails & "BirthDate_Text "
        If Nz(rs!PersonUID, "") = "" Then strEmptyDetails = strEmptyDetails & "PersonUID "

        LogValidationError 0, "tbl_Personnel_Master", "EmptyRequiredField", "Empty required fields: " & Trim$(strEmptyDetails)
        lngEmpty = lngEmpty + 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    lngTotal = lngDup + lngOrphan + lngFuture + lngEmpty
    result("Success") = True
    result("TotalErrors") = lngTotal
    result("DuplicateCount") = lngDup
    result("OrphanCount") = lngOrphan
    result("FutureDateCount") = lngFuture
    result("EmptyFieldCount") = lngEmpty
    result("Message") = BuildHealthCheckSummary(lngTotal, lngDup, lngOrphan, lngFuture, lngEmpty)
    result("ShouldNotifyUser") = Not (bSilentIfNoErrors And lngTotal = 0)

    Set db = Nothing
    Set RunDataHealthCheckResult = result
    Exit Function

ErrorHandler:
    Debug.Print "RunDataHealthCheck error: " & Err.Description & " (" & Err.Number & ")"
    If Not rs Is Nothing Then
        If rs.EditMode <> dbEditNone Then rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    result("Success") = False
    result("ErrorNumber") = Err.Number
    result("ErrorMessage") = "Health check failed: " & Err.Description
    result("Message") = CStr(result("ErrorMessage"))
    result("TotalErrors") = -1
    result("ShouldNotifyUser") = True
    Set RunDataHealthCheckResult = result
End Function

' =============================================
' @description Exports tbl_Validation_Log to a new Excel workbook (Late Binding).
' @return [Boolean] True if success
' =============================================
Public Function ExportValidationLogToExcel() As Boolean
    Dim result As Object

    Set result = ExportValidationLogResult()
    ExportValidationLogToExcel = CBool(result("Success")) And CBool(result("HasData"))
    Set result = Nothing
End Function

Public Function ExportValidationLogResult() As Object
    On Error GoTo ErrorHandler

    Dim result As Object
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
    Dim colCount As Long
    Dim recCount As Long

    Set result = CreateOperationResult()
    result("HasData") = False
    result("RecordCount") = 0

    mod_Schema_Manager.CreateValidationLogTable
    Set db = CurrentDb

    strSQL = "SELECT LogID AS [ID], RecordID, TableName AS [Table], ErrorType AS [Error Type], ErrorMessage AS [Message], CheckDate AS [Date] FROM tbl_Validation_Log ORDER BY LogID;"
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)

    If rs.EOF Then
        result("Success") = True
        result("Message") = "No validation log records to export."
        rs.Close
        Set rs = Nothing
        Set db = Nothing
        Set ExportValidationLogResult = result
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

    result("Success") = True
    result("HasData") = True
    result("RecordCount") = recCount
    result("Message") = "Validation log exported: " & recCount & " record(s)."

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
    Set ExportValidationLogResult = result
    Exit Function

ErrorHandler:
    Debug.Print "ExportValidationLogToExcel error: " & Err.Description & " (" & Err.Number & ")"
    result("Success") = False
    result("ErrorNumber") = Err.Number
    result("ErrorMessage") = "Export failed: " & Err.Description
    result("Message") = CStr(result("ErrorMessage"))
    GoTo Cleanup
End Function

' =============================================
' @description Creates a backup copy of the current database in \Backups subfolder.
' @return [Boolean] True on success
' =============================================
Public Function CreateDatabaseBackup() As Boolean
    Dim result As Object

    Set result = CreateDatabaseBackupResult()
    CreateDatabaseBackup = CBool(result("Success"))
    Set result = Nothing
End Function

Public Function CreateDatabaseBackupResult() As Object
    On Error GoTo ErrorHandler

    Dim result As Object
    Dim strBackupDir As String
    Dim strDestPath As String
    Dim fso As Object

    Set result = CreateOperationResult()
    result("BackupPath") = ""

    strBackupDir = CurrentProject.Path & "\Backups"
    If Len(CurrentProject.Path) = 0 Or Len(CurrentProject.FullName) = 0 Then
        result("ErrorMessage") = mod_UI_Helpers.GetMsgBackupPathUndefined()
        result("Message") = CStr(result("ErrorMessage"))
        Set CreateDatabaseBackupResult = result
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

    result("Success") = True
    result("BackupPath") = strDestPath
    result("Message") = mod_UI_Helpers.GetMsgBackupSaved(strDestPath)
    Set CreateDatabaseBackupResult = result
    Exit Function

ErrorHandler:
    Debug.Print "CreateDatabaseBackup error: " & Err.Description & " (" & Err.Number & ")"
    result("Success") = False
    result("ErrorNumber") = Err.Number
    If Err.Number = 70 Then
        result("ErrorMessage") = mod_UI_Helpers.GetMsgBackupError70()
    Else
        result("ErrorMessage") = mod_UI_Helpers.GetMsgBackupFailedGeneric(Err.Description)
    End If
    result("Message") = CStr(result("ErrorMessage"))
    If Not fso Is Nothing Then Set fso = Nothing
    Set CreateDatabaseBackupResult = result
End Function

' =============================================
' @description Clears tbl_Validation_Log.
' =============================================
Public Function ClearValidationLog() As Boolean
    Dim result As Object

    Set result = ClearValidationLogResult()
    ClearValidationLog = CBool(result("Success"))
    Set result = Nothing
End Function

Public Function ClearValidationLogResult() As Object
    On Error GoTo ErrorHandler

    Dim result As Object
    Dim db As DAO.Database

    Set result = CreateOperationResult()

    Set db = CurrentDb
    db.Execute "DELETE * FROM tbl_Validation_Log", dbFailOnError
    Set db = Nothing
    result("Success") = True
    result("Message") = mod_UI_Helpers.GetMsgValidationLogCleared()
    Set ClearValidationLogResult = result
    Exit Function

ErrorHandler:
    Debug.Print "ClearValidationLog error: " & Err.Description & " (" & Err.Number & ")"
    If Not db Is Nothing Then Set db = Nothing
    result("Success") = False
    result("ErrorNumber") = Err.Number
    result("ErrorMessage") = "Clear failed: " & Err.Description
    result("Message") = CStr(result("ErrorMessage"))
    Set ClearValidationLogResult = result
End Function

' =============================================
' @description Wipes all user data (Master, History, Logs, Buffer).
'              Used for development and testing.
' =============================================
Public Function FactoryResetDataResult() As Object
    On Error GoTo ErrorHandler

    Dim result As Object
    Dim db As DAO.Database

    Set result = CreateOperationResult()
    Set db = CurrentDb

    ' Delete all records from linked tables
    db.Execute "DELETE FROM tbl_Personnel_Master;", dbFailOnError
    db.Execute "DELETE FROM tbl_History_Log;", dbFailOnError
    db.Execute "DELETE FROM tbl_Import_Buffer;", dbFailOnError
    db.Execute "DELETE FROM tbl_Validation_Log;", dbFailOnError

    result("Success") = True
    result("Message") = "Database cleared successfully!"
    Set db = Nothing
    Set FactoryResetDataResult = result
    Exit Function

ErrorHandler:
    If Not db Is Nothing Then Set db = Nothing
    result("Success") = False
    result("ErrorNumber") = Err.Number
    result("ErrorMessage") = "Error clearing data: " & Err.Description
    result("Message") = CStr(result("ErrorMessage"))
    Set FactoryResetDataResult = result
End Function

Public Sub FactoryResetData()
    Dim result As Object

    Set result = FactoryResetDataResult()
    Set result = Nothing
End Sub

Private Function CreateOperationResult() As Object
    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d("Success") = False
    d("Message") = ""
    d("ErrorMessage") = ""
    d("ErrorNumber") = 0

    Set CreateOperationResult = d
End Function

Private Function BuildHealthCheckSummary(ByVal lngTotal As Long, ByVal lngDup As Long, ByVal lngOrphan As Long, ByVal lngFuture As Long, ByVal lngEmpty As Long) As String
    Dim strSummary As String

    strSummary = "Health check complete. " & lngTotal & " error(s) found."
    If lngTotal > 0 Then
        strSummary = strSummary & vbCrLf & _
                     "Duplicates: " & lngDup & ", Orphans: " & lngOrphan & _
                     ", Future dates: " & lngFuture & ", Empty fields: " & lngEmpty
    End If

    BuildHealthCheckSummary = strSummary
End Function
