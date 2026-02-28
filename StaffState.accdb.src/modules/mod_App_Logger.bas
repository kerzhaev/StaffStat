Attribute VB_Name = "mod_App_Logger"
Option Compare Database
Option Explicit

' =============================================
' @module mod_App_Logger
' @author Kerzhaev Evgeniy
' @description Centralized error and event logging system
' @note 100% English version. Safe for modern IDEs.
' =============================================

Private Const cstrTableName As String = "tbl_System_Logs"

' =============================================
' @description Checks for tbl_System_Logs table and creates it if necessary
' =============================================
Public Sub InitLogger()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim blnTableExists As Boolean

    Set db = CurrentDb

    ' Check if table exists
    blnTableExists = False
    On Error Resume Next
    Set tdf = db.TableDefs(cstrTableName)
    If Err.Number = 0 Then blnTableExists = True
    On Error GoTo ErrorHandler

    If Not blnTableExists Then
        ' Create table via DAO
        Set tdf = db.CreateTableDef(cstrTableName)

        ' ID field (Counter, Primary Key)
        Set fld = tdf.CreateField("ID", dbLong)
        fld.Attributes = dbAutoIncrField
        tdf.Fields.Append fld

        ' LogDate field (DateTime, Default Now)
        Set fld = tdf.CreateField("LogDate", dbDate)
        fld.defaultValue = "Now()"
        tdf.Fields.Append fld

        ' LogType field (Text 50)
        Set fld = tdf.CreateField("LogType", dbText, 50)
        tdf.Fields.Append fld

        ' Source field (Text 255)
        Set fld = tdf.CreateField("Source", dbText, 255)
        tdf.Fields.Append fld

        ' Description field (Memo)
        Set fld = tdf.CreateField("Description", dbMemo)
        tdf.Fields.Append fld

        ' WinUser field (Text 50)
        Set fld = tdf.CreateField("WinUser", dbText, 50)
        tdf.Fields.Append fld

        ' Create primary key via index
        Set idx = tdf.CreateIndex("PRIMARYKEY")
        Set fld = idx.CreateField("ID")
        idx.Fields.Append fld
        idx.Primary = True
        idx.Unique = True
        tdf.Indexes.Append idx

        ' Add table to database
        db.TableDefs.Append tdf

        Debug.Print "Logger: table " & cstrTableName & " created"
    End If

    Set idx = Nothing
    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "InitLogger error: " & Err.Description & " (" & Err.Number & ")"
    If Not idx Is Nothing Then Set idx = Nothing
    If Not fld Is Nothing Then Set fld = Nothing
    If Not tdf Is Nothing Then Set tdf = Nothing
    If Not db Is Nothing Then Set db = Nothing
End Sub

' =============================================
' @description Writes error to log. Always written (ERROR/INFO/DEBUG all log errors).
' @param sSource [String] Error source (module/procedure name)
' @param sMsg [String] Error message text
' @param bShowUI [Boolean] Whether to show MsgBox to user (default True)
' =============================================
Public Sub LogError(ByVal sSource As String, ByVal sMsg As String, Optional ByVal bShowUI As Boolean = True)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim strSQL As String
    Dim strWinUser As String
    Dim strSafeSource As String
    Dim strSafeMsg As String

    ' Get Windows username
    On Error Resume Next
    Dim objNetwork As Object
    Set objNetwork = CreateObject("WScript.Network")
    strWinUser = objNetwork.UserName
    If Err.Number <> 0 Then strWinUser = Environ("USERNAME")
    On Error GoTo ErrorHandler
    If strWinUser = "" Then strWinUser = "Unknown"

    ' Escape apostrophes for SQL
    strSafeSource = Replace(sSource, "'", "''")
    strSafeMsg = Replace(sMsg, "'", "''")

    ' Output to debug window
    Debug.Print "ERROR [" & sSource & "]: " & sMsg

    InitLogger

    Set db = CurrentDb
    strSQL = "INSERT INTO [" & cstrTableName & "] (LogType, Source, Description, WinUser, LogDate) " & _
             "VALUES ('ERROR', '" & strSafeSource & "', '" & strSafeMsg & "', '" & strWinUser & "', Now());"
    db.Execute strSQL, dbFailOnError

    If bShowUI Then
        MsgBox "An error occurred in module: " & sSource & vbCrLf & vbCrLf & _
               "Details: " & sMsg & vbCrLf & vbCrLf & _
               "The details were saved in the system log.", vbCritical, "System Error"
    End If

    Set db = Nothing
    Set objNetwork = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "LogError write failed: " & Err.Description & " (" & Err.Number & ")"
    Debug.Print "Original message [" & sSource & "]: " & sMsg
    If Not db Is Nothing Then Set db = Nothing
    If Not objNetwork Is Nothing Then Set objNetwork = Nothing
End Sub

' =============================================
' @description Returns current LogLevel from settings (ERROR, INFO, DEBUG). Default INFO.
' =============================================
Private Function GetLogLevel() As String
    On Error Resume Next
    GetLogLevel = UCase(Trim$(CStr(Nz(mod_Maintenance_Logic.GetSetting("LogLevel", "INFO"), "INFO"))))
    If Err.Number <> 0 Or Len(GetLogLevel) = 0 Then GetLogLevel = "INFO"
End Function

' =============================================
' @description Writes informational message to logging table. Only writes if LogLevel is INFO or DEBUG.
' @param sMsg [String] Informational message text
' @param sSource [String] Message source (default "General")
' =============================================
Public Sub LogInfo(ByVal sMsg As String, Optional ByVal sSource As String = "General")
    On Error GoTo ErrorHandler

    Dim strLevel As String
    strLevel = GetLogLevel
    If strLevel <> "INFO" And strLevel <> "DEBUG" Then Exit Sub

    Dim db As DAO.Database
    Dim strSQL As String
    Dim strWinUser As String
    Dim strSafeSource As String
    Dim strSafeMsg As String

    On Error Resume Next
    Dim objNetwork As Object
    Set objNetwork = CreateObject("WScript.Network")
    strWinUser = objNetwork.UserName
    If Err.Number <> 0 Then strWinUser = Environ("USERNAME")
    On Error GoTo ErrorHandler
    If strWinUser = "" Then strWinUser = "Unknown"

    strSafeSource = Replace(sSource, "'", "''")
    strSafeMsg = Replace(sMsg, "'", "''")

    Debug.Print "INFO [" & sSource & "]: " & sMsg

    InitLogger

    Set db = CurrentDb
    strSQL = "INSERT INTO [" & cstrTableName & "] (LogType, Source, Description, WinUser, LogDate) " & _
             "VALUES ('INFO', '" & strSafeSource & "', '" & strSafeMsg & "', '" & strWinUser & "', Now());"
    db.Execute strSQL, dbFailOnError

    Set db = Nothing
    Set objNetwork = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "LogInfo error: " & Err.Description & " (" & Err.Number & ")"
    Debug.Print "Original info [" & sSource & "]: " & sMsg
    If Not db Is Nothing Then Set db = Nothing
    If Not objNetwork Is Nothing Then Set objNetwork = Nothing
End Sub

' =============================================
' @description Writes debug message. Only written when LogLevel = "DEBUG".
' @param sMsg [String] Debug message text
' @param sSource [String] Message source (default "General")
' =============================================
Public Sub LogDebug(ByVal sMsg As String, Optional ByVal sSource As String = "General")
    On Error GoTo ErrorHandler

    If GetLogLevel() <> "DEBUG" Then Exit Sub

    Dim db As DAO.Database
    Dim strSQL As String
    Dim strWinUser As String
    Dim strSafeSource As String
    Dim strSafeMsg As String

    On Error Resume Next
    Dim objNetwork As Object
    Set objNetwork = CreateObject("WScript.Network")
    strWinUser = objNetwork.UserName
    If Err.Number <> 0 Then strWinUser = Environ("USERNAME")
    On Error GoTo ErrorHandler
    If strWinUser = "" Then strWinUser = "Unknown"

    strSafeSource = Replace(sSource, "'", "''")
    strSafeMsg = Replace(sMsg, "'", "''")
    Debug.Print "DEBUG [" & sSource & "]: " & sMsg

    InitLogger
    Set db = CurrentDb
    strSQL = "INSERT INTO [" & cstrTableName & "] (LogType, Source, Description, WinUser, LogDate) " & _
             "VALUES ('DEBUG', '" & strSafeSource & "', '" & strSafeMsg & "', '" & strWinUser & "', Now());"
    db.Execute strSQL, dbFailOnError

    Set db = Nothing
    If Not objNetwork Is Nothing Then Set objNetwork = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print "LogDebug error: " & Err.Description
    If Not db Is Nothing Then Set db = Nothing
    If Not objNetwork Is Nothing Then Set objNetwork = Nothing
End Sub
