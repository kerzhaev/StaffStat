Attribute VB_Name = "mod_App_Logger"
Option Compare Database
Option Explicit

' =============================================
' @module mod_App_Logger
' @description Centralized error and event logging system (QueryDef Parameterized)
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

    Set db = CurrentDb

    On Error Resume Next
    Set tdf = db.TableDefs(cstrTableName)
    If Err.Number = 0 Then Exit Sub ' Table exists
    On Error GoTo ErrorHandler

    ' Create table via DAO
    Set tdf = db.CreateTableDef(cstrTableName)

    Set fld = tdf.CreateField("ID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("LogDate", dbDate)
    fld.defaultValue = "Now()"
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("LogType", dbText, 50)
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("Source", dbText, 255)
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("Description", dbMemo)
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("WinUser", dbText, 50)
    tdf.Fields.Append fld

    Set idx = tdf.CreateIndex("PRIMARYKEY")
    Set fld = idx.CreateField("ID")
    idx.Fields.Append fld
    idx.Primary = True
    idx.Unique = True
    tdf.Indexes.Append idx

    db.TableDefs.Append tdf
    Debug.Print "Logger: table " & cstrTableName & " created"

    Set idx = Nothing
    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "InitLogger error: " & Err.Description
End Sub

' =============================================
' @description Core write mechanism using Parameterized QueryDef
' =============================================
Private Sub WriteLogEntry(ByVal sType As String, ByVal sSource As String, ByVal sMsg As String)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim strSQL As String
    Dim strWinUser As String
    Dim objNetwork As Object

    On Error Resume Next
    Set objNetwork = CreateObject("WScript.Network")
    strWinUser = objNetwork.UserName
    If Err.Number <> 0 Then strWinUser = Environ("USERNAME")
    On Error GoTo ErrorHandler
    If strWinUser = "" Then strWinUser = "Unknown"

    InitLogger

    Set db = CurrentDb
    strSQL = "PARAMETERS prmType Text (50), prmSource Text (255), prmDesc Memo, prmUser Text (50); " & _
             "INSERT INTO [" & cstrTableName & "] (LogType, Source, Description, WinUser, LogDate) " & _
             "VALUES ([prmType], [prmSource], [prmDesc], [prmUser], Now());"

    Set qdf = db.CreateQueryDef("", strSQL)
    qdf.Parameters("prmType").value = sType
    qdf.Parameters("prmSource").value = Left$(sSource, 255)
    qdf.Parameters("prmDesc").value = sMsg
    qdf.Parameters("prmUser").value = Left$(strWinUser, 50)

    qdf.Execute dbFailOnError

    Set qdf = Nothing
    Set db = Nothing
    Set objNetwork = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "WriteLogEntry failed: " & Err.Description
End Sub

Private Function GetLogLevel() As String
    On Error Resume Next
    GetLogLevel = UCase(Trim$(CStr(Nz(mod_Maintenance_Logic.GetSetting("LogLevel", "INFO"), "INFO"))))
    If Err.Number <> 0 Or Len(GetLogLevel) = 0 Then GetLogLevel = "INFO"
End Function

Public Sub LogError(ByVal sSource As String, ByVal sMsg As String, Optional ByVal bShowUI As Boolean = True)
    Debug.Print "ERROR [" & sSource & "]: " & sMsg
    WriteLogEntry "ERROR", sSource, sMsg
    If bShowUI Then
        MsgBox "An error occurred in module: " & sSource & vbCrLf & vbCrLf & _
               "Details: " & sMsg & vbCrLf & vbCrLf & _
               "The details were saved in the system log.", vbCritical, "System Error"
    End If
End Sub

Public Sub LogInfo(ByVal sMsg As String, Optional ByVal sSource As String = "General")
    Dim strLevel As String
    strLevel = GetLogLevel()
    If strLevel <> "INFO" And strLevel <> "DEBUG" Then Exit Sub
    Debug.Print "INFO [" & sSource & "]: " & sMsg
    WriteLogEntry "INFO", sSource, sMsg
End Sub

Public Sub LogDebug(ByVal sMsg As String, Optional ByVal sSource As String = "General")
    If GetLogLevel() <> "DEBUG" Then Exit Sub
    Debug.Print "DEBUG [" & sSource & "]: " & sMsg
    WriteLogEntry "DEBUG", sSource, sMsg
End Sub
