Attribute VB_Name = "mod_UI_Helpers"
Option Compare Database
Option Explicit

' =============================================
' @module mod_UI_Helpers
' @description UI helper functions (friendly captions, history formatting)
' @note 100% English version. Safe for modern IDEs, Git and AI tools.
' =============================================
' =============================================
' LOCALIZATION ENGINE (Phase 30)
' =============================================
Private dictLocales As Object ' Scripting.Dictionary
' =============================================
' @description Shows a message to the user (wrapper for MsgBox).
' @param msg [String] Message text
' @param msgType [VbMsgBoxStyle] vbInformation, vbExclamation, etc. Default vbInformation
' =============================================
Public Sub ShowMessage(ByVal msg As String, Optional ByVal msgType As VbMsgBoxStyle = vbInformation)
    On Error GoTo ErrorHandler
    MsgBox msg, msgType, "StaffState"
    Exit Sub
ErrorHandler:
    Debug.Print "ShowMessage error: " & Err.Description
End Sub

' --- UI Messages and Captions ---

Public Function GetSearchCaption() As String
    GetSearchCaption = "Employee Search"
End Function

Public Function GetMsgBackupPathUndefined() As String
    GetMsgBackupPathUndefined = "Backup directory is not defined: the database is located in a path that cannot be accessed."
End Function

Public Function GetMsgBackupSaved(ByVal strPath As String) As String
    GetMsgBackupSaved = "Backup successfully saved to: " & strPath
End Function

Public Function GetMsgBackupError70() As String
    GetMsgBackupError70 = "File is locked. Cannot create a backup while the database is being used."
End Function

Public Function GetMsgBackupFailedLocked() As String
    GetMsgBackupFailedLocked = "Cannot backup while the database is heavily used. Please close other Access windows and try again."
End Function

Public Function GetMsgBackupFailedGeneric(ByVal strErrDesc As String) As String
    GetMsgBackupFailedGeneric = "Backup failed: " & strErrDesc & vbCrLf & vbCrLf & _
        "If the file is locked, try closing other Access windows or copy the file manually."
End Function

Public Function GetMsgValidationLogCleared() As String
    GetMsgValidationLogCleared = "Validation log has been cleared successfully."
End Function

Public Function GetMsgExportCompleted(ByVal recCount As Long) As String
    GetMsgExportCompleted = "Export completed successfully. Exported " & recCount & " record(s)."
End Function

' =============================================
' @description Asks user Yes/No; use for export/action prompts.
' @param msg [String] Question text
' @param title [String] Optional dialog title, default "StaffState"
' @return [Boolean] True if user chose Yes
' =============================================
Public Function AskUserYesNo(ByVal msg As String, Optional ByVal title As String = "StaffState") As Boolean
    On Error GoTo ErrorHandler
    AskUserYesNo = (MsgBox(msg, vbYesNo + vbQuestion, title) = vbYes)
    Exit Function
ErrorHandler:
    Debug.Print "AskUserYesNo error: " & Err.Description
    AskUserYesNo = False
End Function

' =============================================
' @description Returns "(empty)" token.
' =============================================
Private Function GetEmptyToken() As String
    GetEmptyToken = "(empty)"
End Function

' =============================================
' @description Returns English caption for internal field name.
' @param strInternalName [String] Internal field name (e.g., "RankName")
' @return [String] English UI caption
' =============================================
Public Function GetFieldFriendlyName(ByVal strInternalName As String) As String
    On Error GoTo ErrorHandler

    Dim s As String
    s = Trim$(Nz(strInternalName, ""))

    Select Case s
        Case "_System"
            GetFieldFriendlyName = "System Account"
        Case "RankName"
            GetFieldFriendlyName = "Rank"
        Case "WorkStatus"
            GetFieldFriendlyName = "Work Status"
        Case "PosName"
            GetFieldFriendlyName = "Position Name"
        Case "PosCode"
            GetFieldFriendlyName = "Position Code"
        Case "FullName"
            GetFieldFriendlyName = "Full Name"
        Case "PersonUID"
            GetFieldFriendlyName = "Personal ID"
        Case "SourceID"
            GetFieldFriendlyName = "Source ID"
        Case "OrderDate"
            GetFieldFriendlyName = "Order Date"
        Case "OrderNum"
            GetFieldFriendlyName = "Order Number"
        Case "BirthDate"
            GetFieldFriendlyName = "Date of Birth"
        Case Else
            GetFieldFriendlyName = Replace(s, "_", " ")
    End Select

    Exit Function
ErrorHandler:
    GetFieldFriendlyName = Trim$(Nz(strInternalName, ""))
End Function

' =============================================
' @description Translates system event token to English.
' @param sToken [String] Token stored in DB
' @return [String] English user-facing text
' =============================================
Private Function TranslateSystemToken(ByVal sToken As String) As String
    Dim s As String
    s = UCase$(Trim$(Nz(sToken, "")))

    Select Case s
        Case "ADDED"
            TranslateSystemToken = "Added to database"
        Case "REMOVED"
            TranslateSystemToken = "Removed from database"
        Case Else
            TranslateSystemToken = Trim$(Nz(sToken, ""))
    End Select
End Function

' =============================================
' @description Builds a human-friendly history description line.
' @param strInternalName [String] Field internal name
' @param vOld [Variant] OldValue
' @param vNew [Variant] NewValue
' @return [String] Human-friendly description string
' =============================================
Public Function BuildHistoryDescription(ByVal strInternalName As String, ByVal vOld As Variant, ByVal vNew As Variant) As String
    On Error GoTo ErrorHandler

    Dim sOld As String
    Dim sNew As String
    Dim sNewDisplay As String
    Dim sCaption As String
    Dim sMarker As String

    sOld = Trim$(Nz(vOld, ""))
    sNew = Trim$(Nz(vNew, ""))
    sCaption = GetFieldFriendlyName(strInternalName)

    If Trim$(Nz(strInternalName, "")) = "_System" Then
        sNewDisplay = TranslateSystemToken(sNew)
    Else
        sNewDisplay = sNew
    End If

    If sOld = "" And sNew <> "" Then
        sMarker = "[+]"
        sOld = GetEmptyToken()
    ElseIf sOld <> "" And sNew = "" Then
        sMarker = "[-]"
        sNewDisplay = GetEmptyToken()
    Else
        sMarker = "[*]"
        If sOld = "" Then sOld = GetEmptyToken()
        If sNewDisplay = "" Then sNewDisplay = GetEmptyToken()
    End If

    BuildHistoryDescription = sMarker & " " & sCaption & ": " & sOld & " -> " & sNewDisplay
    Exit Function

ErrorHandler:
    BuildHistoryDescription = "[*] " & Trim$(Nz(strInternalName, "")) & ": " & Trim$(Nz(vOld, "")) & " -> " & Trim$(Nz(vNew, ""))
End Function

' =============================================
' @description Loads all translations into memory
' =============================================
Public Sub InitLocalization()
    On Error GoTo ErrorHandler

    ' Initialize Dictionary (Late Binding)
    Set dictLocales = CreateObject("Scripting.Dictionary")
    dictLocales.CompareMode = 1 ' vbTextCompare (case-insensitive)

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb

    ' Safely attempt to open table
    On Error Resume Next
    Set rs = db.OpenRecordset("SELECT MsgKey, LocalValue FROM tbl_Localization", dbOpenForwardOnly)
    On Error GoTo ErrorHandler

    If Not rs Is Nothing Then
        Do While Not rs.EOF
            ' Load keys and values into memory
            dictLocales.Add Nz(rs!MsgKey, ""), Nz(rs!LocalValue, "")
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If

    Debug.Print "Localization Engine initialized. Loaded " & dictLocales.count & " keys."
    Exit Sub

ErrorHandler:
    Debug.Print "InitLocalization Error " & Err.Number & ": " & Err.Description
End Sub

' =============================================
' @description Fast UI text retriever
' @param strKey [String] The English message key
' @return [String] Localized string or [KEY] if missing
' =============================================
Public Function GetLoc(ByVal strKey As String) As String
    ' Auto-initialize if not loaded yet
    If dictLocales Is Nothing Then InitLocalization

    If dictLocales.Exists(strKey) Then
        GetLoc = dictLocales(strKey)
    Else
        ' Return bracketed key so developer can spot missing translations
        GetLoc = "[" & strKey & "]"
    End If
End Function

' =============================================
' @description Clears cache (useful after editing the table)
' =============================================
Public Sub ReloadLocalization()
    Set dictLocales = Nothing
    InitLocalization
End Sub
