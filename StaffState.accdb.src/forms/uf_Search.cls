Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================
' @form uf_Search

Option Compare Database
Option Explicit

' True when form opened with OpenArgs MODE=DUPLICATES
Private m_fDuplicatesMode As Boolean
' Mode string from OpenArgs (e.g. "MODE=DUPLICATES"); empty = normal search
Private m_strMode As String
Private m_fDebugSearchLogging As Boolean

' --- FORM LOAD: delegate to OpenArgs handler ---
Private Sub Form_Load()
    On Error GoTo ErrorHandler

    ' --- PHASE 30: Apply UI Localization ---
    Me.btnSearch.Caption = mod_UI_Helpers.GetLoc("BTN_SEARCH")
    Me.btnClear.Caption = mod_UI_Helpers.GetLoc("BTN_CLEAR")
    Me.btnExportExcel.Caption = mod_UI_Helpers.GetLoc("BTN_EXPORT")

    m_fDuplicatesMode = False
    m_strMode = ""
    m_fDebugSearchLogging = ShouldWriteDebugSearchLog()
    HandleOpenArgsMode
    If Not m_fDuplicatesMode Then Me.Caption = mod_UI_Helpers.GetLoc("TITLE_SEARCH")

ExitHere:
    Exit Sub
ErrorHandler:
    Debug.Print "Form_Load (uf_Search) error: " & Err.Number & " - " & Err.Description
    m_fDuplicatesMode = False
    Resume ExitHere
End Sub

' --- Entry point ---
Private Sub HandleOpenArgsMode()
    On Error GoTo ErrorHandler

    Dim strArgs As String

    strArgs = Nz(Me.OpenArgs, "")
    If IsDuplicatesMode(strArgs) Then
        LoadDuplicatesModeAndNotify
    End If

ExitHere:
    Exit Sub
ErrorHandler:
    Debug.Print "HandleOpenArgsMode error: " & Err.Number & " - " & Err.Description
    Resume ExitHere
End Sub

Private Function IsDuplicatesMode(ByVal strOpenArgs As String) As Boolean
    On Error GoTo ErrorHandler
    IsDuplicatesMode = (InStr(1, Nz(strOpenArgs, ""), "MODE=DUPLICATES", vbTextCompare) > 0)
    Exit Function
ErrorHandler:
    IsDuplicatesMode = False
    Debug.Print "IsDuplicatesMode error: " & Err.Number & " - " & Err.Description
End Function

Private Sub LoadDuplicatesModeAndNotify()
    On Error GoTo ErrorHandler

    Const cSqlGroups As String = "SELECT Count(*) AS GroupCount FROM (SELECT FullName, BirthDate FROM tbl_Personnel_Master WHERE FullName Is Not Null AND Len(Trim(FullName))>0 AND BirthDate Is Not Null GROUP BY FullName, BirthDate HAVING Count(*)>1) AS d;"
    Const cSqlRecords As String = "SELECT Count(*) AS RecordCount FROM tbl_Personnel_Master AS pm INNER JOIN (SELECT FullName, BirthDate FROM tbl_Personnel_Master WHERE FullName Is Not Null AND Len(Trim(FullName))>0 AND BirthDate Is Not Null GROUP BY FullName, BirthDate HAVING Count(*)>1) AS d ON pm.FullName=d.FullName AND pm.BirthDate=d.BirthDate;"

    Dim strSQL As String
    Dim lngGroups As Long
    Dim lngRecords As Long

    m_fDuplicatesMode = True
    m_strMode = "MODE=DUPLICATES"
    Me.txtFilter.value = ""
    Me.txtFilter.Enabled = False
    Me.Caption = mod_UI_Helpers.GetLoc("TITLE_DUPLICATES")
    On Error Resume Next
    Me.btnExportExcel.Enabled = False
    On Error GoTo ErrorHandler

    strSQL = BuildDuplicatesSQL()
    If strSQL <> "" Then
        Me.lstResults.RowSource = ""
        Me.lstResults.RowSource = strSQL
        Me.lstResults.Requery
    End If
    Me.btnExportExcel.Enabled = (Me.lstResults.ListCount > 0)

    lngGroups = ExecuteScalarLong(cSqlGroups)
    lngRecords = ExecuteScalarLong(cSqlRecords)
    Debug.Print "Duplicates mode: groups=" & lngGroups & ", records=" & lngRecords

    If lngGroups = 0 Then
        MsgBox mod_UI_Helpers.GetLoc("MSG_NO_DUPLICATES"), vbInformation, mod_UI_Helpers.GetLoc("TITLE_INFO")
    Else
        MsgBox mod_UI_Helpers.GetLoc("MSG_DUPLICATES_FOUND") & " " & _
               mod_UI_Helpers.GetLoc("LBL_GROUPS") & " = " & lngGroups & ", " & _
               mod_UI_Helpers.GetLoc("LBL_RECORDS") & " = " & lngRecords & ".", _
               vbInformation, mod_UI_Helpers.GetLoc("TITLE_INFO")
    End If

ExitHere:
    Exit Sub
ErrorHandler:
    Debug.Print "LoadDuplicatesModeAndNotify error: " & Err.Number & " - " & Err.Description
    m_fDuplicatesMode = False
    Resume ExitHere
End Sub

Private Function ExecuteScalarLong(ByVal strSQL As String) As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim v As Variant

    ExecuteScalarLong = 0
    If Trim$(strSQL) = "" Then Exit Function

    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If Not rs.EOF Then
        v = rs.Fields(0).value
        If Not IsNull(v) Then ExecuteScalarLong = CLng(v)
    End If

ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Exit Function
ErrorHandler:
    Debug.Print "ExecuteScalarLong error: " & Err.Number & " - " & Err.Description & " | SQL: " & Left$(strSQL, 80)
    ExecuteScalarLong = 0
    Resume ExitHere
End Function

Private Sub btnSearch_Click()
    If m_fDebugSearchLogging Then
        DebugLog "uf_Search.cls:btnSearch_Click", "btnSearch_Click entry", "H4", "{""event"":""btnSearch_Click""}"
    End If
    PerformSearch
End Sub

Private Sub btnClear_Click()
    On Error GoTo ErrorHandler

    If m_fDuplicatesMode Or (m_strMode <> "") Then
        m_fDuplicatesMode = False
        m_strMode = ""
        Me.lstResults.RowSource = ""
        Me.lstResults.Requery
        Me.txtFilter.Enabled = True
        Me.txtFilter.value = ""
        Me.Caption = mod_UI_Helpers.GetLoc("TITLE_SEARCH")
        On Error Resume Next
        Me.btnExportExcel.Enabled = False
        Me.btnSearch.Enabled = True
        Me.btnClear.Enabled = True
        On Error GoTo ErrorHandler
        Me.txtFilter.SetFocus
    Else
        Me.txtFilter.value = ""
        Me.lstResults.RowSource = ""
        Me.lstResults.Requery
        Me.txtFilter.SetFocus
    End If
    Exit Sub
ErrorHandler:
    Debug.Print "btnClear_Click error: " & Err.Number & " - " & Err.Description
End Sub

Private Sub txtFilter_Change()
    If m_fDuplicatesMode Then Exit Sub
    If m_fDebugSearchLogging Then
        DebugLog "uf_Search.cls:txtFilter_Change", "txtFilter_Change entry", "H4", "{""event"":""txtFilter_Change""}"
    End If
    PerformSearch bSilent:=True
End Sub

Private Sub txtFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        PerformSearch
        KeyCode = 0
    End If
End Sub

Private Sub txtFilter_KeyUp(KeyCode As Integer, Shift As Integer)
    If m_fDuplicatesMode Then Exit Sub
    If KeyCode <> 13 Then
        If m_fDebugSearchLogging Then
            DebugLog "uf_Search.cls:txtFilter_KeyUp", "txtFilter_KeyUp entry", "H4", "{""event"":""txtFilter_KeyUp"",""keyCode"":" & CStr(KeyCode) & "}"
        End If
        PerformSearch bSilent:=True
    End If
End Sub

Private Sub lstResults_DblClick(Cancel As Integer)
    On Error GoTo ErrorHandler

    Dim strUID As Variant

    If Me.lstResults.ItemsSelected.count = 0 Then Exit Sub
    strUID = Me.lstResults.value

    If IsNull(strUID) Or strUID = "" Then
        MsgBox mod_UI_Helpers.GetLoc("ERR_EMP_ID_NOT_READ"), vbExclamation, mod_UI_Helpers.GetLoc("TITLE_WARNING")
        Exit Sub
    End If

    DoCmd.OpenForm "uf_PersonCard", , , , , acDialog, CStr(strUID)

    Exit Sub
ErrorHandler:
    MsgBox mod_UI_Helpers.GetLoc("TITLE_ERROR") & ": " & Err.Description, vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
End Sub

Private Sub btnExportExcel_Click()
    On Error GoTo ErrorHandler
    Dim result As Object

    If m_fDuplicatesMode Then
        ExportDuplicatesToExcel
        Exit Sub
    End If

    Set result = mod_Export_Logic.ExportFullSearchToExcelResult()
    If Len(Trim$(CStr(Nz(result("Message"), "")))) > 0 Then
        mod_UI_Helpers.ShowMessage CStr(result("Message")), IIf(CBool(result("Success")), vbInformation, vbExclamation)
    End If
    Set result = Nothing

    Exit Sub
ErrorHandler:
    Set result = Nothing
    MsgBox mod_UI_Helpers.GetLoc("TITLE_ERROR") & ": " & Err.Description, vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
End Sub

Private Sub ExportDuplicatesToExcel()
    On Error GoTo ErrorHandler

    Const cQdfName As String = "qtmpExportDuplicates"

    Dim strSQL   As String
    Dim filePath As String

    If Me.lstResults.ListCount = 0 Then
        MsgBox mod_UI_Helpers.GetLoc("MSG_EXPORT_EMPTY"), vbExclamation, mod_UI_Helpers.GetLoc("TITLE_WARNING")
        Exit Sub
    End If

    strSQL = BuildDuplicatesExportSql()
    If strSQL = "" Then
        Debug.Print "ExportDuplicatesToExcel: BuildDuplicatesExportSql returned empty"
        MsgBox mod_UI_Helpers.GetLoc("ERR_EXPORT_PREP"), vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
        Exit Sub
    End If

    SaveOrReplaceQueryDef cQdfName, strSQL

    filePath = CurrentProject.Path & "\Duplicates_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, cQdfName, filePath, True

    MsgBox mod_UI_Helpers.GetLoc("MSG_EXPORT_SUCCESS") & " " & filePath, vbInformation, mod_UI_Helpers.GetLoc("TITLE_INFO")

ExitHere:
    Exit Sub
ErrorHandler:
    Debug.Print "ExportDuplicatesToExcel error: " & Err.Number & " - " & Err.Description
    MsgBox mod_UI_Helpers.GetLoc("TITLE_ERROR") & ": " & Err.Description, vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
    Resume ExitHere
End Sub

Private Function BuildDuplicatesExportSql() As String
    On Error GoTo ErrorHandler

    BuildDuplicatesExportSql = "SELECT pm.PersonUID, pm.FullName, pm.BirthDate, pm.RankName, pm.PosName, pm.IsActive " & _
        "FROM tbl_Personnel_Master AS pm " & _
        "INNER JOIN (" & _
        "SELECT FullName, BirthDate FROM tbl_Personnel_Master " & _
        "WHERE FullName Is Not Null AND Len(Trim(FullName))>0 AND BirthDate Is Not Null " & _
        "GROUP BY FullName, BirthDate HAVING Count(*)>1" & _
        ") AS d ON pm.FullName = d.FullName AND pm.BirthDate = d.BirthDate " & _
        "ORDER BY pm.FullName, pm.BirthDate, pm.PersonUID;"

    Exit Function
ErrorHandler:
    BuildDuplicatesExportSql = ""
    Debug.Print "BuildDuplicatesExportSql error: " & Err.Number & " - " & Err.Description
End Function

Private Sub SaveOrReplaceQueryDef(ByVal strName As String, ByVal strSQL As String)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database

    If Trim$(strName) = "" Or Trim$(strSQL) = "" Then Exit Sub

    Set db = CurrentDb
    On Error Resume Next
    db.QueryDefs.Delete strName
    On Error GoTo ErrorHandler
    db.CreateQueryDef strName, strSQL

ExitHere:
    Set db = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print "SaveOrReplaceQueryDef error: " & Err.Number & " - " & Err.Description
    Resume ExitHere
End Sub

Private Sub PerformSearch(Optional bSilent As Boolean = False)
    On Error GoTo ErrorHandler

    If m_fDuplicatesMode Then Exit Sub

    Dim strSQL As String

    strSQL = GetSearchSQL(bAddTop50:=True)

    If m_fDebugSearchLogging Then
        DebugLog "uf_Search.cls:PerformSearch", "PerformSearch after GetSearchSQL", "H2", "{""bSilent"":" & LCase$(CStr(bSilent)) & ",""sqlLen"":" & CStr(Len(strSQL)) & ",""hasSQL"":" & LCase$(CStr(strSQL <> "")) & "}"
    End If

    If strSQL = "" Then
        Me.lstResults.RowSource = ""
        Me.lstResults.Requery
        Exit Sub
    End If

    Me.lstResults.RowSourceType = "Table/Query"
    Me.lstResults.RowSource = ""
    Me.lstResults.RowSource = strSQL
    Me.lstResults.Requery

    If m_fDebugSearchLogging Then
        DebugLog "uf_Search.cls:PerformSearch", "PerformSearch after Requery", "H3", "{""listCount"":" & CStr(Me.lstResults.ListCount) & "}"
    End If

    If Not bSilent And Me.lstResults.ListCount = 0 Then
        MsgBox mod_UI_Helpers.GetLoc("MSG_NO_EMPLOYEES"), vbExclamation, mod_UI_Helpers.GetLoc("TITLE_SEARCH")
    End If

    Exit Sub
ErrorHandler:
    If Not bSilent Then
        MsgBox mod_UI_Helpers.GetLoc("TITLE_ERROR") & ": " & Err.Description, vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
    End If
    If m_fDebugSearchLogging Then
        DebugLog "uf_Search.cls:PerformSearch", "PerformSearch error", "H2", "{""errNumber"":" & CStr(Err.Number) & ",""hasError"":" & LCase$(CStr(Err.Number <> 0)) & "}"
    End If
End Sub

Public Function GetCurrentFilterText() As String
    On Error Resume Next
    GetCurrentFilterText = Trim$(Nz(Me.txtFilter.Text, ""))
    If Err.Number <> 0 Or GetCurrentFilterText = "" Then
        GetCurrentFilterText = Trim$(Nz(Me.txtFilter.value, ""))
    End If
    If m_fDebugSearchLogging Then
        DebugLog "uf_Search.cls:GetCurrentFilterText", "GetCurrentFilterText result", "H1", "{""textErr"":" & CStr(Err.Number) & ",""filterLen"":" & CStr(Len(GetCurrentFilterText)) & "}"
    End If
End Function

Public Function GetSearchSQL(Optional ByVal bAddTop50 As Boolean = False) As String
    On Error GoTo ErrorHandler

    Dim strSQL As String
    Dim strFilter As String
    Dim selCols As String
    Dim whereClause As String

    strFilter = GetCurrentFilterText()
    strFilter = Replace(strFilter, "'", "''")

    If Len(strFilter) > 0 And Len(strFilter) < 2 Then
        GetSearchSQL = ""
        Exit Function
    End If

    If bAddTop50 Then
        selCols = "TOP 50 [PersonUID], [FullName], [RankName], [PosName]"
    Else
        selCols = "[PersonUID], [FullName], [RankName], [PosName], [SourceID]"
    End If

    If Trim$(strFilter) = "" Then
        whereClause = "1=1"
    Else
        whereClause = "(Nz([PersonUID],"""") LIKE '*" & strFilter & "*') " & _
            "OR (Nz([FullName],"""") LIKE '*" & strFilter & "*') " & _
            "OR (Nz([RankName],"""") LIKE '*" & strFilter & "*') " & _
            "OR (Nz([PosName],"""") LIKE '*" & strFilter & "*') " & _
            "OR (CStr(Nz([SourceID],"""")) LIKE '*" & strFilter & "*')"
    End If

    strSQL = "SELECT " & selCols & " " & _
        "FROM tbl_Personnel_Master " & _
        "WHERE " & whereClause & " " & _
        "ORDER BY [FullName];"

    GetSearchSQL = strSQL
    Exit Function
ErrorHandler:
    GetSearchSQL = ""
    Debug.Print "GetSearchSQL error: " & Err.Number & " - " & Err.Description
End Function

' --- DEBUG LOGGING ---
Private Sub DebugLog(ByVal location As String, ByVal message As String, ByVal hypothesisId As String, ByVal dataJson As String)
    On Error GoTo ErrorHandler
    Dim f As Integer
    Dim logPath As String
    Dim json As String
    Dim ts As Currency
    Dim logId As String

    logPath = CurrentProject.Path & "\debug_search.log"
    ts = CLng(Timer * 1000)
    Randomize
    logId = "log_" & Format$(Now, "yyyymmddhhnnss") & "_" & CStr(Int(Rnd() * 100000))

    json = "{""id"":""" & logId & """,""timestamp"":" & CStr(ts) & ",""location"":""" & JsonEscape(location) & """,""message"":""" & JsonEscape(message) & """,""data"":" & dataJson & ",""sessionId"":""debug-session"",""runId"":""pre-fix"",""hypothesisId"":""" & JsonEscape(hypothesisId) & """}"

    f = FreeFile
    Open logPath For Append As #f
    Print #f, json
    Close #f
    Exit Sub
ErrorHandler:
    On Error Resume Next
    If f <> 0 Then Close #f
End Sub

Private Function ShouldWriteDebugSearchLog() As Boolean
    On Error GoTo ErrorHandler

    ShouldWriteDebugSearchLog = (UCase$(Trim$(CStr(Nz(mod_Maintenance_Logic.GetSetting("LogLevel", "INFO"), "INFO")))) = "DEBUG")
    Exit Function

ErrorHandler:
    ShouldWriteDebugSearchLog = False
End Function

Private Function JsonEscape(ByVal s As String) As String
    On Error GoTo ErrorHandler
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCrLf, "\n")
    t = Replace(t, vbCr, "\n")
    t = Replace(t, vbLf, "\n")
    JsonEscape = t
    Exit Function
ErrorHandler:
    JsonEscape = ""
End Function

Private Sub DebugInspectTableDef(ByVal tableName As String)
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim hasTable As Boolean
    Dim countTdf As Long

    On Error Resume Next
    Set db = CurrentDb
    If Err.Number <> 0 Or db Is Nothing Then
        DebugLog "uf_Search.cls:DebugInspectTableDef", "CurrentDb not available", "H2", "{""errNumber"":" & CStr(Err.Number) & ",""dbSet"":" & LCase$(CStr(Not db Is Nothing)) & "}"
        Err.Clear
        Exit Sub
    End If

    countTdf = db.TableDefs.count
    For Each tdf In db.TableDefs
        If tdf.Name = tableName Then
            hasTable = True
            Exit For
        End If
    Next tdf
    DebugLog "uf_Search.cls:DebugInspectTableDef", "TableDef inspect", "H2", "{""tableFound"":" & LCase$(CStr(hasTable)) & ",""tableCount"":" & CStr(countTdf) & "}"
    Exit Sub
ErrorHandler:
    DebugLog "uf_Search.cls:DebugInspectTableDef", "DebugInspectTableDef error", "H2", "{""errNumber"":" & CStr(Err.Number) & ",""hasError"":" & LCase$(CStr(Err.Number <> 0)) & "}"
End Sub

Private Function SafeFieldCount(ByVal tdf As DAO.TableDef, ByRef errNumber As Long) As Long
    On Error Resume Next
    errNumber = 0
    SafeFieldCount = 0
    If tdf Is Nothing Then
        errNumber = 91
        Exit Function
    End If
    SafeFieldCount = tdf.Fields.count
    If Err.Number <> 0 Then
        errNumber = Err.Number
        Err.Clear
    End If
End Function

' =============================================
' @description Builds SQL for duplicates mode
' =============================================
Private Function BuildDuplicatesSQL() As String
    On Error GoTo ErrorHandler

    ' Ignore blank names and null birth dates; group by FullName+BirthDate, HAVING Count>1
    BuildDuplicatesSQL = "SELECT pm.PersonUID, " & _
        "pm.FullName & "" ("" & Format(pm.BirthDate,""dd.mm.yyyy"") & "")"" AS FullName, " & _
        "pm.RankName, pm.PosName " & _
        "FROM tbl_Personnel_Master AS pm " & _
        "INNER JOIN (" & _
        "SELECT FullName, BirthDate FROM tbl_Personnel_Master " & _
        "WHERE FullName Is Not Null AND Len(Trim(FullName))>0 AND BirthDate Is Not Null " & _
        "GROUP BY FullName, BirthDate HAVING Count(*)>1" & _
        ") AS d ON pm.FullName = d.FullName AND pm.BirthDate = d.BirthDate " & _
        "ORDER BY pm.FullName, pm.BirthDate, pm.PersonUID;"

    Exit Function
ErrorHandler:
    BuildDuplicatesSQL = ""
    Debug.Print "BuildDuplicatesSQL error: " & Err.Number & " - " & Err.Description
End Function
