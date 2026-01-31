Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Option Explicit



' --- SEARCH BUTTON ---
Private Sub btnSearch_Click()
    ' #Region agent log
    DebugLog "uf_Search.cls:btnSearch_Click", "btnSearch_Click entry", "H4", "{""event"":""btnSearch_Click""}"
    ' #End Region
    PerformSearch
End Sub

' --- CLEAR BUTTON ---
Private Sub btnClear_Click()
    On Error Resume Next
    Me.txtFilter.Value = ""
    Me.lstResults.RowSource = ""
    Me.txtFilter.SetFocus
End Sub

' --- AUTO-SEARCH ON TEXT CHANGE ---
Private Sub txtFilter_Change()
    ' Perform search automatically as user types (without showing "not found" message)
    ' #Region agent log
    DebugLog "uf_Search.cls:txtFilter_Change", "txtFilter_Change entry", "H4", "{""event"":""txtFilter_Change""}"
    ' #End Region
    PerformSearch bSilent:=True
End Sub

' --- ENTER KEY PRESS IN INPUT FIELD ---
Private Sub txtFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    ' If Enter pressed (Code 13)
    If KeyCode = 13 Then
        PerformSearch
        KeyCode = 0 ' Suppress standard Enter sound
    End If
End Sub

' --- KEY UP EVENT (More reliable for auto-search) ---
Private Sub txtFilter_KeyUp(KeyCode As Integer, Shift As Integer)
    ' Trigger search after key is released (text is already updated)
    ' Skip Enter key (already handled in KeyDown)
    If KeyCode <> 13 Then
        ' #Region agent log
        DebugLog "uf_Search.cls:txtFilter_KeyUp", "txtFilter_KeyUp entry", "H4", "{""event"":""txtFilter_KeyUp"",""keyCode"":" & CStr(KeyCode) & "}"
        ' #End Region
        PerformSearch bSilent:=True
    End If
End Sub

' --- DOUBLE CLICK ON LIST ---
Private Sub lstResults_DblClick(Cancel As Integer)
    On Error GoTo ErrorHandler

    Dim strUID As Variant ' Use Variant to catch Null

    ' Check if there's a selection
    If Me.lstResults.ItemsSelected.count = 0 Then Exit Sub

    ' Get value (it's taken from Bound Column = 1)
    strUID = Me.lstResults.Value

    ' DIAGNOSTICS: If empty - report it
    If IsNull(strUID) Or strUID = "" Then
        MsgBox "Warning: Employee ID was not read from the list." & vbCrLf & _
               "Check 'Column Count' (should be 4) and 'Bound Column' (should be 1).", vbExclamation
        Exit Sub
    End If

    ' If all ok - open card
    ' Pass ID in OpenArgs (last parameter)
    DoCmd.OpenForm "uf_PersonCard", , , , , acDialog, CStr(strUID)

    Exit Sub
ErrorHandler:
    MsgBox "Open error: " & Err.Description
End Sub

' --- EXPORT TO EXCEL ---
Private Sub btnExportExcel_Click()
    On Error GoTo ErrorHandler

    Dim strSQL   As String
    Dim strFilter As String
    Dim db       As DAO.Database
    Dim rs       As DAO.Recordset

    strFilter = GetCurrentFilterText()
    strSQL = BuildDynamicSearchSQL(strFilter, bAddTop50:=False)
    If strSQL = "" Then
        MsgBox "No data to export.", vbExclamation
        Exit Sub
    End If

    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        Set db = Nothing
        MsgBox "No data to export.", vbExclamation
        Exit Sub
    End If

    If mod_Export_Logic.ExportSearchToExcel(rs) Then
        ' Success (message shown inside ExportSearchToExcel)
    End If

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Export error: " & Err.Description, vbCritical
    GoTo Cleanup
End Sub

' --- SEARCH LOGIC (dynamic: TableDef.Fields, searchable types only) ---
' @param bSilent [Boolean] If True, suppress "not found" message (for auto-search)
Private Sub PerformSearch(Optional bSilent As Boolean = False)
    On Error GoTo ErrorHandler

    Dim strSQL    As String
    Dim strFilter As String

    strFilter = GetCurrentFilterText()
    ' #Region agent log
    DebugLog "uf_Search.cls:PerformSearch", "PerformSearch after GetCurrentFilterText", "H1", "{""bSilent"":" & LCase$(CStr(bSilent)) & ",""filterLen"":" & CStr(Len(strFilter)) & "}"
    ' #End Region

    ' Escape single quotes for SQL safety
    strFilter = Replace(strFilter, "'", "''")

    ' MINIMUM LENGTH CHECK: Start search only after 2 characters
    If Len(strFilter) > 0 And Len(strFilter) < 2 Then
        Me.lstResults.RowSource = ""
        Me.lstResults.Requery
        Exit Sub
    End If

    strSQL = BuildDynamicSearchSQL(strFilter, bAddTop50:=True)
    ' #Region agent log
    DebugLog "uf_Search.cls:PerformSearch", "PerformSearch after BuildDynamicSearchSQL", "H2", "{""sqlLen"":" & CStr(Len(strSQL)) & ",""hasSQL"":" & LCase$(CStr(strSQL <> "")) & "}"
    ' #End Region
    If strSQL = "" Then
        Me.lstResults.RowSource = ""
        Me.lstResults.Requery
        Exit Sub
    End If

    Me.lstResults.RowSource = ""
    Me.lstResults.RowSource = strSQL
    Me.lstResults.Requery
    ' #Region agent log
    DebugLog "uf_Search.cls:PerformSearch", "PerformSearch after Requery", "H3", "{""listCount"":" & CStr(Me.lstResults.ListCount) & "}"
    ' #End Region

    If Not bSilent And Me.lstResults.ListCount = 0 Then
        MsgBox "No employees found.", vbExclamation
    End If

    Exit Sub
ErrorHandler:
    If Not bSilent Then
        MsgBox "Search error: " & Err.Description, vbCritical
    End If
    ' #Region agent log
    DebugLog "uf_Search.cls:PerformSearch", "PerformSearch error", "H2", "{""errNumber"":" & CStr(Err.Number) & ",""hasError"":" & LCase$(CStr(Err.Number <> 0)) & "}"
    ' #End Region
End Sub

' --- Get filter text from txtFilter ---
Private Function GetCurrentFilterText() As String
    On Error Resume Next
    GetCurrentFilterText = Trim$(Nz(Me.txtFilter.Text, ""))
    If Err.Number <> 0 Or GetCurrentFilterText = "" Then
        GetCurrentFilterText = Trim$(Nz(Me.txtFilter.Value, ""))
    End If
    ' #Region agent log
    DebugLog "uf_Search.cls:GetCurrentFilterText", "GetCurrentFilterText result", "H1", "{""textErr"":" & CStr(Err.Number) & ",""filterLen"":" & CStr(Len(GetCurrentFilterText)) & "}"
    ' #End Region
End Function

' --- Build dynamic SELECT and WHERE from tbl_Personnel_Master schema ---
' Searchable types: Text(10), Memo(12), Long(4), Date(8). Blacklist: ID, SourceID, IsActive, LastUpdated, LogID, PersonUID_Raw.
' @param strFilter [String] Escaped search text (caller must replace ')
' @param bAddTop50 [Boolean] If True, add TOP 50 for list
' @return Full SQL or empty string on error
Private Function BuildDynamicSearchSQL(ByVal strFilter As String, ByVal bAddTop50 As Boolean) As String
    On Error GoTo ErrorHandler

    Dim db     As DAO.Database
    Dim tdf    As DAO.TableDef
    Dim fld    As DAO.Field
    Dim selParts  As String
    Dim whereParts As String
    Dim fName As String
    Dim dt     As Integer
    Dim preferred As String
    Dim rest   As String
    Dim n      As String
    Dim orderField As String
    Dim firstField As String
    Dim fc As Long
    Dim fcErr As Long

    DebugInspectTableDef "tbl_Personnel_Master"
    Set db = CurrentDb
    ' #Region agent log
    DebugLog "uf_Search.cls:BuildDynamicSearchSQL", "BuildDynamicSearchSQL after Set db", "H2", "{""dbSet"":" & LCase$(CStr(Not db Is Nothing)) & "}"
    ' #End Region
    Set tdf = db.TableDefs("tbl_Personnel_Master")
    ' #Region agent log
    DebugLog "uf_Search.cls:BuildDynamicSearchSQL", "BuildDynamicSearchSQL after Set tdf", "H2", "{""tableName"":""" & JsonEscape(tdf.Name) & """}"
    ' #End Region
    fc = SafeFieldCount(tdf, fcErr)
    ' #Region agent log
    DebugLog "uf_Search.cls:BuildDynamicSearchSQL", "BuildDynamicSearchSQL field count", "H2", "{""fieldCount"":" & CStr(fc) & ",""fcErr"":" & CStr(fcErr) & "}"
    ' #End Region
    If fcErr <> 0 Then GoTo ErrorHandler
    preferred = ""
    rest = ""
    orderField = ""
    firstField = ""

    ' Build SELECT: prefer PersonUID, FullName, RankName, PosName first; then rest (non-blacklist)
    For Each fld In tdf.Fields
        fName = fld.Name
        If mod_Export_Logic.IsTechnicalField(fName) Then GoTo NextField
        n = "[" & fName & "]"
        If firstField = "" Then firstField = n
        Select Case fName
            Case "PersonUID", "FullName", "RankName", "PosName"
                If preferred <> "" Then preferred = preferred & ", "
                preferred = preferred & n
                If fName = "FullName" Then orderField = n
            Case Else
                If rest <> "" Then rest = rest & ", "
                rest = rest & n
        End Select
NextField:
    Next fld

    selParts = preferred
    If rest <> "" Then
        If selParts <> "" Then selParts = selParts & ", "
        selParts = selParts & rest
    End If
    If selParts = "" Then Exit Function

    BuildDynamicSearchSQL = "SELECT "
    If bAddTop50 Then BuildDynamicSearchSQL = BuildDynamicSearchSQL & "TOP 50 "
    BuildDynamicSearchSQL = BuildDynamicSearchSQL & selParts & " FROM tbl_Personnel_Master"

    ' Build WHERE: only searchable types (Text=10, Memo=12, Long=4, Date=8)
    If Trim$(strFilter) <> "" Then
        whereParts = ""
        For Each fld In tdf.Fields
            fName = fld.Name
            If mod_Export_Logic.IsTechnicalField(fName) Then GoTo NextW
            dt = fld.Type
            Select Case dt
                Case 10, 12  ' dbText, dbMemo
                    n = "([" & fName & "] LIKE '*" & strFilter & "*')"
                Case 4       ' dbLong
                    n = "(CStr([" & fName & "]) LIKE '*" & strFilter & "*')"
                Case 8       ' dbDate
                    n = "(Format([" & fName & "],'yyyy-mm-dd') LIKE '*" & strFilter & "*')"
                Case Else
                    GoTo NextW
            End Select
            If whereParts <> "" Then whereParts = whereParts & " OR "
            whereParts = whereParts & n
NextW:
        Next fld
        If whereParts <> "" Then
            BuildDynamicSearchSQL = BuildDynamicSearchSQL & " WHERE " & whereParts
        End If
    End If

    If orderField = "" Then orderField = firstField
    If orderField <> "" Then
        BuildDynamicSearchSQL = BuildDynamicSearchSQL & " ORDER BY " & orderField & ";"
    End If
    ' #Region agent log
    DebugLog "uf_Search.cls:BuildDynamicSearchSQL", "BuildDynamicSearchSQL exit", "H2", "{""sqlLen"":" & CStr(Len(BuildDynamicSearchSQL)) & ",""orderFieldSet"":" & LCase$(CStr(orderField <> "")) & "}"
    ' #End Region

Cleanup:
    On Error Resume Next
    Set tdf = Nothing
    Exit Function

ErrorHandler:
    BuildDynamicSearchSQL = ""
    ' #Region agent log
    DebugLog "uf_Search.cls:BuildDynamicSearchSQL", "BuildDynamicSearchSQL error", "H2", "{""errNumber"":" & CStr(Err.Number) & ",""hasError"":" & LCase$(CStr(Err.Number <> 0)) & ",""errDesc"":""" & JsonEscape(Err.Description) & """}"
    ' #End Region
    GoTo Cleanup
End Function

' --- DEBUG LOGGING (agent) ---
Private Sub DebugLog(ByVal location As String, ByVal message As String, ByVal hypothesisId As String, ByVal dataJson As String)
    On Error GoTo ErrorHandler
    Dim f As Integer
    Dim logPath As String
    Dim json As String
    Dim ts As Currency
    Dim logId As String

    logPath = "c:\Users\Nachfin\Desktop\Projets\StaffState\.cursor\debug.log"
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
