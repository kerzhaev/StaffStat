Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Option Explicit

' =============================================
' @form uf_Search
' @author Кержаев Евгений
' @description Employee search form
' =============================================

' --- SEARCH BUTTON ---
Private Sub btnSearch_Click()
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
        PerformSearch bSilent:=True
    End If
End Sub

' --- DOUBLE CLICK ON LIST ---
Private Sub lstResults_DblClick(Cancel As Integer)
    On Error GoTo ErrorHandler

    Dim strUID As Variant ' Use Variant to catch Null

    ' Check if there's a selection
    If Me.lstResults.ItemsSelected.Count = 0 Then Exit Sub

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

' --- SEARCH LOGIC ---
' @param bSilent [Boolean] If True, suppress "not found" message (for auto-search)
Private Sub PerformSearch(Optional bSilent As Boolean = False)
    On Error GoTo ErrorHandler

    Dim strSQL As String
    Dim strFilter As String

    ' Get current text (use .Text for real-time, .Value as fallback)
    On Error Resume Next
    strFilter = Me.txtFilter.Text
    If Err.Number <> 0 Or IsNull(strFilter) Then
        strFilter = Nz(Me.txtFilter.Value, "")
    End If
    On Error GoTo ErrorHandler
    
    strFilter = Trim(strFilter)
    
    ' Escape single quotes for SQL safety (replace ' with '')
    strFilter = Replace(strFilter, "'", "''")

    ' MINIMUM LENGTH CHECK: Start search only after 2 characters
    If Len(strFilter) > 0 And Len(strFilter) < 2 Then
        Me.lstResults.RowSource = ""
        Me.lstResults.Requery
        Exit Sub
    End If

    ' Build base SQL query
    ' IMPORTANT: In lstResults the "Row Source Type" property must be "Table/Query"
    strSQL = "SELECT TOP 50 PersonUID, FullName, RankName, PosName " & _
             "FROM tbl_Personnel_Master"

    ' Add filter conditions (with proper WHERE clause)
    If strFilter <> "" Then
        strSQL = strSQL & " WHERE (FullName LIKE '*" & strFilter & "*') " & _
                 "OR (PersonUID LIKE '*" & strFilter & "*') " & _
                 "OR (CStr(SourceID) LIKE '*" & strFilter & "*')"
    End If

    ' Sort
    strSQL = strSQL & " ORDER BY FullName;"

    ' Clear list first (fix for Access caching issues)
    Me.lstResults.RowSource = ""
    
    ' Assign new SQL to list
    Me.lstResults.RowSource = strSQL
    
    ' Force list refresh
    Me.lstResults.Requery

    ' If nothing found (show message only if not silent mode)
    If Not bSilent And Me.lstResults.ListCount = 0 Then
        MsgBox "No employees found.", vbExclamation
    End If

    Exit Sub
ErrorHandler:
    ' Show errors only in non-silent mode
    If Not bSilent Then
        MsgBox "Search error: " & Err.Description, vbCritical
    End If
End Sub
