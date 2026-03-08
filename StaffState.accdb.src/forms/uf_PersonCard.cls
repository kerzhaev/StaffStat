Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================
' @form uf_PersonCard
' @description Employee card form (view only, read-only data)
' @note 100% English code version. Explicit error handling (Phase 28).
' =============================================

Option Compare Database
Option Explicit

Private m_strUID As String

Private Sub Form_Open(Cancel As Integer)
    On Error Resume Next
    ' --- PHASE 30: Apply UI Localization ---
    Me.Caption = mod_UI_Helpers.GetLoc("TITLE_PERSON_CARD")

    If Not IsNull(Me.OpenArgs) Then
        m_strUID = Me.OpenArgs
        LoadPersonData
        SetupHistoryFilters
    End If
End Sub

' --- Load Person Data ---

Public Sub LoadPersonData()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim strDismissed As String

    Set db = CurrentDb
    strSQL = "SELECT * FROM tbl_Personnel_Master WHERE PersonUID = '" & Replace(m_strUID, "'", "''") & "'"
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)

    If rs.EOF Then
        MsgBox mod_UI_Helpers.GetLoc("ERR_EMP_NOT_FOUND") & " " & m_strUID, vbExclamation, mod_UI_Helpers.GetLoc("TITLE_WARNING")
        GoTo Cleanup
    End If

    ' Header
    SafeLoad rs, "FullName", Me.txtFullName
    SafeLoad rs, "PersonUID", Me.txtPersonUID
    SafeLoad rs, "RankName", Me.txtRank
    SafeLoad rs, "PosName", Me.txtPosition
    SafeLoad rs, "WorkStatus", Me.txtWorkStatus

    ' Tab 1: Service
    SafeLoad rs, "PersonnelDivision", Me.PersonnelDivision
    SafeLoad rs, "StaffPosition", Me.StaffPosition
    SafeLoad rs, "VUS", Me.VUS
    SafeLoad rs, "SalaryGrade", Me.SalaryGrade

    ' Tab 2: Contract
    SafeLoad rs, "ContractType", Me.ContractType
    SafeLoad rs, "ContractKind", Me.ContractKind
    SafeLoad rs, "ContractStartDate", Me.ContractStartDate
    SafeLoad rs, "ContractEndDate", Me.ContractEndDate

    ' Tab 3: Personal Info
    SafeLoad rs, "BirthDate_Text", Me.BirthDate
    SafeLoad rs, "EmployeeAge", Me.EmployeeAge
    SafeLoad rs, "Gender", Me.Gender
    SafeLoad rs, "Address", Me.Address

    ' Tab 4: Bank
    SafeLoad rs, "BankAccountNumber", Me.BankAccountNumber
    SafeLoad rs, "Payee", Me.Payee

    ' Visual styling based on status
    If HasField(rs, "WorkStatus") Then
        ' Get localized string from Dictionary instead of using ChrW()
        strDismissed = mod_UI_Helpers.GetLoc("STATUS_DISMISSED")
        If InStr(1, Nz(rs!WorkStatus, ""), strDismissed, vbTextCompare) > 0 Then
            Me.txtFullName.ForeColor = vbRed
        Else
            Me.txtFullName.ForeColor = 0
        End If
    End If

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    mod_App_Logger.LogError "uf_PersonCard.LoadPersonData", Err.Description, True
    Resume Cleanup
End Sub

' =============================================
' @description Safe load from RS to UI control with explicit checks
' =============================================
Private Sub SafeLoad(rs As DAO.Recordset, strFieldName As String, ByRef ctrl As Control)
    On Error GoTo ErrorHandler

    ' 1. Check if control exists
    If ctrl Is Nothing Then
        mod_App_Logger.LogDebug "Control for field [" & strFieldName & "] is missing on the form.", "uf_PersonCard"
        Exit Sub
    End If

    ' 2. Check if field exists in Recordset
    If Not HasField(rs, strFieldName) Then
        mod_App_Logger.LogDebug "Field [" & strFieldName & "] is missing from the database.", "uf_PersonCard"
        Exit Sub
    End If

    ' 3. Safe assignment
    ctrl.value = Nz(rs.Fields(strFieldName).value, "")
    Exit Sub

ErrorHandler:
    mod_App_Logger.LogError "uf_PersonCard.SafeLoad", "Error mapping [" & strFieldName & "]: " & Err.Description, False
End Sub

' =============================================
' @description Helper: Checks if a field exists in a DAO Recordset
' =============================================
Private Function HasField(rs As DAO.Recordset, strName As String) As Boolean
    On Error Resume Next
    Dim x As Variant
    x = rs.Fields(strName).Name
    HasField = (Err.Number = 0)
    Err.Clear
End Function

' --- History Filter Setup (History Tab) ---

Private Sub SetupHistoryFilters()
    On Error GoTo ErrorHandler

    Dim strAll As String
    strAll = mod_UI_Helpers.GetLoc("FILTER_ALL")

    Me.cboFilterHistory.RowSourceType = "Value List"
    Me.cboFilterHistory.RowSource = "'" & strAll & "';'RankName';'PosName';'WorkStatus';'FullName';'BirthDate'"
    Me.cboFilterHistory = strAll

    Me.lstHistory.ColumnCount = 4
    Me.lstHistory.ColumnWidths = "2cm;3cm;4cm;4cm"

    ApplyHistoryFilter
    Exit Sub
ErrorHandler:
    mod_App_Logger.LogError "uf_PersonCard.SetupHistoryFilters", Err.Description, False
End Sub

Private Sub ApplyHistoryFilter()
    On Error GoTo ErrorHandler
    Dim strSQL As String
    Dim strWhere As String
    Dim strAll As String

    strAll = mod_UI_Helpers.GetLoc("FILTER_ALL")

    strWhere = " WHERE PersonUID = '" & Replace(m_strUID, "'", "''") & "'"

    If Nz(Me.cboFilterHistory, strAll) <> strAll Then
        strWhere = strWhere & " AND FieldName LIKE '*" & Replace(Me.cboFilterHistory, "'", "''") & "*'"
    End If

    strSQL = "SELECT Format(ChangeDate, 'dd.mm.yyyy') AS ChangeDateFmt, FieldName, " & _
             "OldValue & " & ChrW(34) & " -> " & ChrW(34) & " & NewValue AS ChangeDesc " & _
             "FROM tbl_History_Log " & _
             strWhere & " ORDER BY ChangeDate DESC;"

    Me.lstHistory.RowSource = strSQL
    Exit Sub
ErrorHandler:
    mod_App_Logger.LogError "uf_PersonCard.ApplyHistoryFilter", Err.Description, False
End Sub

Private Sub cboFilterHistory_Change()
    ApplyHistoryFilter
End Sub

Private Sub btnResetDates_Click()
    Me.cboFilterHistory = mod_UI_Helpers.GetLoc("FILTER_ALL")
    ApplyHistoryFilter
End Sub
