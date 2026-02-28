Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================
' @form uf_PersonCard
' @description Employee card form (view only, read-only data)
' @note 100% English code version. Safe for modern IDEs.
' =============================================

Private m_strUID As String

Private Sub Form_Open(Cancel As Integer)
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
        MsgBox "Employee with UID " & m_strUID & " not found.", vbExclamation, "StaffState"
        Exit Sub
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
    On Error Resume Next
    ' Check for Russian word "Dismissed" using ASCII-safe ChrW to prevent encoding breaks
    strDismissed = ChrW(1059) & ChrW(1074) & ChrW(1086) & ChrW(1083) & ChrW(1077) & ChrW(1085)
    If InStr(1, Nz(rs!WorkStatus, ""), strDismissed, vbTextCompare) > 0 Then
        Me.txtFullName.ForeColor = vbRed
    Else
        Me.txtFullName.ForeColor = 0
    End If
    On Error GoTo ErrorHandler

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error in LoadPersonData: " & Err.Description, vbCritical, "StaffState"
End Sub

' --- Helper function (safe load from RS to UI control) ---
Private Sub SafeLoad(rs As DAO.Recordset, strFieldName As String, ByRef ctrl As Control)
    On Error Resume Next
    Dim val As Variant
    val = rs.Fields(strFieldName).value

    If Err.Number <> 0 Then
        Debug.Print "DB Error: Field [" & strFieldName & "] not found in table."
        Err.Clear
        Exit Sub
    End If

    ctrl.value = Nz(val, "")

    If Err.Number <> 0 Then
        Debug.Print "UI Error: Control for [" & strFieldName & "] not found or incorrectly named on form."
        Err.Clear
    End If
End Sub

' --- History Filter Setup (History Tab) ---

Private Sub SetupHistoryFilters()
    On Error Resume Next
    Me.cboFilterHistory.RowSourceType = "Value List"
    Me.cboFilterHistory.RowSource = "'All';'RankName';'PosName';'WorkStatus';'FullName';'BirthDate'"
    Me.cboFilterHistory = "All"

    Me.lstHistory.ColumnCount = 4
    Me.lstHistory.ColumnWidths = "2cm;3cm;4cm;4cm"

    ApplyHistoryFilter
End Sub

Private Sub ApplyHistoryFilter()
    On Error Resume Next
    Dim strSQL As String
    Dim strWhere As String

    strWhere = " WHERE PersonUID = '" & m_strUID & "'"

    If Me.cboFilterHistory <> "All" Then
        strWhere = strWhere & " AND FieldName LIKE '*" & Me.cboFilterHistory & "*'"
    End If

    strSQL = "SELECT Format(ChangeDate, 'dd.mm.yyyy') AS ChangeDateFmt, FieldName, " & _
             "OldValue & " & ChrW(34) & " -> " & ChrW(34) & " & NewValue AS ChangeDesc " & _
             "FROM tbl_History_Log " & _
             strWhere & " ORDER BY ChangeDate DESC;"

    Me.lstHistory.RowSource = strSQL
End Sub

Private Sub cboFilterHistory_Change()
    ApplyHistoryFilter
End Sub

Private Sub btnResetDates_Click()
    Me.cboFilterHistory = "All"
    ApplyHistoryFilter
End Sub
