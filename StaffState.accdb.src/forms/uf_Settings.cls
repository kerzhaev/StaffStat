Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================
' @form uf_Settings
' @description Settings UI with tab control: General, Mapping, Maintenance.
'              All settings read/write via mod_Maintenance_Logic.GetSetting/SetSetting.
' @note 100% English version. Safe for modern IDEs.
' =============================================

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    LoadSettingsFromStorage
    Me.cboTargetField.LimitToList = False
    RefreshMappingList
    Exit Sub
ErrorHandler:
    Debug.Print "uf_Settings Form_Load: " & Err.Number & " - " & Err.Description
End Sub

' Refresh mapping list when user switches to Mapping tab (pgMapping = page index 1).
Private Sub tabSettings_Change()
    On Error Resume Next
    If Me.tabSettings.value = 1 Then RefreshMappingList
End Sub

' Load all settings from tbl_Settings into controls.
Private Sub LoadSettingsFromStorage()
    On Error GoTo ErrorHandler

    Debug.Print "--- Loading Form Settings ---"
    Me.cboLogLevel.RowSource = "INFO;DEBUG;ERROR"

    Me.txtOrganizationName.value = Nz(mod_Maintenance_Logic.GetSetting("OrganizationName", "N/A"), "N/A")
    Debug.Print "Loaded Org: " & Me.txtOrganizationName.value

    Me.txtImportFolderPath.value = Nz(mod_Maintenance_Logic.GetSetting("ImportFolderPath", "N/A"), "N/A")
    Debug.Print "Loaded Path: " & Me.txtImportFolderPath.value

    Me.chkAutoCheckEnabled.value = CBool(Nz(mod_Maintenance_Logic.GetSetting("AutoCheckEnabled", "False"), "False"))
    Debug.Print "Loaded AutoCheck: " & Me.chkAutoCheckEnabled.value

    Me.cboLogLevel.value = CStr(Nz(mod_Maintenance_Logic.GetSetting("LogLevel", "INFO"), "INFO"))
    Debug.Print "Loaded LogLevel: " & Me.cboLogLevel.value

    Debug.Print "--- Settings Loaded: " & Me.txtOrganizationName.value & " ---"
    Exit Sub
ErrorHandler:
    Debug.Print "LoadSettingsFromStorage: " & Err.Description
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler
    Debug.Print "Attempting to save: " & Me.txtOrganizationName.value
    mod_Maintenance_Logic.SetSetting "OrganizationName", CStr(Nz(Me.txtOrganizationName.value, "")), "General"
    mod_Maintenance_Logic.SetSetting "ImportFolderPath", CStr(Nz(Me.txtImportFolderPath.value, "")), "General"
    mod_Maintenance_Logic.SetSetting "AutoCheckEnabled", IIf(Nz(Me.chkAutoCheckEnabled.value, False) = True, "True", "False"), "General"
    mod_Maintenance_Logic.SetSetting "LogLevel", CStr(Nz(Me.cboLogLevel.value, "INFO")), "General"

    mod_UI_Helpers.ShowMessage "Settings saved successfully.", vbInformation
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
ErrorHandler:
    Debug.Print "cmdSave_Click: " & Err.Number & " - " & Err.Description
    mod_UI_Helpers.ShowMessage "Save failed: " & Err.Description, vbExclamation
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

' --- Mapping tab (pgMapping): list is bound; Requery refreshes. ---
Public Sub RefreshMappingList()
    On Error GoTo ErrorHandler
    If Not (Me.lstMapping Is Nothing) Then
        Me.lstMapping.Requery
    End If
    Exit Sub
ErrorHandler:
    Debug.Print "RefreshMappingList: " & Err.Description
End Sub

Private Sub btnAddMapping_Click()
    On Error GoTo ErrorHandler
    Dim strExcel As String
    Dim strTarget As String
    Dim db As DAO.Database
    Dim strSQL As String

    strExcel = Trim$(Nz(Me.txtExcelHeader.value, ""))
    strTarget = Trim$(Nz(Me.cboTargetField.value, ""))

    If Len(strExcel) = 0 Or Len(strTarget) = 0 Then
        mod_UI_Helpers.ShowMessage "Please fill in both Excel Header and Target Field.", vbExclamation
        Exit Sub
    End If

    If Not mod_Schema_Manager.FieldExists("tbl_Personnel_Master", strTarget) Then
        mod_Schema_Manager.AddNewFieldToSchema strTarget
        mod_UI_Helpers.ShowMessage "Field [" & strTarget & "] created in the database structure.", vbInformation
    End If

    Set db = CurrentDb
    strSQL = "INSERT INTO tbl_Import_Mapping (ProfileID, ExcelHeader, TargetField) VALUES (1, '" & Replace(strExcel, "'", "''") & "', '" & Replace(strTarget, "'", "''") & "')"
    db.Execute strSQL, dbFailOnError
    Set db = Nothing

    RefreshMappingList
    Me.txtExcelHeader = ""
    Me.cboTargetField = ""
    Exit Sub
ErrorHandler:
    If Not db Is Nothing Then Set db = Nothing
    Debug.Print "btnAddMapping_Click: " & Err.Description
    mod_UI_Helpers.ShowMessage "Add mapping failed: " & Err.Description, vbExclamation
End Sub

Private Sub btnDeleteMapping_Click()
    On Error GoTo ErrorHandler
    Dim varId As Variant
    Dim strTarget As String
    Dim db As DAO.Database
    Dim strSQL As String

    If Me.lstMapping.ListIndex < 0 Then
        mod_UI_Helpers.ShowMessage "Please select a row to delete.", vbExclamation
        Exit Sub
    End If

    varId = Me.lstMapping.Column(0)
    strTarget = Nz(Me.lstMapping.Column(2), "")

    If UCase(Trim$(strTarget)) = "PERSONUID" Then
        If Not mod_UI_Helpers.AskUserYesNo("Removing PersonUID mapping may break the import process. Continue?", "Confirm") Then Exit Sub
    End If

    Set db = CurrentDb
    strSQL = "DELETE FROM tbl_Import_Mapping WHERE MappingID = " & CLng(varId)
    db.Execute strSQL, dbFailOnError
    Set db = Nothing

    RefreshMappingList
    Exit Sub
ErrorHandler:
    If Not db Is Nothing Then Set db = Nothing
    Debug.Print "btnDeleteMapping_Click: " & Err.Description
    mod_UI_Helpers.ShowMessage "Delete failed: " & Err.Description, vbExclamation
End Sub

Private Sub btnReSeedMapping_Click()
    On Error GoTo ErrorHandler
    If Not mod_UI_Helpers.AskUserYesNo("Restore default mapping? Current custom links will be deleted.", "Confirm") Then Exit Sub

    mod_Schema_Manager.ReSeedMapping
    RefreshMappingList
    mod_UI_Helpers.ShowMessage "Mapping restored to default.", vbInformation
    Exit Sub
ErrorHandler:
    Debug.Print "btnReSeedMapping_Click: " & Err.Description
    mod_UI_Helpers.ShowMessage "ReSeed failed: " & Err.Description, vbExclamation
End Sub

' --- Maintenance tab (pgMaintenance) ---
Private Sub cmdCreateBackup_Click()
    On Error GoTo ErrorHandler
    mod_Maintenance_Logic.CreateDatabaseBackup
    Exit Sub
ErrorHandler:
    Debug.Print "cmdCreateBackup_Click: " & Err.Description
End Sub

Private Sub cmdClearLogs_Click()
    On Error GoTo ErrorHandler
    mod_Maintenance_Logic.ClearValidationLog
    Exit Sub
ErrorHandler:
    Debug.Print "cmdClearLogs_Click: " & Err.Description
End Sub

Private Sub btnRunHealthCheck_Click()
    On Error GoTo ErrorHandler
    mod_Maintenance_Logic.RunDataHealthCheck False
    Exit Sub
ErrorHandler:
    Debug.Print "btnRunHealthCheck_Click: " & Err.Description
End Sub
