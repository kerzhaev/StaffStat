Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================
' @form uf_Settings
' @description Settings UI with tab control: General, Mapping, Maintenance.
'              All settings read/write via mod_Maintenance_Logic.GetSetting/SetSetting.
' @note 100% English version. Safe for modern IDEs.
' =============================================

Option Compare Database
Option Explicit

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    ' --- PHASE 30: Apply UI Localization ---
    Me.Caption = mod_UI_Helpers.GetLoc("TITLE_SETTINGS")

    ' Wrap control caption assignment in Resume Next
    ' in case some buttons/tabs are renamed or missing
    On Error Resume Next
    Me.cmdSave.Caption = mod_UI_Helpers.GetLoc("BTN_SAVE")
    Me.cmdCancel.Caption = mod_UI_Helpers.GetLoc("BTN_CANCEL")

    Me.btnAddMapping.Caption = mod_UI_Helpers.GetLoc("BTN_ADD_MAPPING")
    Me.btnDeleteMapping.Caption = mod_UI_Helpers.GetLoc("BTN_DEL_MAPPING")
    Me.btnEditMapping.Caption = mod_UI_Helpers.GetLoc("BTN_EDIT_MAPPING")
    Me.btnReSeedMapping.Caption = mod_UI_Helpers.GetLoc("BTN_RESEED_MAPPING")

    Me.cmdCreateBackup.Caption = mod_UI_Helpers.GetLoc("BTN_CREATE_BACKUP")
    Me.cmdClearLogs.Caption = mod_UI_Helpers.GetLoc("BTN_CLEAR_LOGS")
    Me.btnRunHealthCheck.Caption = mod_UI_Helpers.GetLoc("BTN_RUN_HEALTH_CHECK")

    ' Перевод новых лейблов маппинга
    Me.lblExcel.Caption = mod_UI_Helpers.GetLoc("LBL_EXCEL_HEADER")
    Me.lblDB.Caption = mod_UI_Helpers.GetLoc("LBL_DB_FIELD")

    ' Localize Tabs
    Me.pgGeneral.Caption = mod_UI_Helpers.GetLoc("TAB_GENERAL")
    Me.pgMapping.Caption = mod_UI_Helpers.GetLoc("TAB_MAPPING")
    Me.pgMaintenance.Caption = mod_UI_Helpers.GetLoc("TAB_MAINTENANCE")
    On Error GoTo ErrorHandler

    ' --- PHASE 31: Initialize Profile Combo Box ---
    On Error Resume Next
    Me.cboProfile.RowSourceType = "Table/Query"
    Me.cboProfile.RowSource = "SELECT ProfileID, ProfileID & ' - ' & ProfileName FROM tbl_Import_Profiles ORDER BY ProfileID;"
    Me.cboProfile.ColumnCount = 2
    Me.cboProfile.ColumnWidths = "0cm;5cm"
    Me.cboProfile.BoundColumn = 1
    If Me.cboProfile.ListCount > 0 Then Me.cboProfile.value = Me.cboProfile.ItemData(0) ' Выбираем первый профиль по умолчанию
    On Error GoTo ErrorHandler

    LoadSettingsFromStorage
    Me.cboTargetField.LimitToList = False
    RefreshMappingList
    Exit Sub
ErrorHandler:
    Debug.Print "uf_Settings Form_Load: " & Err.Number & " - " & Err.Description
End Sub

' --- PHASE 31: Change mapping list when profile is selected ---
Private Sub cboProfile_Change()
    RefreshMappingList
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

    mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetLoc("MSG_SETTINGS_SAVED"), vbInformation
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
ErrorHandler:
    Debug.Print "cmdSave_Click: " & Err.Number & " - " & Err.Description
    mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetLoc("ERR_SAVE_FAILED") & " " & Err.Description, vbExclamation
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

' --- Mapping tab (pgMapping): list is bound; Requery refreshes. ---
Public Sub RefreshMappingList()
    On Error GoTo ErrorHandler
    Dim lngProfile As Long

    ' Получаем выбранный профиль (если комбобокс пуст, используем 1)
    On Error Resume Next
    lngProfile = Nz(Me.cboProfile.value, 1)
    On Error GoTo ErrorHandler

    If Not (Me.lstMapping Is Nothing) Then
        ' Inject our custom function to show Data Type in the 4th column. Filter by ProfileID!
        Me.lstMapping.RowSource = "SELECT MappingID, ExcelHeader, TargetField, GetFieldTypeFriendly('tbl_Personnel_Master', TargetField) AS DataType FROM tbl_Import_Mapping WHERE ProfileID = " & lngProfile & " ORDER BY ExcelHeader;"
        Me.lstMapping.ColumnCount = 4
        Me.lstMapping.ColumnWidths = "0;2500;2500;2000"
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
    Dim qdf As DAO.QueryDef
    Dim strSQL As String
    Dim strTypeSelection As String
    Dim strSqlType As String
    Dim lngProfile As Long

    strExcel = Trim$(Nz(Me.txtExcelHeader.value, ""))
    strTarget = Trim$(Nz(Me.cboTargetField.value, ""))
    lngProfile = Nz(Me.cboProfile.value, 1) ' Получаем текущий профиль

    If Len(strExcel) = 0 Or Len(strTarget) = 0 Then
        mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetLoc("MSG_FILL_MAPPING_FIELDS"), vbExclamation
        Exit Sub
    End If

    ' --- PHASE 27: Smart Typing (Prompt user for data type) ---
    If Not mod_Schema_Manager.FieldExists("tbl_Personnel_Master", strTarget) Then
        strTypeSelection = InputBox( _
            mod_UI_Helpers.GetLoc("MSG_FIELD_MISSING") & " [" & strTarget & "]" & vbCrLf & vbCrLf & _
            mod_UI_Helpers.GetLoc("PROMPT_SELECT_DATA_TYPE"), _
            mod_UI_Helpers.GetLoc("TITLE_SCHEMA_MANAGER"), "1")

        If Trim$(strTypeSelection) = "" Then
            mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetLoc("MSG_MAPPING_CANCELED"), vbExclamation
            Exit Sub
        End If

        Select Case Trim$(strTypeSelection)
            Case "2": strSqlType = "DATETIME"
            Case "3": strSqlType = "LONG"
            Case "4": strSqlType = "LONGTEXT"
            Case "5": strSqlType = "BIT"
            Case Else: strSqlType = "VARCHAR(255)"
        End Select

        mod_Schema_Manager.AddNewFieldToSchema strTarget, strSqlType
        mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetLoc("MSG_FIELD_CREATED") & " " & strSqlType & ".", vbInformation
    End If

    ' --- PHASE 26 & 31: Parameterized Query for mapping (Safe Insert into selected Profile) ---
    Set db = CurrentDb
    strSQL = "PARAMETERS prmProfile Long, prmExcel Text (255), prmTarget Text (100); " & _
             "INSERT INTO tbl_Import_Mapping (ProfileID, ExcelHeader, TargetField) " & _
             "VALUES ([prmProfile], [prmExcel], [prmTarget]);"

    Set qdf = db.CreateQueryDef("", strSQL)
    qdf.Parameters("prmProfile").value = lngProfile
    qdf.Parameters("prmExcel").value = Left$(strExcel, 255)
    qdf.Parameters("prmTarget").value = Left$(strTarget, 100)

    qdf.Execute dbFailOnError

    Set qdf = Nothing
    Set db = Nothing

    RefreshMappingList
    Me.txtExcelHeader = ""
    Me.cboTargetField = ""
    Exit Sub
ErrorHandler:
    If Not qdf Is Nothing Then Set qdf = Nothing
    If Not db Is Nothing Then Set db = Nothing
    Debug.Print "btnAddMapping_Click: " & Err.Description
    mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetLoc("ERR_ADD_MAPPING_FAILED") & " " & Err.Description, vbExclamation
End Sub

' ========================================================
' NEW: Edit Header button
' ========================================================
Private Sub btnEditMapping_Click()
    On Error GoTo ErrorHandler
    Dim varId As Variant
    Dim strOldHeader As String
    Dim strNewHeader As String
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef

    If Me.lstMapping.ListIndex < 0 Then
        mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetLoc("MSG_SELECT_ROW_DELETE"), vbExclamation
        Exit Sub
    End If

    varId = Me.lstMapping.Column(0)
    strOldHeader = Nz(Me.lstMapping.Column(1), "")

    strNewHeader = InputBox(mod_UI_Helpers.GetLoc("PROMPT_EDIT_EXCEL_HEADER"), mod_UI_Helpers.GetLoc("TITLE_EDIT_MAPPING"), strOldHeader)
    If Trim$(strNewHeader) = "" Or strNewHeader = strOldHeader Then Exit Sub

    Set db = CurrentDb
    Set qdf = db.CreateQueryDef("", "PARAMETERS prmNew Text (255), prmId Long; UPDATE tbl_Import_Mapping SET ExcelHeader = [prmNew] WHERE MappingID = [prmId];")
    qdf.Parameters("prmNew").value = Left$(Trim$(strNewHeader), 255)
    qdf.Parameters("prmId").value = CLng(varId)
    qdf.Execute dbFailOnError
    Set qdf = Nothing
    Set db = Nothing

    RefreshMappingList
    mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetLoc("MSG_MAPPING_UPDATED"), vbInformation
    Exit Sub
ErrorHandler:
    mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetLoc("TITLE_ERROR") & ": " & Err.Description, vbCritical
End Sub

Private Sub btnDeleteMapping_Click()
    On Error GoTo ErrorHandler
    Dim varId As Variant
    Dim strTarget As String
    Dim db As DAO.Database

    If Me.lstMapping.ListIndex < 0 Then
        mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetLoc("MSG_SELECT_ROW_DELETE"), vbExclamation
        Exit Sub
    End If

    varId = Me.lstMapping.Column(0)
    strTarget = Nz(Me.lstMapping.Column(2), "")

    If UCase(Trim$(strTarget)) = "PERSONUID" Then
        If Not mod_UI_Helpers.AskUserYesNo(mod_UI_Helpers.GetLoc("PROMPT_DEL_PERSONUID"), mod_UI_Helpers.GetLoc("TITLE_CONFIRM")) Then Exit Sub
    End If

    Set db = CurrentDb
    db.Execute "DELETE FROM tbl_Import_Mapping WHERE MappingID = " & CLng(varId), dbFailOnError
    Set db = Nothing

    RefreshMappingList
    Exit Sub
ErrorHandler:
    mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetLoc("ERR_DELETE_FAILED") & " " & Err.Description, vbExclamation
End Sub

Private Sub btnReSeedMapping_Click()
    On Error GoTo ErrorHandler

    ' Защита от случайного сброса других профилей
    Dim lngProfile As Long
    lngProfile = Nz(Me.cboProfile.value, 1)
    If lngProfile <> 1 Then
        mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetLoc("MSG_RESEED_ONLY_MAIN"), vbExclamation
        Exit Sub
    End If

    If Not mod_UI_Helpers.AskUserYesNo(mod_UI_Helpers.GetLoc("PROMPT_RESEED_MAPPING"), mod_UI_Helpers.GetLoc("TITLE_CONFIRM")) Then Exit Sub

    mod_Schema_Manager.ReSeedMapping
    RefreshMappingList
    mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetLoc("MSG_MAPPING_RESTORED"), vbInformation
    Exit Sub
ErrorHandler:
    mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetLoc("ERR_RESEED_FAILED") & " " & Err.Description, vbExclamation
End Sub

' --- Maintenance tab (pgMaintenance) ---
Private Sub cmdCreateBackup_Click()
    On Error GoTo ErrorHandler
    Dim result As Object

    Set result = mod_Maintenance_Logic.CreateDatabaseBackupResult()
    mod_UI_Helpers.ShowMessage CStr(result("Message")), IIf(CBool(result("Success")), vbInformation, vbExclamation)
    Set result = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print "cmdCreateBackup_Click: " & Err.Description
End Sub

Private Sub cmdClearLogs_Click()
    On Error GoTo ErrorHandler
    Dim result As Object

    If Not mod_UI_Helpers.AskUserYesNo("Are you sure you want to clear the validation log?", "StaffState") Then Exit Sub

    Set result = mod_Maintenance_Logic.ClearValidationLogResult()
    mod_UI_Helpers.ShowMessage CStr(result("Message")), IIf(CBool(result("Success")), vbInformation, vbExclamation)
    Set result = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print "cmdClearLogs_Click: " & Err.Description
End Sub

Private Sub btnRunHealthCheck_Click()
    On Error GoTo ErrorHandler
    Dim result As Object

    Set result = mod_Maintenance_Logic.RunDataHealthCheckResult(False)
    mod_UI_Helpers.ShowMessage CStr(result("Message")), IIf(CBool(result("Success")), vbInformation, vbExclamation)
    Set result = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print "btnRunHealthCheck_Click: " & Err.Description
End Sub

' =============================================
' @description Double click to change Data Type of an existing field
' =============================================
Private Sub lstMapping_DblClick(Cancel As Integer)
    On Error GoTo ErrorHandler
    Dim strTarget As String
    Dim strTypeSelection As String
    Dim strSqlType As String

    If Me.lstMapping.ListIndex < 0 Then Exit Sub

    ' Get TargetField from the 3rd column (index 2)
    strTarget = Nz(Me.lstMapping.Column(2), "")
    If strTarget = "" Then Exit Sub

    If Not mod_UI_Helpers.AskUserYesNo(mod_UI_Helpers.GetLoc("PROMPT_CHANGE_TYPE_WARN") & " [" & strTarget & "]?" & vbCrLf & vbCrLf & mod_UI_Helpers.GetLoc("PROMPT_CHANGE_TYPE_WARN2"), mod_UI_Helpers.GetLoc("TITLE_CHANGE_TYPE")) Then Exit Sub

    strTypeSelection = InputBox( _
        mod_UI_Helpers.GetLoc("PROMPT_CHANGE_TYPE_INPUT") & " [" & strTarget & "]" & vbCrLf & vbCrLf & _
        mod_UI_Helpers.GetLoc("PROMPT_SELECT_DATA_TYPE"), _
        mod_UI_Helpers.GetLoc("TITLE_SCHEMA_MANAGER"), "1")

    If Trim$(strTypeSelection) = "" Then Exit Sub

    Select Case Trim$(strTypeSelection)
        Case "2": strSqlType = "DATETIME"
        Case "3": strSqlType = "LONG"
        Case "4": strSqlType = "LONGTEXT"
        Case "5": strSqlType = "BIT"
        Case Else: strSqlType = "VARCHAR(255)"
    End Select

    mod_Schema_Manager.AlterFieldType strTarget, strSqlType
    mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetLoc("MSG_FIELD_TYPE_CHANGED") & " " & strSqlType & ".", vbInformation

    RefreshMappingList
    Exit Sub
ErrorHandler:
    mod_UI_Helpers.ShowMessage mod_UI_Helpers.GetLoc("ERR_CHANGE_TYPE_FAILED") & " " & Err.Description, vbCritical
End Sub
