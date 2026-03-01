Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================
' @form uf_Dashboard
' @author Kerzhaev Evgeniy
' @description Main control panel for StaffState
' @note 100% English version.
' =============================================
Option Compare Database
Option Explicit

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    ' === PHASE 25: Ensure Back-End is connected before loading dashboard ===
    If Not mod_Table_Relinker.VerifyAndRelinkTables() Then
        DoCmd.Quit acQuitSaveNone
        Exit Sub
    End If
    ' =======================================================================

    ' --- PHASE 30: Apply UI Localization ---
    Me.Caption = mod_UI_Helpers.GetLoc("TITLE_DASHBOARD")

    On Error Resume Next
    ' Перевод кнопок
    Me.cmdSettings.Caption = mod_UI_Helpers.GetLoc("BTN_SETTINGS")
    Me.btnImport.Caption = mod_UI_Helpers.GetLoc("BTN_IMPORT")
    Me.cmdHealthCheck.Caption = mod_UI_Helpers.GetLoc("BTN_HEALTH_CHECK")
    Me.btnAnalyze.Caption = mod_UI_Helpers.GetLoc("BTN_ANALYZE")
    Me.btnOpenLog.Caption = mod_UI_Helpers.GetLoc("BTN_OPEN_LOG")
    Me.btnCreateIndexes.Caption = mod_UI_Helpers.GetLoc("BTN_CREATE_INDEXES")
    Me.btnFullSync.Caption = mod_UI_Helpers.GetLoc("BTN_FULL_SYNC")
    Me.btnFindDuplicates.Caption = mod_UI_Helpers.GetLoc("BTN_FIND_DUPLICATES")
    Me.btnGenerateChangeReport.Caption = mod_UI_Helpers.GetLoc("BTN_CHANGE_REPORT")
    Me.btnSnapshotReport.Caption = mod_UI_Helpers.GetLoc("BTN_SNAPSHOT")

    ' Перевод рамок (Frames/Groups) - добавлено несколько вариантов имен для подстраховки
    Me.fraManualControls.Caption = mod_UI_Helpers.GetLoc("LBL_MANUAL_CONTROLS")
    Me.fraManualControls.Caption = mod_UI_Helpers.GetLoc("LBL_MANUAL_CONTROLS")
    On Error GoTo ErrorHandler

    Me.txtEndDate.value = Date
    Me.txtStartDate.value = Date - 30
    RefreshMetrics

    Exit Sub
ErrorHandler:
    MsgBox mod_UI_Helpers.GetLoc("ERR_UI") & Err.Description, vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrorHandler
    RefreshMetrics
    Exit Sub
ErrorHandler:
    Debug.Print "Form_Activate RefreshMetrics: " & Err.Description
End Sub

' =============================================
' @description Load metrics from GetDashboardStats
' =============================================
Public Sub RefreshMetrics()
    On Error GoTo ErrorHandler

    Dim d As Object
    Dim lngErr As Long

    Set d = mod_Maintenance_Logic.GetDashboardStats()
    If d Is Nothing Then GoTo SafeFallback

    Me.lblTotalCount.Caption = mod_UI_Helpers.GetLoc("LBL_TOTAL") & CStr(d("TotalCount"))
    Me.lblActiveCount.Caption = mod_UI_Helpers.GetLoc("LBL_ACTIVE") & CStr(d("ActiveCount"))
    lngErr = CLng(d("ErrorCount"))
    Me.lblErrorCount.Caption = mod_UI_Helpers.GetLoc("LBL_ERRORS") & CStr(lngErr)

    If lngErr > 0 Then
        Me.lblErrorCount.ForeColor = vbRed
    Else
        Me.lblErrorCount.ForeColor = RGB(0, 128, 0)
    End If

    Set d = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "RefreshMetrics error: " & Err.Number & " - " & Err.Description
    If Not d Is Nothing Then Set d = Nothing
    GoTo SafeFallback
SafeFallback:
    On Error Resume Next
    Me.lblTotalCount.Caption = mod_UI_Helpers.GetLoc("LBL_TOTAL") & "--"
    Me.lblActiveCount.Caption = mod_UI_Helpers.GetLoc("LBL_ACTIVE") & "--"
    Me.lblErrorCount.Caption = mod_UI_Helpers.GetLoc("LBL_ERRORS") & "--"
    Me.lblErrorCount.ForeColor = vbBlack
End Sub

Private Sub cmdSettings_Click()
    On Error GoTo ErrorHandler
    DoCmd.OpenForm "uf_Settings", acNormal
    Exit Sub
ErrorHandler:
    MsgBox mod_UI_Helpers.GetLoc("ERR_UI") & Err.Description, vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
End Sub

Private Sub btnImport_Click()
    On Error GoTo ErrorHandler

    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("MSG_RUNNING_IMPORT")
    DoEvents

    If mod_Import_Logic.ImportExcelData(False) Then
        Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("MSG_IMPORT_SUCCESS")
    Else
        Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("MSG_IMPORT_FAILED")
    End If

    Exit Sub
ErrorHandler:
    MsgBox mod_UI_Helpers.GetLoc("ERR_UI") & Err.Description, vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("STATUS_ERROR")
End Sub

Private Sub cmdHealthCheck_Click()
    On Error GoTo ErrorHandler

    Dim lngErrors As Long
    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("MSG_RUNNING_HEALTH")
    DoEvents

    lngErrors = mod_Maintenance_Logic.RunDataHealthCheck(False)

    If lngErrors > 0 Then
        If mod_UI_Helpers.AskUserYesNo(mod_UI_Helpers.GetLoc("PROMPT_EXPORT_ERRORS"), mod_UI_Helpers.GetLoc("TITLE_HEALTH_CHECK")) Then
            mod_Maintenance_Logic.ExportValidationLogToExcel
        End If
    End If

    If lngErrors >= 0 Then
        Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("MSG_HEALTH_DONE") & lngErrors
        RefreshMetrics
    End If
    Exit Sub

ErrorHandler:
    MsgBox mod_UI_Helpers.GetLoc("ERR_UI") & Err.Description, vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("STATUS_ERROR")
End Sub

Private Sub btnAnalyze_Click()
    On Error GoTo ErrorHandler

    Dim iNew As Long
    Dim iUpd As Long

    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("MSG_RUNNING_ANALYSIS")
    DoEvents

    mod_Analysis_Logic.SyncBufferToMaster iNew, iUpd, False

    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("MSG_ANALYSIS_DONE") & iNew & mod_UI_Helpers.GetLoc("LBL_UPDATED") & iUpd

    Exit Sub
ErrorHandler:
    MsgBox mod_UI_Helpers.GetLoc("ERR_UI") & Err.Description, vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("STATUS_ERROR")
End Sub

Private Sub btnOpenLog_Click()
    On Error GoTo ErrorHandler
    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("MSG_OPENING_LOG")
    DoCmd.OpenTable "tbl_History_Log"
    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("MSG_LOG_OPENED")
    Exit Sub
ErrorHandler:
    MsgBox mod_UI_Helpers.GetLoc("ERR_UI") & Err.Description, vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
End Sub

Private Sub btnCreateIndexes_Click()
    On Error GoTo ErrorHandler
    Dim iCreated As Long
    Dim iSkipped As Long

    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("MSG_CREATING_INDEXES")
    DoEvents

    mod_App_Init.CreatePerformanceIndexes iCreated, iSkipped, False
    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("MSG_INDEXES_DONE") & iCreated & mod_UI_Helpers.GetLoc("LBL_SKIPPED") & iSkipped

    Exit Sub
ErrorHandler:
    MsgBox mod_UI_Helpers.GetLoc("ERR_UI") & Err.Description, vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("STATUS_ERROR")
End Sub

' --- FULL UPDATE BUTTON FIX ---
Private Sub btnFullSync_Click()
    On Error GoTo ErrorHandler

    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("MSG_RUNNING_FULL_SYNC")
    DoEvents

    mod_Analysis_Logic.RunFullSyncProcess

    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("MSG_FULL_SYNC_DONE")
    RefreshMetrics
    Exit Sub

ErrorHandler:
    MsgBox mod_UI_Helpers.GetLoc("ERR_UI") & Err.Description, vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("STATUS_ERROR")
End Sub

Private Sub btnFindDuplicates_Click()
    On Error GoTo ErrorHandler

    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("MSG_OPENING_DUPLICATES")
    DoEvents

    If CurrentProject.AllForms("uf_Search").IsLoaded Then
        DoCmd.Close acForm, "uf_Search", acSaveNo
    End If
    DoCmd.OpenForm "uf_Search", acNormal, , , , , "MODE=DUPLICATES"

    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("MSG_DUPLICATES_OPENED")
    Exit Sub

ErrorHandler:
    On Error Resume Next
    mod_App_Logger.LogError "uf_Dashboard.btnFindDuplicates_Click", CStr(Err.Number) & " - " & Err.Description, False
    MsgBox mod_UI_Helpers.GetLoc("ERR_OPEN_DUPLICATES") & Err.Description, vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("STATUS_ERROR")
End Sub

Private Sub btnGenerateChangeReport_Click()
    On Error GoTo ErrorHandler

    Dim dtStart As Date
    Dim dtEnd As Date
    Dim sStart As String
    Dim sEnd As String

    sStart = Trim$(Replace(Nz(Me.txtStartDate.value, ""), ".", "/"))
    sEnd = Trim$(Replace(Nz(Me.txtEndDate.value, ""), ".", "/"))
    If sStart = "" Or sEnd = "" Then
        MsgBox mod_UI_Helpers.GetLoc("ERR_ENTER_DATES"), vbExclamation, mod_UI_Helpers.GetLoc("TITLE_WARNING")
        Exit Sub
    End If
    If Not IsDate(sStart) Or Not IsDate(sEnd) Then
        MsgBox mod_UI_Helpers.GetLoc("ERR_INVALID_DATE_FORMAT"), vbExclamation, mod_UI_Helpers.GetLoc("TITLE_WARNING")
        Exit Sub
    End If
    dtStart = CDate(sStart)
    dtEnd = CDate(sEnd)
    If dtStart > dtEnd Then
        MsgBox mod_UI_Helpers.GetLoc("ERR_START_AFTER_END"), vbExclamation, mod_UI_Helpers.GetLoc("TITLE_WARNING")
        Exit Sub
    End If

    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("MSG_GEN_AUDIT")
    DoEvents
    mod_Reports_Logic.GenerateAuditReport dtStart, dtEnd
    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("MSG_AUDIT_DONE")
    Exit Sub
ErrorHandler:
    MsgBox mod_UI_Helpers.GetLoc("ERR_UI") & Err.Description, vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("STATUS_ERROR")
End Sub

Private Sub btnSnapshotReport_Click()
    On Error GoTo ErrorHandler
    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("MSG_GEN_SNAPSHOT")
    DoEvents
    mod_Reports_Logic.GenerateCurrentStaffReport
    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("MSG_SNAPSHOT_DONE")
    Exit Sub
ErrorHandler:
    MsgBox mod_UI_Helpers.GetLoc("ERR_UI") & Err.Description, vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
    Me.lblStatus.Caption = mod_UI_Helpers.GetLoc("STATUS_ERROR")
End Sub
