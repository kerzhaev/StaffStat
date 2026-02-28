Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================
' @form uf_Dashboard
' @author Kerzhaev Evgeniy
' @description Main control panel for StaffState
' @note 100% English version.
' =============================================

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    ' === PHASE 25: Ensure Back-End is connected before loading dashboard ===
    If Not mod_Table_Relinker.VerifyAndRelinkTables() Then
        DoCmd.Quit acQuitSaveNone
        Exit Sub
    End If
    ' =======================================================================

    Me.txtEndDate.value = Date
    Me.txtStartDate.value = Date - 30
    RefreshMetrics

    Exit Sub
ErrorHandler:
    MsgBox "UI error: " & Err.Description, vbCritical
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

    Me.lblTotalCount.Caption = "Total: " & CStr(d("TotalCount"))
    Me.lblActiveCount.Caption = "Active: " & CStr(d("ActiveCount"))
    lngErr = CLng(d("ErrorCount"))
    Me.lblErrorCount.Caption = "Errors: " & CStr(lngErr)

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
    Me.lblTotalCount.Caption = "Total: --"
    Me.lblActiveCount.Caption = "Active: --"
    Me.lblErrorCount.Caption = "Errors: --"
    Me.lblErrorCount.ForeColor = vbBlack
End Sub

Private Sub cmdSettings_Click()
    On Error GoTo ErrorHandler
    DoCmd.OpenForm "uf_Settings", acNormal
    Exit Sub
ErrorHandler:
    MsgBox "Unable to open Settings: " & Err.Description, vbCritical
End Sub

Private Sub btnImport_Click()
    On Error GoTo ErrorHandler

    Me.lblStatus.Caption = "Running import..."
    DoEvents

    If mod_Import_Logic.ImportExcelData(False) Then
        Me.lblStatus.Caption = "Import completed successfully."
    Else
        Me.lblStatus.Caption = "Import canceled or failed."
    End If

    Exit Sub
ErrorHandler:
    MsgBox "UI error: " & Err.Description, vbCritical
    Me.lblStatus.Caption = "Error."
End Sub

Private Sub cmdHealthCheck_Click()
    On Error GoTo ErrorHandler

    Dim lngErrors As Long
    Me.lblStatus.Caption = "Running health check..."
    DoEvents

    lngErrors = mod_Maintenance_Logic.RunDataHealthCheck(False)

    If lngErrors > 0 Then
        If mod_UI_Helpers.AskUserYesNo("Export error report to Excel?", "Health Check") Then
            mod_Maintenance_Logic.ExportValidationLogToExcel
        End If
    End If

    If lngErrors >= 0 Then
        Me.lblStatus.Caption = "Health check done. Errors: " & lngErrors
        RefreshMetrics
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Health check error: " & Err.Description, vbCritical
    Me.lblStatus.Caption = "Error."
End Sub

Private Sub btnAnalyze_Click()
    On Error GoTo ErrorHandler

    Dim iNew As Long
    Dim iUpd As Long

    Me.lblStatus.Caption = "Running analysis..."
    DoEvents

    mod_Analysis_Logic.SyncBufferToMaster iNew, iUpd, False

    Me.lblStatus.Caption = "Analysis completed. New: " & iNew & ", Updated: " & iUpd

    Exit Sub
ErrorHandler:
    MsgBox "UI error: " & Err.Description, vbCritical
    Me.lblStatus.Caption = "Error."
End Sub

Private Sub btnOpenLog_Click()
    On Error GoTo ErrorHandler
    Me.lblStatus.Caption = "Opening log..."
    DoCmd.OpenTable "tbl_History_Log"
    Me.lblStatus.Caption = "Log opened."
    Exit Sub
ErrorHandler:
    MsgBox "Unable to open log: " & Err.Description, vbCritical
End Sub

Private Sub btnCreateIndexes_Click()
    On Error GoTo ErrorHandler
    Dim iCreated As Long
    Dim iSkipped As Long

    Me.lblStatus.Caption = "Creating performance indexes..."
    DoEvents

    mod_App_Init.CreatePerformanceIndexes iCreated, iSkipped, False
    Me.lblStatus.Caption = "Indexes created. Created: " & iCreated & ", Skipped: " & iSkipped

    Exit Sub
ErrorHandler:
    MsgBox "Index creation error: " & Err.Description, vbCritical
    Me.lblStatus.Caption = "Error."
End Sub

' --- FULL UPDATE BUTTON FIX ---
Private Sub btnFullSync_Click()
    On Error GoTo ErrorHandler

    Me.lblStatus.Caption = "Running full update..."
    DoEvents

    mod_Analysis_Logic.RunFullSyncProcess

    Me.lblStatus.Caption = "Full update completed."
    RefreshMetrics
    Exit Sub

ErrorHandler:
    MsgBox "Full update error: " & Err.Description, vbCritical
    Me.lblStatus.Caption = "Error."
End Sub

Private Sub btnFindDuplicates_Click()
    On Error GoTo ErrorHandler

    Me.lblStatus.Caption = "Opening duplicates finder..."
    DoEvents

    If CurrentProject.AllForms("uf_Search").IsLoaded Then
        DoCmd.Close acForm, "uf_Search", acSaveNo
    End If
    DoCmd.OpenForm "uf_Search", acNormal, , , , , "MODE=DUPLICATES"

    Me.lblStatus.Caption = "Duplicates finder opened."
    Exit Sub

ErrorHandler:
    On Error Resume Next
    mod_App_Logger.LogError "uf_Dashboard.btnFindDuplicates_Click", CStr(Err.Number) & " - " & Err.Description, False
    MsgBox "Unable to open duplicates finder: " & Err.Description, vbCritical
    Me.lblStatus.Caption = "Error."
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
        MsgBox "Enter Start Date and End Date.", vbExclamation
        Exit Sub
    End If
    If Not IsDate(sStart) Or Not IsDate(sEnd) Then
        MsgBox "Invalid date format. Use DD.MM.YYYY or DD/MM/YYYY.", vbExclamation
        Exit Sub
    End If
    dtStart = CDate(sStart)
    dtEnd = CDate(sEnd)
    If dtStart > dtEnd Then
        MsgBox "Start date must be earlier than or equal to End date.", vbExclamation
        Exit Sub
    End If

    Me.lblStatus.Caption = "Generating audit report..."
    DoEvents
    mod_Reports_Logic.GenerateAuditReport dtStart, dtEnd
    Me.lblStatus.Caption = "Audit report done."
    Exit Sub
ErrorHandler:
    MsgBox "Report error: " & Err.Description, vbCritical
    Me.lblStatus.Caption = "Error."
End Sub

Private Sub btnSnapshotReport_Click()
    On Error GoTo ErrorHandler
    Me.lblStatus.Caption = "Generating snapshot report..."
    DoEvents
    mod_Reports_Logic.GenerateCurrentStaffReport
    Me.lblStatus.Caption = "Snapshot report done."
    Exit Sub
ErrorHandler:
    MsgBox "Snapshot report error: " & Err.Description, vbCritical
    Me.lblStatus.Caption = "Error."
End Sub
