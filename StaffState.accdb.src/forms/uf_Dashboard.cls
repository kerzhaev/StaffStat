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
' =============================================

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    Me.txtEndDate.Value = Date
    Me.txtStartDate.Value = Date - 30

    Exit Sub
ErrorHandler:
    MsgBox "UI error: " & Err.Description, vbCritical
End Sub

' --- BUTTON 1: IMPORT ---
Private Sub btnImport_Click()
    On Error GoTo ErrorHandler

    Me.lblStatus.Caption = "Running import..."
    DoEvents ' Allow UI to refresh

    ' Call import logic
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

' --- BUTTON 2: ANALYZE ---
Private Sub btnAnalyze_Click()
    On Error GoTo ErrorHandler

    Dim iNew As Long
    Dim iUpd As Long

    Me.lblStatus.Caption = "Running analysis..."
    DoEvents

    ' Call analysis logic
    mod_Analysis_Logic.SyncBufferToMaster iNew, iUpd, False

    Me.lblStatus.Caption = "Analysis completed. New: " & iNew & ", Updated: " & iUpd

    Exit Sub
ErrorHandler:
    MsgBox "UI error: " & Err.Description, vbCritical
    Me.lblStatus.Caption = "Error."
End Sub

' --- BUTTON 3: LOG ---
Private Sub btnOpenLog_Click()
    On Error GoTo ErrorHandler

    Me.lblStatus.Caption = "Opening log..."

    ' Open history table
    DoCmd.OpenTable "tbl_History_Log"

    Me.lblStatus.Caption = "Log opened."

    Exit Sub
ErrorHandler:
    MsgBox "Unable to open log: " & Err.Description, vbCritical
End Sub

' --- BUTTON 4: CREATE INDEXES (Phase 7) ---
Private Sub btnCreateIndexes_Click()
    On Error GoTo ErrorHandler

    Dim iCreated As Long
    Dim iSkipped As Long

    Me.lblStatus.Caption = "Creating performance indexes..."
    DoEvents

    ' Call index creation routine
    mod_App_Init.CreatePerformanceIndexes iCreated, iSkipped, False

    Me.lblStatus.Caption = "Indexes created. Created: " & iCreated & ", Skipped: " & iSkipped

    Exit Sub
ErrorHandler:
    MsgBox "Index creation error: " & Err.Description, vbCritical
    Me.lblStatus.Caption = "Error."
End Sub

' --- BUTTON 5: FULL UPDATE ---
Private Sub btnFullSync_Click()
    On Error GoTo ErrorHandler

    Me.lblStatus.Caption = "Running full update..."
    DoEvents

    mod_Analysis_Logic.RunFullSyncProcess

    Me.lblStatus.Caption = "Full update completed."
    Exit Sub

ErrorHandler:
    MsgBox "Full update error: " & Err.Description, vbCritical
    Me.lblStatus.Caption = "Error."
End Sub

' --- BUTTON 6: CHANGES REPORT ---
Private Sub btnGenerateChangeReport_Click()
    On Error GoTo ErrorHandler

    Dim dtStart As Date
    Dim dtEnd As Date
    Dim sStart As String
    Dim sEnd As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim lCount As Long

    sStart = NormalizeDateInput(Me.txtStartDate.Value)
    sEnd = NormalizeDateInput(Me.txtEndDate.Value)

    If sStart = "" Or sEnd = "" Then
        MsgBox "Please enter start and end dates.", vbExclamation
        Exit Sub
    End If

    If Not IsDate(sStart) Or Not IsDate(sEnd) Then
        MsgBox "Please enter valid dates.", vbExclamation
        Exit Sub
    End If

    dtStart = CDate(sStart)
    dtEnd = CDate(sEnd)

    Me.txtStartDate.Value = sStart
    Me.txtEndDate.Value = sEnd

    If dtStart > dtEnd Then
        MsgBox "Start date must be earlier than or equal to end date.", vbExclamation
        Exit Sub
    End If

    Me.lblStatus.Caption = "Generating report..."
    DoEvents

    strSQL = "SELECT COUNT(*) AS C " & _
             "FROM tbl_History_Log " & _
             "WHERE ChangeDate BETWEEN #" & Format(dtStart, "mm\/dd\/yyyy hh:nn:ss") & "# " & _
             "AND #" & Format(dtEnd, "mm\/dd\/yyyy hh:nn:ss") & "#;"

    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If Not rs.EOF Then
        lCount = Nz(rs!C, 0)
    End If
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    If lCount = 0 Then
        Me.lblStatus.Caption = "No changes found."
        MsgBox "No changes found for selected period.", vbInformation
        Exit Sub
    End If

    If mod_Export_Logic.ExportChangeReport(dtStart, dtEnd) Then
        Me.lblStatus.Caption = "Report generated."
        MsgBox "Report generated for " & lCount & " changes.", vbInformation
    Else
        Me.lblStatus.Caption = "Report failed."
    End If

    Exit Sub
ErrorHandler:
    MsgBox "Report error: " & Err.Description, vbCritical
    Me.lblStatus.Caption = "Error."
End Sub

Private Function NormalizeDateInput(ByVal vValue As Variant) As String
    On Error GoTo ErrorHandler
    Dim s As String
    s = Trim$(Nz(vValue, ""))
    If s = "" Then
        NormalizeDateInput = ""
        Exit Function
    End If
    NormalizeDateInput = Replace(s, ".", "/")
    Exit Function
ErrorHandler:
    NormalizeDateInput = ""
End Function
