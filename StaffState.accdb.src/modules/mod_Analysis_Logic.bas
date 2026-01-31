Attribute VB_Name = "mod_Analysis_Logic"
Option Explicit

' =============================================
' @module mod_Analysis_Logic (PRODUCTION)
' @description Universal data synchronization
' =============================================

Public Sub SyncBufferToMaster(ByRef outNew As Long, ByRef outUpdated As Long, Optional ByVal blnSuppressMsgBox As Boolean = False)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rsBuffer As DAO.Recordset
    Dim rsMaster As DAO.Recordset
    Dim fld As DAO.Field
    Dim strUID As String
    Dim iNew As Long, iUpd As Long
    Dim dtChangeDate As Date

    ' 1. Extend Master structure
    DoCmd.Close acTable, "tbl_Personnel_Master", acSaveYes
    DoCmd.Close acTable, "tbl_Import_Buffer", acSaveYes

    mod_Schema_Manager.SyncMasterStructure

    Set db = CurrentDb
    db.TableDefs.Refresh ' Must clear table cache

    ' 2. Get export file date from metadata (or Now() as fallback)
    dtChangeDate = GetExportFileDate()
    Debug.Print "Using ChangeDate: " & dtChangeDate

    Set rsBuffer = db.OpenRecordset("tbl_Import_Buffer", dbOpenSnapshot)
    Set rsMaster = db.OpenRecordset("tbl_Personnel_Master", dbOpenDynaset)

    outNew = 0
    outUpdated = 0

    ' If buffer is empty - exit
    If rsBuffer.EOF Then GoTo ExitHandler

    Do While Not rsBuffer.EOF
        strUID = Nz(rsBuffer!PersonUID_Raw, "")

        If strUID <> "" Then
            rsMaster.FindFirst "PersonUID = '" & strUID & "'"

            If rsMaster.NoMatch Then
                ' --- NEW EMPLOYEE ---
                rsMaster.AddNew
                rsMaster!PersonUID = strUID
                rsMaster!LastUpdated = dtChangeDate
                rsMaster!IsActive = True
                CopyAllFields rsBuffer, rsMaster
                rsMaster.Update

                ' NOTE: Avoid Cyrillic literals in DB writes for first-run stability.
                ' UI will translate these tokens to Russian.
                LogChange strUID, "_System", "", "Added", dtChangeDate
                iNew = iNew + 1
            Else
                ' --- EXISTING ---
                rsMaster.Edit

                Dim bChanged As Boolean
                bChanged = False

                For Each fld In rsBuffer.Fields
                    Dim sBufName As String
                    Dim sMasName As String

                    sBufName = fld.Name
                    sMasName = ""

                    ' 1. Name logic (remove _Raw suffix)
                    If Right(sBufName, 4) = "_Raw" Then
                        Select Case sBufName
                            Case "SourceID_Raw": sMasName = "SourceID"
                            Case "PersonUID_Raw": sMasName = "PersonUID"
                            Case "Rank_Raw": sMasName = "RankName"
                            Case "FullName_Raw": sMasName = "FullName"
                            Case "BirthDate_Raw": sMasName = "BirthDate"
                            Case "WorkStatus_Raw": sMasName = "WorkStatus"
                            Case "PosCode_Raw": sMasName = "PosCode"
                            Case "PosName_Raw": sMasName = "PosName"
                            Case "OrderDate_Raw": sMasName = "OrderDate"
                            Case "OrderNum_Raw": sMasName = "OrderNum"
                            Case Else: sMasName = Left(sBufName, Len(sBufName) - 4)
                        End Select
                    Else
                        sMasName = sBufName ' Dynamic fields
                    End If

                    ' 2. Check and update
                    If sMasName <> "" Then
                        If FieldExistsInRS(rsMaster, sMasName) Then
                            ' Exclude technical fields from change tracking
                            If sMasName <> "PersonUID" And sMasName <> "LogID" And sMasName <> "LastUpdated" _
                               And sMasName <> "ID" And sMasName <> "IsActive" Then
                                Dim valBuf As String, valMas As String
                                valBuf = Nz(fld.Value, "")
                                valMas = Nz(rsMaster.Fields(sMasName).Value, "")

                                If valBuf <> valMas Then
                                    rsMaster.Fields(sMasName).Value = fld.Value
                                    LogChange strUID, sMasName, valMas, valBuf, dtChangeDate
                                    bChanged = True
                                End If
                            End If
                        End If
                    End If
                Next fld

                rsMaster!LastUpdated = dtChangeDate
                rsMaster.Update
                If bChanged Then iUpd = iUpd + 1
            End If
        End If

        rsBuffer.MoveNext
    Loop

ExitHandler:
    outNew = iNew
    outUpdated = iUpd

    If Not blnSuppressMsgBox Then
        MsgBox "Synchronization completed!" & vbCrLf & _
               "New: " & iNew & vbCrLf & _
               "Updated: " & iUpd, vbInformation
    End If

    rsBuffer.Close
    rsMaster.Close
    Set rsBuffer = Nothing
    Set rsMaster = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    If Not blnSuppressMsgBox Then
        MsgBox "Analysis error: " & Err.Description, vbCritical
    Else
        Debug.Print "Analysis error: " & Err.Description & " (" & Err.Number & ")"
    End If
    Resume ExitHandler
End Sub

' =============================================
' @description Runs the full import -> sync -> index pipeline.
'              Shows a single final summary message.
' =============================================
Public Sub RunFullSyncProcess()
    On Error GoTo ErrorHandler

    Dim blnImported As Boolean
    Dim iNew As Long
    Dim iUpd As Long
    Dim iIdxCreated As Long
    Dim iIdxSkipped As Long
    Dim strSummary As String

    blnImported = mod_Import_Logic.ImportExcelData(True)

    If blnImported Then
        SyncBufferToMaster iNew, iUpd, True
        mod_App_Init.CreatePerformanceIndexes iIdxCreated, iIdxSkipped, True
        strSummary = "Full Update Summary" & vbCrLf & _
                     "Import: OK" & vbCrLf & _
                     "Sync: New=" & iNew & ", Updated=" & iUpd & vbCrLf & _
                     "Indexes: Created=" & iIdxCreated & ", Skipped=" & iIdxSkipped
    Else
        strSummary = "Full Update Summary" & vbCrLf & _
                     "Import: FAILED or CANCELED" & vbCrLf & _
                     "Sync: SKIPPED" & vbCrLf & _
                     "Indexes: SKIPPED"
    End If

    MsgBox strSummary, vbInformation, "Full Update"
    Exit Sub

ErrorHandler:
    MsgBox "Full Update failed: " & Err.Description, vbCritical, "Full Update"
End Sub

' --- HELPER FUNCTIONS REMAIN THE SAME ---
' (CopyAllFields, FieldExistsInRS, LogChange)
' Make sure CopyAllFields is the latest version (where name logic matches the loop above)
' If needed - I can duplicate them here, but they haven't changed.

' =============================================
' @description Universally copies fields from rsSource to rsDest
'              Takes into account that new fields may have the same names.
' =============================================
Private Sub CopyAllFields(rsSource As DAO.Recordset, rsDest As DAO.Recordset)
    Dim fldSource As DAO.Field
    Dim strDestFieldName As String

    For Each fldSource In rsSource.Fields
        strDestFieldName = ""

        ' --- LOGIC FOR DETERMINING FIELD NAME IN DESTINATION (Master) ---
        ' 1. If field ends with _Raw, find its counterpart without _Raw
        If Right(fldSource.Name, 4) = "_Raw" Then
            Select Case fldSource.Name
                Case "SourceID_Raw": strDestFieldName = "SourceID"
                Case "PersonUID_Raw": strDestFieldName = "PersonUID"
                Case "Rank_Raw": strDestFieldName = "RankName"
                Case "FullName_Raw": strDestFieldName = "FullName"
                Case "BirthDate_Raw": strDestFieldName = "BirthDate"
                Case "WorkStatus_Raw": strDestFieldName = "WorkStatus"
                Case "PosCode_Raw": strDestFieldName = "PosCode"
                Case "PosName_Raw": strDestFieldName = "PosName"
                Case "OrderDate_Raw": strDestFieldName = "OrderDate"
                Case "OrderNum_Raw": strDestFieldName = "OrderNum"
                Case Else: strDestFieldName = Left(fldSource.Name, Len(fldSource.Name) - 4) ' General case
            End Select
        Else
            ' 2. If field doesn't end with _Raw (e.g., "??????_?????"),
            '    look for it in Master with the same name.
            strDestFieldName = fldSource.Name
        End If

        ' --- COPY VALUE ---
        If strDestFieldName <> "" And FieldExistsInRS(rsDest, strDestFieldName) Then
            ' Exclude technical fields (PersonUID already set, ID is auto-increment)
            If strDestFieldName <> "PersonUID" And strDestFieldName <> "ID" _
               And strDestFieldName <> "LogID" And strDestFieldName <> "LastUpdated" Then
                On Error Resume Next ' Ignore type errors (e.g., text to date)
                rsDest.Fields(strDestFieldName).Value = fldSource.Value
                On Error GoTo 0
            End If
        End If
    Next fldSource
End Sub

Private Function FieldExistsInRS(rs As DAO.Recordset, strName As String) As Boolean
    On Error Resume Next
    Dim x As Variant
    x = rs.Fields(strName).Name
    FieldExistsInRS = (Err.Number = 0)
End Function

Private Sub LogChange(strUID As String, strField As String, strOld As String, strNew As String, dtDate As Date)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbl_History_Log", dbOpenDynaset, dbAppendOnly)

    rs.AddNew
    rs!PersonUID = strUID
    rs!FieldName = strField
    rs!OldValue = Left(strOld, 255)
    rs!NewValue = Left(strNew, 255)
    rs!ChangeDate = dtDate
    rs.Update

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "LogChange error: " & Err.Description & " (" & Err.Number & ")"
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not db Is Nothing Then Set db = Nothing
End Sub

' =============================================
' @description Gets export file date from import metadata table.
' @return [Date] ExportFileDate from tbl_Import_Meta, or Now() if not available.
' =============================================
Private Function GetExportFileDate() As Date
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb

    ' Check if metadata table exists
    If Not TableExists("tbl_Import_Meta") Then
        GetExportFileDate = Now()
        Exit Function
    End If

    Set rs = db.OpenRecordset("SELECT TOP 1 ExportFileDate FROM tbl_Import_Meta;", dbOpenSnapshot)

    If Not rs.EOF Then
        If Not IsNull(rs!ExportFileDate) Then
            GetExportFileDate = rs!ExportFileDate
        Else
            GetExportFileDate = Now()
        End If
    Else
        ' No metadata record - use current date
        GetExportFileDate = Now()
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    ' Fallback to current date on any error
    GetExportFileDate = Now()
End Function

' =============================================
' @description Helper function to check if table exists.
' =============================================
Private Function TableExists(strTableName As String) As Boolean
    Dim tdf As DAO.TableDef
    On Error Resume Next
    Set tdf = CurrentDb.TableDefs(strTableName)
    TableExists = (Err.Number = 0)
    Err.Clear
End Function
