Attribute VB_Name = "mod_Analysis_Logic"
Option Compare Database
Option Explicit

' =============================================
' @description Universal data synchronization with Batch Transactions
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
    Dim lTotal As Long
    Dim lRecNum As Long
    Dim strOrderDateContext As String

    ' --- PHASE 29: Batch Transaction Variables ---
    Dim lngTransCount As Long
    Const c_BatchSize As Long = 2000 ' Commit every 2000 records

    ' Increase system lock limit to prevent RAM crashes on massive imports
    DAO.DBEngine.SetOption dbMaxLocksPerFile, 200000

    DoCmd.Close acTable, "tbl_Personnel_Master", acSaveYes
    DoCmd.Close acTable, "tbl_Import_Buffer", acSaveYes

    mod_Schema_Manager.SyncMasterStructure

    Set db = CurrentDb
    db.TableDefs.Refresh

    mod_Schema_Manager.CreateValidationLogTable

    dtChangeDate = GetExportFileDate()
    Debug.Print "Using ChangeDate: " & dtChangeDate

    lTotal = GetBufferRecordCount(db)
    If lTotal = 0 Then
        outNew = 0
        outUpdated = 0
        If Not blnSuppressMsgBox Then MsgBox "Synchronization completed! Buffer is empty.", vbInformation
        Exit Sub
    End If

    Set rsBuffer = db.OpenRecordset(GetOrderedBufferSQL(), dbOpenSnapshot)
    Set rsMaster = db.OpenRecordset("tbl_Personnel_Master", dbOpenDynaset)

    outNew = 0
    outUpdated = 0
    lRecNum = 0
    lngTransCount = 0

    ' Start the first transaction
    DBEngine.Workspaces(0).BeginTrans

    Do While Not rsBuffer.EOF
        lRecNum = lRecNum + 1
        strUID = Trim$(Nz(rsBuffer!PersonUID, ""))

        If strUID = "" Then
            ' skip empty UID
        ElseIf Not mod_Maintenance_Logic.IsValidPersonUID(strUID) Then
            mod_Maintenance_Logic.LogValidationError 0, "tbl_Import_Buffer", "InvalidPersonUID", "Invalid PersonUID format: " & strUID
        Else
            strOrderDateContext = GetBufferOrderDateContext(rsBuffer)

            rsMaster.FindFirst "PersonUID = '" & Replace(strUID, "'", "''") & "'"

            If rsMaster.NoMatch Then
                ' --- NEW EMPLOYEE ---
                rsMaster.AddNew
                rsMaster!PersonUID = strUID
                rsMaster!LastUpdated = dtChangeDate
                rsMaster!IsActive = True
                CopyAllFields rsBuffer, rsMaster
                rsMaster.Update

                LogChange strUID, "_System", "", "Added", dtChangeDate, ""
                iNew = iNew + 1
            Else
                ' --- EXISTING ---
                Dim bChanged As Boolean
                bChanged = False
                rsMaster.Edit

                For Each fld In rsBuffer.Fields
                    Dim sBufName As String
                    Dim sMasName As String

                    sBufName = fld.Name
                    sMasName = BufferFieldToMasterName(sBufName)

                    If sMasName <> "" And FieldExistsInRS(rsMaster, sMasName) Then
                        If sMasName <> "PersonUID" And sMasName <> "LogID" And sMasName <> "LastUpdated" _
                           And sMasName <> "ID" And sMasName <> "IsActive" Then

                            Dim vBuf As Variant, vMas As Variant
                            vBuf = fld.value
                            vMas = rsMaster.Fields(sMasName).value

                            If Not IsNull(rsBuffer.Fields(fld.Name).value) And rsBuffer.Fields(fld.Name).value <> "" Then
                                If Not ValuesEqual(vBuf, vMas) Then
                                    rsMaster.Fields(sMasName).value = fld.value
                                    LogChange strUID, sMasName, ValueToLogString(vMas), ValueToLogString(vBuf), dtChangeDate, strOrderDateContext
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

        ' --- BATCH TRANSACTION & UI UNFREEZE ---
        lngTransCount = lngTransCount + 1
        If lngTransCount >= c_BatchSize Then
            DBEngine.Workspaces(0).CommitTrans  ' Save the batch
            DoEvents                            ' Let Windows breathe and UI update
            DBEngine.Workspaces(0).BeginTrans   ' Start new batch
            lngTransCount = 0

            ' Optional debug to see progress in Immediate Window
            Debug.Print "Processed " & lRecNum & " of " & lTotal & " records..."
        End If

        rsBuffer.MoveNext
    Loop

    ' Commit the final remaining batch
    DBEngine.Workspaces(0).CommitTrans

ExitHandler:
    outNew = iNew
    outUpdated = iUpd

    If Not blnSuppressMsgBox Then
        MsgBox "Synchronization completed successfully!" & vbCrLf & _
               "New: " & iNew & vbCrLf & _
               "Updated: " & iUpd, vbInformation, "StaffState Import"
    End If

    On Error Resume Next
    rsBuffer.Close
    rsMaster.Close
    Set rsBuffer = Nothing
    Set rsMaster = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    ' Rollback only the current uncommitted batch (max 2000 records)
    DBEngine.Workspaces(0).Rollback
    If Not blnSuppressMsgBox Then
        MsgBox "Analysis error at record " & lRecNum & ": " & Err.Description, vbCritical
    Else
        Debug.Print "Analysis error: " & Err.Description & " (" & Err.Number & ")"
    End If
    Resume ExitHandler
End Sub

' =============================================
' @description Runs the full import -> sync -> index pipeline.
' =============================================
Public Sub RunFullSyncProcess()
    On Error GoTo ErrorHandler

    Dim blnImported As Boolean
    Dim iNew As Long
    Dim iUpd As Long
    Dim iIdxCreated As Long
    Dim iIdxSkipped As Long
    Dim strSummary As String

    ' 1. Import
    blnImported = mod_Import_Logic.ImportExcelData(True)

    If blnImported Then
        ' 2. Sync
        SyncBufferToMaster iNew, iUpd, True

        ' 3. Rebuild Indexes
        mod_App_Init.CreatePerformanceIndexes iIdxCreated, iIdxSkipped, True

        strSummary = "Full Update Summary" & vbCrLf & _
                     "Import: OK" & vbCrLf & _
                     "Sync: New=" & iNew & ", Updated=" & iUpd & vbCrLf & _
                     "Indexes: Created=" & iIdxCreated & ", Skipped=" & iIdxSkipped

        ' 4. Auto Health Check
        If UCase(Trim$(CStr(Nz(mod_Maintenance_Logic.GetSetting("AutoCheckEnabled", "False"), "False")))) = "TRUE" Then
            mod_Maintenance_Logic.RunDataHealthCheck True
        End If
    Else
        strSummary = "Full Update Summary" & vbCrLf & _
                     "Import: FAILED or CANCELED" & vbCrLf & _
                     "Sync: SKIPPED" & vbCrLf & _
                     "Indexes: SKIPPED"
    End If

    mod_UI_Helpers.ShowMessage strSummary, vbInformation
    Exit Sub

ErrorHandler:
    mod_UI_Helpers.ShowMessage "Full Update failed: " & Err.Description, vbCritical
End Sub

Private Sub CopyAllFields(rsSource As DAO.Recordset, rsDest As DAO.Recordset)
    Dim fldSource As DAO.Field
    Dim strDestFieldName As String

    For Each fldSource In rsSource.Fields
        strDestFieldName = fldSource.Name
        If strDestFieldName = "ID" Then strDestFieldName = ""

        If strDestFieldName <> "" And FieldExistsInRS(rsDest, strDestFieldName) Then
            If strDestFieldName <> "PersonUID" And strDestFieldName <> "LogID" And strDestFieldName <> "LastUpdated" Then
                On Error Resume Next
                rsDest.Fields(strDestFieldName).value = fldSource.value
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

Private Sub LogChange(strUID As String, strField As String, strOld As String, strNew As String, dtDate As Date, Optional strOrderDateContext As String = "")
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strNewVal As String

    strNewVal = Left(strNew, 255)
    If strOrderDateContext <> "" Then strNewVal = Left(strNew, 200) & " [OrderDate: " & Left(strOrderDateContext, 40) & "]"

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbl_History_Log", dbOpenDynaset, dbAppendOnly)

    rs.AddNew
    rs!PersonUID = strUID
    rs!fieldName = strField
    rs!OldValue = Left(strOld, 255)
    rs!NewValue = strNewVal
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

Private Function GetExportFileDate() As Date
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
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
        GetExportFileDate = Now()
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    GetExportFileDate = Now()
End Function

Private Function TableExists(strTableName As String) As Boolean
    Dim tdf As DAO.TableDef
    On Error Resume Next
    Set tdf = CurrentDb.TableDefs(strTableName)
    TableExists = (Err.Number = 0)
    Err.Clear
End Function

Private Function GetOrderedBufferSQL() As String
    GetOrderedBufferSQL = "SELECT * FROM tbl_Import_Buffer ORDER BY PersonUID ASC, Nz([OrderDate_Text],'') ASC, [ID] ASC;"
End Function

Private Function GetBufferRecordCount(db As DAO.Database) As Long
    On Error Resume Next
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS Cnt FROM tbl_Import_Buffer;", dbOpenSnapshot)
    If Not rs Is Nothing And Not rs.EOF Then
        GetBufferRecordCount = Nz(rs!Cnt, 0)
        rs.Close
    Else
        GetBufferRecordCount = 0
    End If
    Set rs = Nothing
End Function

Private Function GetBufferOrderDateContext(rsBuffer As DAO.Recordset) As String
    On Error Resume Next
    If FieldExistsInRS(rsBuffer, "OrderDate_Text") Then
        GetBufferOrderDateContext = Trim$(Nz(rsBuffer!OrderDate_Text, ""))
    Else
        GetBufferOrderDateContext = ""
    End If
End Function

Private Function BufferFieldToMasterName(sBufName As String) As String
    If sBufName = "ID" Then BufferFieldToMasterName = "": Exit Function
    BufferFieldToMasterName = sBufName
End Function

Private Function ValuesEqual(v1 As Variant, v2 As Variant) As Boolean
    Dim s1 As String, s2 As String
    If IsNull(v1) And IsNull(v2) Then ValuesEqual = True: Exit Function
    If IsNull(v1) Then
        s2 = ValueToLogString(v2)
        ValuesEqual = (s2 = "")
        Exit Function
    End If
    If IsNull(v2) Then
        s1 = ValueToLogString(v1)
        ValuesEqual = (s1 = "")
        Exit Function
    End If
    s1 = ValueToLogString(v1)
    s2 = ValueToLogString(v2)
    ValuesEqual = (s1 = s2)
End Function

Private Function ValueToLogString(v As Variant) As String
    If IsNull(v) Then ValueToLogString = "": Exit Function
    ValueToLogString = Trim$(CStr(v))
End Function

Private Function IsBufferValueEmpty(v As Variant) As Boolean
    If IsNull(v) Then IsBufferValueEmpty = True: Exit Function
    IsBufferValueEmpty = (Trim$(CStr(v)) = "")
End Function
