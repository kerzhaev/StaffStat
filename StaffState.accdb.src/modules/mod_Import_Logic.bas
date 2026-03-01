Attribute VB_Name = "mod_Import_Logic"
Option Compare Database
Option Explicit

' =============================================
' @module mod_Import_Logic (DYNAMIC v3.2 - Interactive Wizard)
' @description Dynamic auto-detect import + Interactive mapping wizard
' @note 100% English version. Safe for modern IDEs.
' =============================================

Private Const cstrLinkedTableName As String = "tmp_Excel_Link"

Public Function ImportExcelData(Optional ByVal blnSuppressMsgBox As Boolean = False) As Boolean
    On Error GoTo ErrorHandler

    ImportExcelData = False

    Dim strFilePath As String
    Dim strSkipped As String
    Dim strFinalMsg As String

    strFilePath = SelectExcelFile()
    If strFilePath = "" Then GoTo ExitHandler

    ' --- FIX: FORCE CLOSE TABLE BEFORE OPERATION ---
    DoCmd.Close acTable, "tbl_Import_Buffer", acSaveYes
    DoCmd.Close acTable, "tbl_Personnel_Master", acSaveYes
    DoEvents

    ' 1. Clear buffer
    CurrentDb.Execute "DELETE FROM tbl_Import_Buffer;", dbFailOnError

    ' 2. Link
    If Not LinkExcelFile(strFilePath, blnSuppressMsgBox) Then GoTo ExitHandler

    ' 3. Dynamic import (with Auto-Detect and Interactive Wizard)
    If Not RunDynamicImport(strSkipped, blnSuppressMsgBox) Then GoTo ExitHandler

    ' 4. Save import metadata
    UpdateImportMetadata strFilePath

    ImportExcelData = True

    If Not blnSuppressMsgBox Then
        strFinalMsg = mod_UI_Helpers.GetLoc("MSG_IMPORT_SUCCESS")

        If Len(strSkipped) > 0 Then
            strFinalMsg = strFinalMsg & vbCrLf & vbCrLf & _
                          mod_UI_Helpers.GetLoc("MSG_SKIPPED_COLS") & vbCrLf & _
                          strSkipped
        End If

        MsgBox strFinalMsg, vbInformation, mod_UI_Helpers.GetLoc("TITLE_INFO")
    End If

ExitHandler:
    DeleteExcelLink
    Exit Function

ErrorHandler:
    If Not blnSuppressMsgBox Then
        MsgBox mod_UI_Helpers.GetLoc("TITLE_ERROR") & " " & Err.Description, vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
    Else
        Debug.Print "Import error: " & Err.Description & " (" & Err.Number & ")"
    End If
    Resume ExitHandler
End Function

Private Function NormalizeString(ByVal s As String) As String
    NormalizeString = UCase(Trim$(s))
End Function

Private Function DetectBestProfile(tdfLink As DAO.TableDef) As Long
    Dim db As DAO.Database
    Dim rsProfiles As DAO.Recordset
    Dim rsMapping As DAO.Recordset
    Dim fld As DAO.Field
    Dim currentProfile As Long
    Dim matchCount As Long
    Dim maxMatches As Long
    Dim bestProfile As Long
    Dim hasUID As Boolean
    Dim normExcel As String
    Dim normHeader As String
    Dim targetField As String

    Set db = CurrentDb
    Set rsProfiles = db.OpenRecordset("SELECT DISTINCT ProfileID FROM tbl_Import_Profiles", dbOpenSnapshot)

    bestProfile = 0
    maxMatches = 0

    Do While Not rsProfiles.EOF
        currentProfile = rsProfiles!ProfileID
        matchCount = 0
        hasUID = False

        Set rsMapping = db.OpenRecordset("SELECT ExcelHeader, TargetField FROM tbl_Import_Mapping WHERE ProfileID = " & currentProfile, dbOpenSnapshot)

        Do While Not rsMapping.EOF
            normHeader = NormalizeString(Nz(rsMapping!ExcelHeader, ""))
            targetField = NormalizeString(Nz(rsMapping!targetField, ""))

            For Each fld In tdfLink.Fields
                normExcel = NormalizeString(fld.Name)
                If normExcel = normHeader Then
                    matchCount = matchCount + 1
                    If targetField = "PERSONUID" Then hasUID = True
                    Exit For
                End If
            Next fld
            rsMapping.MoveNext
        Loop
        rsMapping.Close

        If hasUID And matchCount > maxMatches Then
            maxMatches = matchCount
            bestProfile = currentProfile
        End If

        rsProfiles.MoveNext
    Loop
    rsProfiles.Close

    Set rsProfiles = Nothing
    Set rsMapping = Nothing
    Set db = Nothing

    DetectBestProfile = bestProfile
End Function

Private Function RunDynamicImport(ByRef outSkippedColumns As String, Optional ByVal blnSuppressMsgBox As Boolean = False) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim tdfLink As DAO.TableDef
    Dim fld As DAO.Field
    Dim strExcelField As String
    Dim strAccessField As String
    Dim strSelectPart As String
    Dim strInsertPart As String
    Dim strPersonUIDExcelName As String
    Dim strAllColumns As String
    Dim colUsedFields As Collection
    Dim lngBestProfile As Long

    outSkippedColumns = ""
    strPersonUIDExcelName = ""
    strAllColumns = ""
    Set colUsedFields = New Collection
    Set db = CurrentDb

    If Not mod_Schema_Manager.TableExists("tbl_Import_Mapping") Then
        If Not blnSuppressMsgBox Then MsgBox "tbl_Import_Mapping table is missing.", vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
        RunDynamicImport = False
        Exit Function
    End If

    Set tdfLink = db.TableDefs(cstrLinkedTableName)

    lngBestProfile = DetectBestProfile(tdfLink)

    If lngBestProfile = 0 Then
        For Each fld In tdfLink.Fields
            strAllColumns = strAllColumns & "[" & fld.Name & "] "
        Next fld

        If Not blnSuppressMsgBox Then
            MsgBox "Import failed: Could not auto-detect a matching mapping profile." & vbCrLf & _
                   "Ensure mapping exists for PersonUID." & vbCrLf & vbCrLf & _
                   "Columns found: " & Left(strAllColumns, 400), vbCritical, "Auto-Detect Failed"
        End If
        RunDynamicImport = False
        Exit Function
    End If

    strAllColumns = ""

    ' --- LOOP THROUGH ALL EXCEL COLUMNS ---
    For Each fld In tdfLink.Fields
        strExcelField = fld.Name

        If Len(strAllColumns) > 0 Then strAllColumns = strAllColumns & ", "
        strAllColumns = strAllColumns & "[" & strExcelField & "]"

        strAccessField = GetMappedFieldFromTable(strExcelField, lngBestProfile)

        ' ========================================================
        ' INTERACTIVE MAPPING WIZARD
        ' ========================================================
        If Len(strAccessField) = 0 Then
            If Not blnSuppressMsgBox Then
                ' Спрашиваем пользователя: нашли новую колонку, добавить?
                If mod_UI_Helpers.AskUserYesNo(mod_UI_Helpers.GetLoc("PROMPT_MAP_NEW_COL") & " [" & strExcelField & "]" & vbCrLf & vbCrLf & mod_UI_Helpers.GetLoc("PROMPT_MAP_NEW_COL2"), mod_UI_Helpers.GetLoc("TITLE_NEW_COL")) Then

                    Dim strNewTarget As String
                    strNewTarget = InputBox(mod_UI_Helpers.GetLoc("PROMPT_ENTER_EN_NAME"), mod_UI_Helpers.GetLoc("TITLE_NEW_COL"), mod_Schema_Manager.SanitizeFieldName(strExcelField))

                    If Len(Trim$(strNewTarget)) > 0 Then
                        Dim strTypeSel As String
                        Dim strSqlType As String

                        strTypeSel = InputBox(mod_UI_Helpers.GetLoc("PROMPT_SELECT_DATA_TYPE"), mod_UI_Helpers.GetLoc("TITLE_SCHEMA_MANAGER"), "1")
                        Select Case Trim$(strTypeSel)
                            Case "2": strSqlType = "DATETIME"
                            Case "3": strSqlType = "LONG"
                            Case "4": strSqlType = "LONGTEXT"
                            Case "5": strSqlType = "BIT"
                            Case "": strSqlType = "" ' Отмена
                            Case Else: strSqlType = "VARCHAR(255)"
                        End Select

                        If strSqlType <> "" Then
                            ' 1. Добавляем поле в саму базу
                            If Not mod_Schema_Manager.FieldExists("tbl_Personnel_Master", strNewTarget) Then
                                mod_Schema_Manager.AddNewFieldToSchema strNewTarget, strSqlType
                            End If

                            ' 2. Записываем связь в маппинг
                            Dim qdfAdd As DAO.QueryDef
                            Set qdfAdd = db.CreateQueryDef("", "PARAMETERS prmP Long, prmE Text(255), prmT Text(100); INSERT INTO tbl_Import_Mapping (ProfileID, ExcelHeader, TargetField) VALUES ([prmP], [prmE], [prmT]);")
                            qdfAdd.Parameters("prmP").value = lngBestProfile
                            qdfAdd.Parameters("prmE").value = Left$(strExcelField, 255)
                            qdfAdd.Parameters("prmT").value = Left$(strNewTarget, 100)
                            qdfAdd.Execute dbFailOnError
                            Set qdfAdd = Nothing

                            ' 3. Подхватываем поле и идем дальше!
                            strAccessField = strNewTarget
                        End If
                    End If
                End If
            End If

            ' Если поле так и осталось пустым (пользователь нажал НЕТ) - пропускаем
            If Len(strAccessField) = 0 Then
                Debug.Print "Skipping unmapped column: " & strExcelField
                If Len(outSkippedColumns) > 0 Then outSkippedColumns = outSkippedColumns & ", "
                outSkippedColumns = outSkippedColumns & "[" & strExcelField & "]"
                GoTo NextColumn
            End If
        End If
        ' ========================================================

        If Len(strPersonUIDExcelName) = 0 And NormalizeString(strAccessField) = "PERSONUID" Then
            strPersonUIDExcelName = strExcelField
        End If

        If Not RegisterDestField(colUsedFields, strAccessField) Then
            GoTo NextColumn
        End If

        mod_Schema_Manager.EnsureFieldExists "tbl_Import_Buffer", strAccessField, "TEXT(255)"

        If Len(strSelectPart) > 0 Then strSelectPart = strSelectPart & ", "
        If Len(strInsertPart) > 0 Then strInsertPart = strInsertPart & ", "

        strSelectPart = strSelectPart & "[" & strExcelField & "]"
        strInsertPart = strInsertPart & "[" & strAccessField & "]"

NextColumn:
    Next fld

    If Len(strInsertPart) = 0 Then
        If Not blnSuppressMsgBox Then MsgBox "No mapped columns matched.", vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
        RunDynamicImport = False
        Exit Function
    End If

    If Len(strPersonUIDExcelName) = 0 Then
        If Not blnSuppressMsgBox Then MsgBox "CRITICAL: No column mapped to PersonUID.", vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
        RunDynamicImport = False
        Exit Function
    End If

    Dim strSQL As String
    strSQL = "INSERT INTO tbl_Import_Buffer (" & strInsertPart & ") " & _
             "SELECT " & strSelectPart & " " & _
             "FROM [" & cstrLinkedTableName & "] " & _
             "WHERE [" & strPersonUIDExcelName & "] IS NOT NULL;"
    db.Execute strSQL, dbFailOnError

    RunDynamicImport = True
    Exit Function

ErrorHandler:
    If Not blnSuppressMsgBox Then
        MsgBox "Import error: " & Err.Description, vbCritical, mod_UI_Helpers.GetLoc("TITLE_ERROR")
    End If
    RunDynamicImport = False
End Function

Private Function GetMappedFieldFromTable(strExcelField As String, lngProfileID As Long) As String
    On Error GoTo ErrorHandler
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim strNormExcel As String

    GetMappedFieldFromTable = ""
    strNormExcel = NormalizeString(strExcelField)
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT ExcelHeader, TargetField FROM tbl_Import_Mapping WHERE ProfileID = " & lngProfileID, dbOpenSnapshot)
    Do While Not rs.EOF
        If NormalizeString(Nz(rs!ExcelHeader, "")) = strNormExcel Then
            GetMappedFieldFromTable = Trim$(Nz(rs!targetField, ""))
            Exit Do
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
ErrorHandler:
    On Error Resume Next: rs.Close: Set rs = Nothing: Set db = Nothing
End Function

Private Function RegisterDestField(colUsed As Collection, strFieldName As String) As Boolean
    On Error GoTo AlreadyExists
    colUsed.Add strFieldName, UCase(strFieldName)
    RegisterDestField = True
    Exit Function
AlreadyExists:
    RegisterDestField = False
    Err.Clear
End Function

Private Function LinkExcelFile(strPath As String, Optional ByVal blnSuppressMsgBox As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    Dim db As DAO.Database, tdf As DAO.TableDef
    Dim strConnect As String, strSheet As String

    DeleteExcelLink
    Set db = CurrentDb

    strSheet = GetFirstSheetName(strPath)
    If strSheet = "" Then Exit Function
    If Right(strSheet, 1) <> "$" Then strSheet = strSheet & "$"

    If Right(strPath, 4) = ".xls" Then
        strConnect = "Excel 8.0;HDR=YES;IMEX=1;DATABASE=" & strPath
    Else
        strConnect = "Excel 12.0 Xml;HDR=YES;IMEX=1;DATABASE=" & strPath
    End If

    Set tdf = db.CreateTableDef(cstrLinkedTableName)
    tdf.Connect = strConnect
    tdf.SourceTableName = strSheet
    db.TableDefs.Append tdf

    LinkExcelFile = True
    Exit Function
ErrorHandler:
    If Not blnSuppressMsgBox Then MsgBox "Link error: " & Err.Description, vbCritical, "Error"
End Function

Private Function GetFirstSheetName(strPath As String) As String
    Dim dbExcel As DAO.Database
    Dim tdf As DAO.TableDef

    GetFirstSheetName = ""
    On Error Resume Next

    Dim strConnect As String
    If Right(LCase(strPath), 4) = ".xls" Then
        strConnect = "Excel 8.0;HDR=YES;IMEX=1;"
    Else
        strConnect = "Excel 12.0 Xml;HDR=YES;IMEX=1;"
    End If

    Set dbExcel = DBEngine.Workspaces(0).OpenDatabase(strPath, False, True, strConnect)

    If Err.Number = 0 Then
        For Each tdf In dbExcel.TableDefs
            If Right(tdf.Name, 1) = "$" Or Right(tdf.Name, 2) = "$'" Then
                GetFirstSheetName = Replace(tdf.Name, "'", "")
                Exit For
            End If
        Next tdf
        dbExcel.Close
    End If

    Set tdf = Nothing
    Set dbExcel = Nothing
End Function

Public Function SelectExcelFile() As String
    On Error GoTo ErrorHandler
    Dim fd As Object
    Dim strInitialPath As String
    strInitialPath = Trim$(Nz(mod_Maintenance_Logic.GetSetting("ImportFolderPath", ""), ""))
    If Len(strInitialPath) = 0 Then strInitialPath = CurrentProject.Path
    If Len(strInitialPath) > 0 And Right(strInitialPath, 1) <> "\" Then strInitialPath = strInitialPath & "\"
    Set fd = Application.FileDialog(3)
    With fd
        .Filters.Clear
        .Filters.Add "Excel files", "*.xls;*.xlsx"
        If Len(strInitialPath) > 0 Then .InitialFileName = strInitialPath
        If .Show = -1 Then SelectExcelFile = .SelectedItems(1)
    End With
    Exit Function
ErrorHandler:
    SelectExcelFile = ""
End Function

Private Sub DeleteExcelLink()
    On Error Resume Next
    CurrentDb.TableDefs.Delete cstrLinkedTableName
End Sub

Private Function GetFileModificationDate(strFilePath As String) As Date
    On Error GoTo ErrorHandler
    Dim fso As Object, oFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile = fso.GetFile(strFilePath)
    GetFileModificationDate = oFile.DateLastModified
    Set oFile = Nothing
    Set fso = Nothing
    Exit Function
ErrorHandler:
    GetFileModificationDate = Now()
End Function

Private Sub UpdateImportMetadata(strFilePath As String)
    On Error Resume Next
    Dim db As DAO.Database, qdf As DAO.QueryDef
    Set db = CurrentDb
    db.Execute "DELETE FROM tbl_Import_Meta;", dbFailOnError
    Set qdf = db.CreateQueryDef("", "PARAMETERS prmE DateTime, prmI DateTime, prmP Text(255); INSERT INTO tbl_Import_Meta (ExportFileDate, ImportRunAt, SourceFilePath) VALUES ([prmE], [prmI], [prmP]);")
    qdf.Parameters("prmE").value = GetFileModificationDate(strFilePath)
    qdf.Parameters("prmI").value = Now()
    qdf.Parameters("prmP").value = Left$(strFilePath, 255)
    qdf.Execute dbFailOnError
End Sub
