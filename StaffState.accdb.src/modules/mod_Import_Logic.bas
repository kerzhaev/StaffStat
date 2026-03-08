Attribute VB_Name = "mod_Import_Logic"
Option Compare Database
Option Explicit

' =============================================
' @module mod_Import_Logic (DYNAMIC v4.2 - Bugfix & Smart Routing)
' @description Dynamic auto-detect import + Interactive wizard + Content-based routing
' @note 100% English version. Safe for modern IDEs.
' =============================================

Private Const cstrLinkedTableName As String = "tmp_Excel_Link"

Public Function ImportExcelData(Optional ByVal blnSuppressMsgBox As Boolean = False) As Boolean
    Dim result As Object

    Set result = ImportExcelDataResult("", blnSuppressMsgBox)
    ImportExcelData = CBool(result("Success"))
    Set result = Nothing
End Function

Public Function ImportExcelDataResult(Optional ByVal strFilePath As String = "", Optional ByVal blnSuppressMsgBox As Boolean = False) As Object
    On Error GoTo ErrorHandler

    Dim result As Object
    Dim strSkipped As String
    Dim strErrorMessage As String

    Set result = CreateImportResult()

    If Len(Trim$(strFilePath)) = 0 Then
        strFilePath = SelectExcelFile()
    End If
    result("FilePath") = strFilePath

    If strFilePath = "" Then
        result("Cancelled") = True
        result("Message") = "Import canceled."
        GoTo ExitHandler
    End If

    DoCmd.Close acTable, "tbl_Import_Buffer", acSaveYes
    DoCmd.Close acTable, "tbl_Personnel_Master", acSaveYes
    DoEvents

    ' 1. Clear buffer
    CurrentDb.Execute "DELETE FROM tbl_Import_Buffer;", dbFailOnError

    ' 2. Link
    If Not LinkExcelFile(strFilePath, strErrorMessage, blnSuppressMsgBox) Then
        result("ErrorMessage") = strErrorMessage
        result("Message") = strErrorMessage
        GoTo ExitHandler
    End If

    ' 3. Dynamic import (with Auto-Detect and Interactive Wizard)
    If Not RunDynamicImport(strSkipped, strErrorMessage, blnSuppressMsgBox) Then
        result("ErrorMessage") = strErrorMessage
        result("Message") = strErrorMessage
        GoTo ExitHandler
    End If

    ' 4. Save import metadata
    UpdateImportMetadata strFilePath

    result("Success") = True
    result("SkippedColumns") = strSkipped
    result("Message") = BuildImportSuccessMessage(strSkipped)

ExitHandler:
    DeleteExcelLink
    Set ImportExcelDataResult = result
    Exit Function

ErrorHandler:
    result("Success") = False
    result("ErrorNumber") = Err.Number
    result("ErrorMessage") = "Import error: " & Err.Description
    result("Message") = CStr(result("ErrorMessage"))
    Debug.Print "Import error: " & Err.Description & " (" & Err.Number & ")"
    Resume ExitHandler
End Function

Private Function NormalizeString(ByVal s As String) As String
    NormalizeString = UCase(Trim$(s))
End Function

' --- ???????? ??????????? ??????? ---
Private Function IsColumnNumeric(strColName As String) As Boolean
    On Error Resume Next
    Dim rs As DAO.Recordset
    Dim bNumeric As Boolean
    bNumeric = False

    ' ????? ????? ?????? ???????? ?????? ?? ???? ??????? Excel
    Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 [" & strColName & "] FROM [" & cstrLinkedTableName & "] WHERE [" & strColName & "] IS NOT NULL", dbOpenSnapshot)

    If Not rs.EOF Then
        bNumeric = IsNumeric(rs.Fields(0).value)
    End If

    rs.Close
    Set rs = Nothing
    IsColumnNumeric = bNumeric
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

    mod_Schema_Manager.CreateImportProfilesTable

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

                ' ?????? ??????????
                If normExcel = normHeader Then
                    matchCount = matchCount + 1
                    If targetField = "PERSONUID" Then hasUID = True
                    Exit For
                ' ????????? ????-????????? ?????????? Excel (????, ????1, ????2)
                ElseIf normHeader = "????" And Left$(normExcel, 4) = "????" Then
                    matchCount = matchCount + 1
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

Private Function RunDynamicImport(ByRef outSkippedColumns As String, Optional ByRef outErrorMessage As String = "", Optional ByVal blnSuppressMsgBox As Boolean = False) As Boolean
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
    outErrorMessage = ""

    If Not mod_Schema_Manager.TableExists("tbl_Import_Mapping") Then
        outErrorMessage = "Import mapping table is missing."
        RunDynamicImport = False
        Exit Function
    End If

    Set tdfLink = db.TableDefs(cstrLinkedTableName)

    lngBestProfile = DetectBestProfile(tdfLink)

    If lngBestProfile = 0 Then
        For Each fld In tdfLink.Fields
            strAllColumns = strAllColumns & "[" & fld.Name & "] "
        Next fld

        outErrorMessage = "Import failed: could not auto-detect a matching mapping profile." & vbCrLf & _
                          "Ensure mapping exists for PersonUID." & vbCrLf & vbCrLf & _
                          "Columns found: " & Left$(strAllColumns, 400)
        RunDynamicImport = False
        Exit Function
    End If

    strAllColumns = ""

    ' --- LOOP THROUGH ALL EXCEL COLUMNS ---
    For Each fld In tdfLink.Fields
        strExcelField = fld.Name

        If Len(strAllColumns) > 0 Then strAllColumns = strAllColumns & ", "
        strAllColumns = strAllColumns & "[" & strExcelField & "]"

        ' ========================================================
        ' SMART ROUTING (?????? ??????????? ??? ?????????? "????")
        ' ========================================================
        If Left$(NormalizeString(strExcelField), 4) = "????" Then
            If IsColumnNumeric(strExcelField) Then
                strAccessField = "SourceID"
            Else
                strAccessField = "FullName"
            End If
        Else
            ' ??????? ????? ? ????????
            strAccessField = GetMappedFieldFromTable(strExcelField, lngBestProfile)
        End If
        ' ========================================================


        ' ========================================================
        ' INTERACTIVE MAPPING WIZARD
        ' ========================================================
        If Len(strAccessField) = 0 Then

            ' ???? ??? ??????? ?????? ?????????? - ?????????? ?????
            If blnSuppressMsgBox Then
                If Len(outSkippedColumns) > 0 Then outSkippedColumns = outSkippedColumns & ", "
                outSkippedColumns = outSkippedColumns & "[" & strExcelField & "]"
                GoTo NextColumn
            End If

            Dim strSuggestedName As String
            Dim blnFieldExists As Boolean
            Dim qdfAdd As DAO.QueryDef

            strSuggestedName = mod_Schema_Manager.SanitizeFieldName(strExcelField)
            blnFieldExists = mod_Schema_Manager.FieldExists("tbl_Personnel_Master", strSuggestedName)

            If blnFieldExists Then
                Dim strPromptRestore As String
                strPromptRestore = mod_UI_Helpers.GetLoc("PROMPT_RESTORE_LINK1") & " [" & strExcelField & "]." & vbCrLf & vbCrLf & _
                                   "???? [" & strSuggestedName & "] " & mod_UI_Helpers.GetLoc("PROMPT_RESTORE_LINK2") & vbCrLf & _
                                   mod_UI_Helpers.GetLoc("PROMPT_RESTORE_LINK3")

                If mod_UI_Helpers.AskUserYesNo(strPromptRestore, mod_UI_Helpers.GetLoc("TITLE_RESTORE_LINK")) Then
                    Set qdfAdd = db.CreateQueryDef("", "PARAMETERS prmP Long, prmE Text(255), prmT Text(100); INSERT INTO tbl_Import_Mapping (ProfileID, ExcelHeader, TargetField) VALUES ([prmP], [prmE], [prmT]);")
                    qdfAdd.Parameters("prmP").value = lngBestProfile
                    qdfAdd.Parameters("prmE").value = Left$(strExcelField, 255)
                    qdfAdd.Parameters("prmT").value = Left$(strSuggestedName, 100)
                    qdfAdd.Execute dbFailOnError
                    Set qdfAdd = Nothing
                    strAccessField = strSuggestedName
                End If
            Else
                If mod_UI_Helpers.AskUserYesNo(mod_UI_Helpers.GetLoc("PROMPT_MAP_NEW_COL") & " [" & strExcelField & "]" & vbCrLf & vbCrLf & mod_UI_Helpers.GetLoc("PROMPT_MAP_NEW_COL2"), mod_UI_Helpers.GetLoc("TITLE_NEW_COL")) Then
                    Dim strNewTarget As String
                    strNewTarget = InputBox(mod_UI_Helpers.GetLoc("PROMPT_ENTER_EN_NAME"), mod_UI_Helpers.GetLoc("TITLE_NEW_COL"), strSuggestedName)

                    If Len(Trim$(strNewTarget)) > 0 Then
                        If mod_Schema_Manager.FieldExists("tbl_Personnel_Master", strNewTarget) Then
                            strAccessField = strNewTarget
                            Set qdfAdd = db.CreateQueryDef("", "PARAMETERS prmP Long, prmE Text(255), prmT Text(100); INSERT INTO tbl_Import_Mapping (ProfileID, ExcelHeader, TargetField) VALUES ([prmP], [prmE], [prmT]);")
                            qdfAdd.Parameters("prmP").value = lngBestProfile
                            qdfAdd.Parameters("prmE").value = Left$(strExcelField, 255)
                            qdfAdd.Parameters("prmT").value = Left$(strNewTarget, 100)
                            qdfAdd.Execute dbFailOnError
                            Set qdfAdd = Nothing
                        Else
                            Dim strTypeSel As String
                            Dim strSqlType As String

                            strTypeSel = InputBox(mod_UI_Helpers.GetLoc("PROMPT_SELECT_DATA_TYPE"), mod_UI_Helpers.GetLoc("TITLE_SCHEMA_MANAGER"), "1")
                            Select Case Trim$(strTypeSel)
                                Case "2": strSqlType = "DATETIME"
                                Case "3": strSqlType = "LONG"
                                Case "4": strSqlType = "LONGTEXT"
                                Case "5": strSqlType = "BIT"
                                Case "": strSqlType = ""
                                Case Else: strSqlType = "VARCHAR(255)"
                            End Select

                            If strSqlType <> "" Then
                                mod_Schema_Manager.AddNewFieldToSchema strNewTarget, strSqlType
                                Set qdfAdd = db.CreateQueryDef("", "PARAMETERS prmP Long, prmE Text(255), prmT Text(100); INSERT INTO tbl_Import_Mapping (ProfileID, ExcelHeader, TargetField) VALUES ([prmP], [prmE], [prmT]);")
                                qdfAdd.Parameters("prmP").value = lngBestProfile
                                qdfAdd.Parameters("prmE").value = Left$(strExcelField, 255)
                                qdfAdd.Parameters("prmT").value = Left$(strNewTarget, 100)
                                qdfAdd.Execute dbFailOnError
                                Set qdfAdd = Nothing
                                strAccessField = strNewTarget
                            End If
                        End If
                    End If
                End If
            End If

            If Len(strAccessField) = 0 Then
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
        outErrorMessage = "No mapped columns matched."
        RunDynamicImport = False
        Exit Function
    End If

    If Len(strPersonUIDExcelName) = 0 Then
        outErrorMessage = "Critical import error: no column is mapped to PersonUID."
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
    outErrorMessage = "Import error: " & Err.Description
    Debug.Print "RunDynamicImport error: " & Err.Description & " (" & Err.Number & ")"
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

Private Function LinkExcelFile(strPath As String, Optional ByRef outErrorMessage As String = "", Optional ByVal blnSuppressMsgBox As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    Dim db As DAO.Database, tdf As DAO.TableDef
    Dim strConnect As String, strSheet As String

    outErrorMessage = ""
    DeleteExcelLink
    Set db = CurrentDb

    strSheet = GetFirstSheetName(strPath)
    If strSheet = "" Then
        outErrorMessage = "Link error: could not detect a worksheet in the selected Excel file."
        Exit Function
    End If
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
    outErrorMessage = "Link error: " & Err.Description
    Debug.Print "LinkExcelFile error: " & Err.Description & " (" & Err.Number & ")"
End Function

' ????????????? ?????? ????? EXCEL ????? ADODB
Private Function GetFirstSheetName(strPath As String) As String
    Dim conn As Object
    Dim rs As Object

    GetFirstSheetName = ""
    On Error Resume Next

    Set conn = CreateObject("ADODB.Connection")
    If Right(LCase(strPath), 4) = ".xls" Then
        conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strPath & ";Extended Properties=""Excel 8.0;HDR=YES;"";"
    Else
        conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strPath & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;"";"
    End If

    If Err.Number <> 0 Then
        Err.Clear
        conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPath & ";Extended Properties=""Excel 8.0;HDR=YES;"";"
    End If

    If conn.State = 1 Then ' ???? ????????? ???????
        Set rs = conn.OpenSchema(20)
        Do While Not rs.EOF
            Dim sName As String
            sName = rs.Fields("TABLE_NAME").value
            If Right(sName, 1) = "$" Or Right(sName, 2) = "$'" Then
                GetFirstSheetName = Replace(sName, "'", "")
                Exit Do
            End If
            rs.MoveNext
        Loop
        rs.Close
        conn.Close
    End If

    Set rs = Nothing
    Set conn = Nothing

    ' Fallback ???? ADODB ???????????? ?? ??
    If GetFirstSheetName = "" Then
        Err.Clear
        Dim xlApp As Object, xlWb As Object
        Set xlApp = CreateObject("Excel.Application")
        xlApp.Visible = False
        Set xlWb = xlApp.Workbooks.Open(strPath, False, True)
        GetFirstSheetName = xlWb.Sheets(1).Name
        xlWb.Close False
        xlApp.Quit
        Set xlWb = Nothing
        Set xlApp = Nothing
    End If
End Function

Public Function SelectExcelFile() As String
    On Error GoTo ErrorHandler
    Dim fd As Object
    Dim strInitialPath As String
    Dim fso As Object

    ' 1. Получаем путь из настроек
    strInitialPath = Trim$(Nz(mod_Maintenance_Logic.GetSetting("ImportFolderPath", ""), ""))
    If UCase$(strInitialPath) = "N/A" Then strInitialPath = ""

    ' 2. УМНАЯ ПРОВЕРКА: существует ли папка физически?
    If Len(strInitialPath) > 0 Then
        On Error Resume Next
        Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FolderExists(strInitialPath) Then
            ' Если пользователь опечатался или папку удалили - сбрасываем путь
            strInitialPath = ""
            Debug.Print "Warning: Import folder from settings does not exist. Falling back to DB path."
        End If
        Set fso = Nothing
        On Error GoTo ErrorHandler
    End If

    ' 3. Если путь пустой (или был сброшен из-за ошибки), берем папку, где лежит сама база
    If Len(strInitialPath) = 0 Then strInitialPath = CurrentProject.Path

    ' 4. Форматируем путь (добавляем слеш на конце)
    If Len(strInitialPath) > 0 And Right(strInitialPath, 1) <> "\" Then strInitialPath = strInitialPath & "\"

    ' 5. Открываем окно выбора файла
    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker
    With fd
        .Filters.Clear
        .Filters.Add "Excel files", "*.xls;*.xlsx"
        .InitialFileName = strInitialPath
        If .Show = -1 Then SelectExcelFile = .SelectedItems(1)
    End With

    Exit Function
ErrorHandler:
    Debug.Print "SelectExcelFile error: " & Err.Description & " (" & Err.Number & ")"
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

Private Function CreateImportResult() As Object
    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d("Success") = False
    d("Cancelled") = False
    d("FilePath") = ""
    d("SkippedColumns") = ""
    d("Message") = ""
    d("ErrorMessage") = ""
    d("ErrorNumber") = 0

    Set CreateImportResult = d
End Function

Private Function BuildImportSuccessMessage(ByVal strSkipped As String) As String
    Dim strMessage As String

    strMessage = mod_UI_Helpers.GetLoc("MSG_IMPORT_SUCCESS")
    If Len(strSkipped) > 0 Then
        strMessage = strMessage & vbCrLf & vbCrLf & _
                     mod_UI_Helpers.GetLoc("MSG_SKIPPED_COLS") & vbCrLf & _
                     strSkipped
    End If

    BuildImportSuccessMessage = strMessage
End Function
