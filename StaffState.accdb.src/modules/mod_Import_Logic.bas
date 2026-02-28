Attribute VB_Name = "mod_Import_Logic"
Option Compare Database
Option Explicit

' =============================================
' @module mod_Import_Logic (DYNAMIC v2.0)
' @description Dynamic import of any columns from Excel
' @note 100% English version. Safe for modern IDEs.
' =============================================

Private Const cstrLinkedTableName As String = "tmp_Excel_Link"

Public Function ImportExcelData(Optional ByVal blnSuppressMsgBox As Boolean = False) As Boolean
    On Error GoTo ErrorHandler

    ImportExcelData = False

    Dim strFilePath As String
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

    ' 3. Dynamic import
    If Not RunDynamicImport(blnSuppressMsgBox) Then GoTo ExitHandler

    ' 4. Save import metadata
    UpdateImportMetadata strFilePath

    ImportExcelData = True

    If Not blnSuppressMsgBox Then
        MsgBox "Import completed successfully. Only mapped English columns were filled.", vbInformation, "StaffState Import"
    End If

ExitHandler:
    DeleteExcelLink
    Exit Function

ErrorHandler:
    If Not blnSuppressMsgBox Then
        MsgBox "Critical Error during import: " & Err.Description, vbCritical, "System Error"
    Else
        Debug.Print "Import error: " & Err.Description & " (" & Err.Number & ")"
    End If
    Resume ExitHandler
End Function

Private Function NormalizeString(ByVal s As String) As String
    NormalizeString = UCase(Trim$(s))
End Function

Private Function RunDynamicImport(Optional ByVal blnSuppressMsgBox As Boolean = False) As Boolean
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

    strPersonUIDExcelName = ""
    strAllColumns = ""
    Set colUsedFields = New Collection
    Set db = CurrentDb

    ' --- EARLY VALIDATE: tbl_Import_Mapping must exist ---
    If Not mod_Schema_Manager.TableExists("tbl_Import_Mapping") Then
        If Not blnSuppressMsgBox Then
            MsgBox "Import failed: tbl_Import_Mapping table is missing. Run InitDatabaseStructure or create the table and run SeedImportMappingProfile1.", vbCritical, "Import Error"
        End If
        RunDynamicImport = False
        Exit Function
    End If
    If GetMappingCountForProfile(1) = 0 Then
        If Not blnSuppressMsgBox Then
            MsgBox "Import failed: tbl_Import_Mapping is empty for Profile 1. Run mod_Schema_Manager.SeedImportMappingProfile1.", vbCritical, "Import Error"
        End If
        RunDynamicImport = False
        Exit Function
    End If

    Set tdfLink = db.TableDefs(cstrLinkedTableName)

    ' --- LOOP THROUGH ALL EXCEL COLUMNS ---
    For Each fld In tdfLink.Fields
        strExcelField = fld.Name

        If Len(strAllColumns) > 0 Then strAllColumns = strAllColumns & ", "
        strAllColumns = strAllColumns & "[" & strExcelField & "]"

        strAccessField = GetMappedFieldFromTable(strExcelField, 1)
        If Len(strAccessField) = 0 Then
            Debug.Print "Skipping unmapped column: " & strExcelField
            GoTo NextColumn
        End If

        If Len(strPersonUIDExcelName) = 0 And NormalizeString(strAccessField) = "PERSONUID" Then
            strPersonUIDExcelName = strExcelField
            Debug.Print "PersonUID column: " & strExcelField
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

    Debug.Print "All Excel columns: " & strAllColumns

    ' --- VALIDATE: at least one mapped column ---
    If Len(strInsertPart) = 0 Then
        If Not blnSuppressMsgBox Then
            MsgBox "No mapped columns. Your Excel headers did not match any row in tbl_Import_Mapping (Profile 1)." & vbCrLf & vbCrLf & _
                   "Excel columns in file: " & Left(strAllColumns, 400), vbCritical, "Import Error"
        End If
        RunDynamicImport = False
        Exit Function
    End If

    ' --- VALIDATE: PersonUID column must be present ---
    If Len(strPersonUIDExcelName) = 0 Then
        If Not blnSuppressMsgBox Then
            MsgBox "CRITICAL: No Excel column maps to PersonUID. Add a mapping in tbl_Import_Mapping." & vbCrLf & vbCrLf & _
                   "Excel columns in your file: " & Left(strAllColumns, 400), vbCritical, "Import Error"
        End If
        RunDynamicImport = False
        Exit Function
    End If

    ' --- FINAL SQL ---
    Dim strSQL As String
    strSQL = "INSERT INTO tbl_Import_Buffer (" & strInsertPart & ") " & _
             "SELECT " & strSelectPart & " " & _
             "FROM [" & cstrLinkedTableName & "] " & _
             "WHERE [" & strPersonUIDExcelName & "] IS NOT NULL;"
    Debug.Print "Dynamic SQL generated: " & strSQL
    db.Execute strSQL, dbFailOnError

    RunDynamicImport = True
    Exit Function

ErrorHandler:
    If Not blnSuppressMsgBox Then
        MsgBox "Import error: " & Err.Description & vbCrLf & "Error number: " & Err.Number, vbCritical, "Dynamic Import Error"
    Else
        Debug.Print "Dynamic import error: " & Err.Description & " (" & Err.Number & ")"
    End If
End Function

Private Function GetMappingCountForProfile(lngProfileID As Long) As Long
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    GetMappingCountForProfile = 0
    If Not mod_Schema_Manager.TableExists("tbl_Import_Mapping") Then Exit Function
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS Cnt FROM tbl_Import_Mapping WHERE ProfileID = " & lngProfileID, dbOpenSnapshot)
    If Not rs.EOF Then GetMappingCountForProfile = Nz(rs!Cnt, 0)
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
ErrorHandler:
    If Not rs Is Nothing Then On Error Resume Next: rs.Close: Set rs = Nothing
    If Not db Is Nothing Then Set db = Nothing
    GetMappingCountForProfile = 0
End Function

Private Function GetMappedFieldFromTable(strExcelField As String, lngProfileID As Long) As String
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim strNormExcel As String

    GetMappedFieldFromTable = ""
    If Not mod_Schema_Manager.TableExists("tbl_Import_Mapping") Then Exit Function

    strNormExcel = NormalizeString(strExcelField)
    Set db = CurrentDb
    strSQL = "SELECT ExcelHeader, TargetField FROM tbl_Import_Mapping WHERE ProfileID = " & lngProfileID
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    Do While Not rs.EOF
        If NormalizeString(Nz(rs!ExcelHeader, "")) = strNormExcel Then
            GetMappedFieldFromTable = Trim$(Nz(rs!TargetField, ""))
            Exit Do
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
ErrorHandler:
    If Not rs Is Nothing Then
        On Error Resume Next
        rs.Close
        Set rs = Nothing
    End If
    If Not db Is Nothing Then Set db = Nothing
    GetMappedFieldFromTable = ""
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

' --- RUSSIAN MAPPINGS (fuzzy by patterns) - Kept for legacy fallback if mapping table fails ---
Private Function MapFieldName(strExcelField As String) As String
    Dim s As String
    s = LCase(strExcelField)

    If s = "personuid" Or s = "uid" Then MapFieldName = "PersonUID": Exit Function
    If s = "sourceid" Then MapFieldName = "SourceID": Exit Function
    If s = "fullname" Then MapFieldName = "FullName": Exit Function
    If s = "rank" Then MapFieldName = "RankName": Exit Function
    If s = "birthdate" Then MapFieldName = "BirthDate_Text": Exit Function
    If s = "workstatus" Then MapFieldName = "WorkStatus": Exit Function
    If s = "poscode" Then MapFieldName = "PosCode": Exit Function
    If s = "posname" Then MapFieldName = "PosName": Exit Function
    If s = "orderdate" Then MapFieldName = "OrderDate_Text": Exit Function
    If s = "ordernum" Or s = "ordernumber" Then MapFieldName = "OrderNumber": Exit Function

    If ContainsCyrillic(strExcelField) Then
        ' Legacy pattern mapping preserved using ChrW to avoid IDE breaking
        If MatchCyrPattern(s, Array(1083, 1080, 1094, 1086)) And Len(s) < 6 Then MapFieldName = "SourceID": Exit Function
        If MatchCyrPattern(s, Array(1092, 1080, 1086)) Then MapFieldName = "FullName": Exit Function
        If MatchCyrPattern(s, Array(1079, 1074, 1072, 1085, 1080, 1077)) Then MapFieldName = "RankName": Exit Function
        If MatchCyrPattern(s, Array(1088, 1086, 1078, 1076, 1077, 1085, 1080, 1103)) Then MapFieldName = "BirthDate_Text": Exit Function
    End If

    MapFieldName = mod_Schema_Manager.SanitizeFieldName(strExcelField)
End Function

Private Function ContainsCyrillic(s As String) As Boolean
    Dim i As Long
    Dim c As Long

    For i = 1 To Len(s)
        c = AscW(Mid(s, i, 1))
        If c >= 1024 And c <= 1279 Then
            ContainsCyrillic = True
            Exit Function
        End If
        If c >= 192 And c <= 255 Then
            ContainsCyrillic = True
            Exit Function
        End If
    Next i
    ContainsCyrillic = False
End Function

Private Function MatchCyrPattern(s As String, pattern As Variant) As Boolean
    Dim patternStr As String
    Dim i As Long

    patternStr = ""
    For i = LBound(pattern) To UBound(pattern)
        patternStr = patternStr & ChrW(pattern(i))
    Next i

    MatchCyrPattern = (InStr(LCase(s), LCase(patternStr)) > 0)
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
    If Not blnSuppressMsgBox Then
        MsgBox "Link error: " & Err.Description, vbCritical, "Error"
    Else
        Debug.Print "Link error: " & Err.Description & " (" & Err.Number & ")"
    End If
End Function

Private Function GetFirstSheetName(strPath As String) As String
    Dim xlApp As Object, xlWb As Object
    On Error Resume Next
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlWb = xlApp.Workbooks.Open(strPath, False, True)
    GetFirstSheetName = xlWb.Sheets(1).Name
    xlWb.Close False: xlApp.Quit
    Set xlApp = Nothing: Set xlWb = Nothing
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
    Debug.Print "SelectExcelFile error: " & Err.Description & " (" & Err.Number & ")"
    SelectExcelFile = ""
End Function

Private Sub DeleteExcelLink()
    On Error Resume Next
    CurrentDb.TableDefs.Delete cstrLinkedTableName
End Sub

Private Function GetFileModificationDate(strFilePath As String) As Date
    On Error GoTo ErrorHandler
    Dim fso As Object
    Dim oFile As Object

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
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim dtFileDate As Date
    Dim strSQL As String

    Set db = CurrentDb
    dtFileDate = GetFileModificationDate(strFilePath)
    db.Execute "DELETE FROM tbl_Import_Meta;", dbFailOnError

    Dim strFileDate As String, strNowDate As String
    strFileDate = Month(dtFileDate) & "/" & Day(dtFileDate) & "/" & Year(dtFileDate) & " " & Format(dtFileDate, "hh:nn:ss")
    strNowDate = Month(Now()) & "/" & Day(Now()) & "/" & Year(Now()) & " " & Format(Now(), "hh:nn:ss")

    strSQL = "INSERT INTO tbl_Import_Meta (ExportFileDate, ImportRunAt, SourceFilePath) " & _
             "VALUES (#" & strFileDate & "#, #" & strNowDate & "#, '" & Replace(strFilePath, "'", "''") & "');"
    db.Execute strSQL, dbFailOnError

    Debug.Print "Import metadata updated. ExportFileDate: " & dtFileDate
    Set db = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print "Warning: Failed to update import metadata: " & Err.Description
End Sub
