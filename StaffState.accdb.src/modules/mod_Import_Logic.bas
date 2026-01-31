Attribute VB_Name = "mod_Import_Logic"
Option Explicit

' =============================================
' @module mod_Import_Logic (DYNAMIC v2.0)
' @description Dynamic import of any columns from Excel
' =============================================

Private Const cstrLinkedTableName As String = "tmp_Excel_Link"

Public Function ImportExcelData() As Boolean
    On Error GoTo ErrorHandler

    Dim strFilePath As String
    strFilePath = SelectExcelFile()
    If strFilePath = "" Then Exit Function

    ' --- FIX: FORCE CLOSE TABLE BEFORE OPERATION ---
    ' acSaveYes will save layout changes if any
    DoCmd.Close acTable, "tbl_Import_Buffer", acSaveYes
    DoCmd.Close acTable, "tbl_Personnel_Master", acSaveYes
    DoEvents ' Give Access time to breathe
    ' -----------------------------------------------------------

    ' 1. Clear buffer
    CurrentDb.Execute "DELETE FROM tbl_Import_Buffer;", dbFailOnError

    ' 2. Link
    If Not LinkExcelFile(strFilePath) Then Exit Function

    ' 3. Dynamic import
    If Not RunDynamicImport() Then Exit Function

    ' 4. Save import metadata (file date for change tracking)
    UpdateImportMetadata strFilePath

    MsgBox "Import completed. New columns were added.", vbInformation
    ImportExcelData = True

    DeleteExcelLink
    Exit Function

ErrorHandler:
    MsgBox "Critical Error: " & Err.Description, vbCritical
    DeleteExcelLink
End Function

' =============================================
' @description Main magic: reads Excel fields, extends Buffer and builds SQL
' @note Uses fuzzy PersonUID field detection (encoding-independent)
' =============================================
Private Function RunDynamicImport() As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim tdfLink As DAO.TableDef
    Dim fld As DAO.Field
    Dim strExcelField As String
    Dim strAccessField As String
    ' Field lists for SQL
    Dim strSelectPart As String
    Dim strInsertPart As String
    
    ' PersonUID field detection (original Excel name)
    Dim strPersonUIDExcelName As String
    strPersonUIDExcelName = ""
    
    ' Debug: collect all column names
    Dim strAllColumns As String
    strAllColumns = ""
    
    ' Track destination fields to prevent duplicates in INSERT
    Dim colUsedFields As Collection
    Set colUsedFields = New Collection
    
    ' Track duplicate counters for base field names
    Dim colDupCounts As Collection
    Set colDupCounts = New Collection

    Set db = CurrentDb
    Set tdfLink = db.TableDefs(cstrLinkedTableName)

    ' --- LOOP THROUGH ALL EXCEL COLUMNS ---
    For Each fld In tdfLink.Fields
        strExcelField = fld.Name
        
        ' Collect column names for debug
        If Len(strAllColumns) > 0 Then strAllColumns = strAllColumns & ", "
        strAllColumns = strAllColumns & "[" & strExcelField & "]"

        ' 1. NAME MAPPING using fuzzy match (encoding-independent)
        strAccessField = MapFieldName(strExcelField)
        
        ' 2. Detect PersonUID column (fuzzy match)
        If Len(strPersonUIDExcelName) = 0 Then
            If IsPersonUIDColumn(strExcelField) Then
                strPersonUIDExcelName = strExcelField
                strAccessField = "PersonUID_Raw"
                Debug.Print "PersonUID detected: " & strExcelField
            End If
        End If

        ' 3. Make destination field unique (keep duplicates with suffix)
        strAccessField = MakeUniqueDestField(colUsedFields, colDupCounts, strAccessField, strExcelField)

        ' 4. Check if field exists in Buffer
        mod_Schema_Manager.EnsureFieldExists "tbl_Import_Buffer", strAccessField, "TEXT(255)"

        ' 5. Build SQL
        If Len(strSelectPart) > 0 Then strSelectPart = strSelectPart & ", "
        If Len(strInsertPart) > 0 Then strInsertPart = strInsertPart & ", "

        strSelectPart = strSelectPart & "[" & strExcelField & "]"
        strInsertPart = strInsertPart & "[" & strAccessField & "]"

    Next fld
    
    Debug.Print "All Excel columns: " & strAllColumns

    ' --- VALIDATE: PersonUID field must exist ---
    If Len(strPersonUIDExcelName) = 0 Then
        MsgBox "CRITICAL: Excel file has no PersonUID column!" & vbCrLf & vbCrLf & _
               "Found columns:" & vbCrLf & strAllColumns & vbCrLf & vbCrLf & _
               "Looking for column containing: 'number', 'UID', 'PersonUID'", _
               vbCritical, "Import Error"
        RunDynamicImport = False
        Exit Function
    End If

    ' --- FINAL SQL (uses dynamically detected PersonUID field name) ---
    Dim strSQL As String
    strSQL = "INSERT INTO tbl_Import_Buffer (" & strInsertPart & ") " & _
             "SELECT " & strSelectPart & " " & _
             "FROM [" & cstrLinkedTableName & "] " & _
             "WHERE [" & strPersonUIDExcelName & "] IS NOT NULL;"

    Debug.Print "Dynamic SQL generated: " & strSQL
    Debug.Print "PersonUID Excel column: " & strPersonUIDExcelName
    db.Execute strSQL, dbFailOnError

    RunDynamicImport = True
    Exit Function

ErrorHandler:
    MsgBox "Import error: " & Err.Description & vbCrLf & _
           "Error number: " & Err.Number, vbCritical, "Dynamic Import"
End Function

' =============================================
' @description Registers destination field, avoids duplicates
' @param colUsed [Collection] Used field names
' @param strFieldName [String] Destination field name
' @return [Boolean] True if field was registered
' =============================================
Private Function RegisterDestField(colUsed As Collection, strFieldName As String) As Boolean
    On Error GoTo AlreadyExists
    colUsed.Add strFieldName, UCase(strFieldName)
    RegisterDestField = True
    Exit Function
AlreadyExists:
    RegisterDestField = False
    Err.Clear
End Function

' =============================================
' @description Ensures destination field is unique by adding suffix
' @param colUsed [Collection] Used field names
' @param colCounts [Collection] Duplicate counters
' @param strBaseField [String] Base destination field name
' @param strExcelField [String] Original Excel field name (for debug)
' @return [String] Unique destination field name
' =============================================
Private Function MakeUniqueDestField(colUsed As Collection, colCounts As Collection, _
                                     ByVal strBaseField As String, ByVal strExcelField As String) As String
    Dim strCandidate As String
    Dim iDup As Long
    
    strCandidate = strBaseField
    
    If RegisterDestField(colUsed, strCandidate) Then
        MakeUniqueDestField = strCandidate
        Exit Function
    End If
    
    Do
        iDup = NextDuplicateIndex(colCounts, strBaseField)
        strCandidate = strBaseField & "_" & iDup
    Loop While Not RegisterDestField(colUsed, strCandidate)
    
    Debug.Print "Duplicate mapped: " & strExcelField & " -> " & strCandidate
    MakeUniqueDestField = strCandidate
End Function

' =============================================
' @description Returns next duplicate index (2,3,4...)
' @param colCounts [Collection] Duplicate counters
' @param strBaseField [String] Base field name
' @return [Long] Next index
' =============================================
Private Function NextDuplicateIndex(colCounts As Collection, ByVal strBaseField As String) As Long
    Dim key As String
    Dim count As Long
    
    key = UCase(strBaseField)
    On Error Resume Next
    count = colCounts.Item(key)
    If Err.Number <> 0 Then
        Err.Clear
        count = 2
    Else
        count = count + 1
    End If
    On Error GoTo 0
    
    On Error Resume Next
    colCounts.Remove key
    On Error GoTo 0
    colCounts.Add count, key
    
    NextDuplicateIndex = count
End Function

' =============================================
' @description Detects if column is PersonUID (encoding-independent fuzzy match)
' @param strFieldName [String] Column name from Excel
' @return [Boolean] True if this is PersonUID column
' =============================================
Private Function IsPersonUIDColumn(strFieldName As String) As Boolean
    Dim s As String
    s = LCase(strFieldName)
    
    ' English patterns
    If InStr(s, "personuid") > 0 Then IsPersonUIDColumn = True: Exit Function
    If InStr(s, "uid") > 0 And Len(s) < 10 Then IsPersonUIDColumn = True: Exit Function
    
    ' Russian patterns (works regardless of encoding)
    ' Check for Cyrillic by looking at byte values
    If ContainsCyrillic(strFieldName) Then
        ' Look for patterns like "личн" + "номер" or just "номер"
        If InStr(s, ChrW(1085) & ChrW(1086) & ChrW(1084) & ChrW(1077) & ChrW(1088)) > 0 Then
            ' Contains "номер" (Unicode)
            IsPersonUIDColumn = True: Exit Function
        End If
        If InStr(s, Chr(237) & Chr(238) & Chr(236) & Chr(229) & Chr(240)) > 0 Then
            ' Contains "номер" (CP1251)
            IsPersonUIDColumn = True: Exit Function
        End If
    End If
    
    IsPersonUIDColumn = False
End Function

' =============================================
' @description Checks if string contains Cyrillic characters
' =============================================
Private Function ContainsCyrillic(s As String) As Boolean
    Dim i As Long
    Dim c As Long
    
    For i = 1 To Len(s)
        c = AscW(Mid(s, i, 1))
        ' Cyrillic Unicode range: U+0400 to U+04FF
        If c >= 1024 And c <= 1279 Then
            ContainsCyrillic = True
            Exit Function
        End If
        ' CP1251 Cyrillic range: 192-255
        If c >= 192 And c <= 255 Then
            ContainsCyrillic = True
            Exit Function
        End If
    Next i
    
    ContainsCyrillic = False
End Function

' =============================================
' @description Maps Excel field name to Access field name (encoding-independent)
' @param strExcelField [String] Original column name from Excel
' @return [String] Mapped field name for Buffer table
' =============================================
Private Function MapFieldName(strExcelField As String) As String
    Dim s As String
    s = LCase(strExcelField)
    
    ' --- ENGLISH MAPPINGS ---
    If s = "personuid" Or s = "uid" Then MapFieldName = "PersonUID_Raw": Exit Function
    If s = "sourceid" Then MapFieldName = "SourceID_Raw": Exit Function
    If s = "fullname" Then MapFieldName = "FullName_Raw": Exit Function
    If s = "rank" Then MapFieldName = "Rank_Raw": Exit Function
    If s = "birthdate" Then MapFieldName = "BirthDate_Raw": Exit Function
    If s = "workstatus" Then MapFieldName = "WorkStatus_Raw": Exit Function
    If s = "poscode" Then MapFieldName = "PosCode_Raw": Exit Function
    If s = "posname" Then MapFieldName = "PosName_Raw": Exit Function
    If s = "orderdate" Then MapFieldName = "OrderDate_Raw": Exit Function
    If s = "ordernum" Then MapFieldName = "OrderNum_Raw": Exit Function
    
    ' --- RUSSIAN MAPPINGS (fuzzy by patterns) ---
    If ContainsCyrillic(strExcelField) Then
        ' SourceID: "Лицо" (not "Лицо1")
        If MatchCyrPattern(s, Array(1083, 1080, 1094, 1086)) And Len(s) < 6 Then
            MapFieldName = "SourceID_Raw": Exit Function
        End If
        ' FullName: "ФИО" or "Лицо1"
        If MatchCyrPattern(s, Array(1092, 1080, 1086)) Then
            MapFieldName = "FullName_Raw": Exit Function
        End If
        If MatchCyrPattern(s, Array(1083, 1080, 1094, 1086, 49)) Then  ' Лицо1
            MapFieldName = "FullName_Raw": Exit Function
        End If
        ' Rank: "звание"
        If MatchCyrPattern(s, Array(1079, 1074, 1072, 1085, 1080, 1077)) Then
            MapFieldName = "Rank_Raw": Exit Function
        End If
        ' BirthDate: "рождения"
        If MatchCyrPattern(s, Array(1088, 1086, 1078, 1076, 1077, 1085, 1080, 1103)) Then
            MapFieldName = "BirthDate_Raw": Exit Function
        End If
        ' WorkStatus: "статус" and "занятости"
        If MatchCyrPattern(s, Array(1089, 1090, 1072, 1090, 1091, 1089)) Then
            If MatchCyrPattern(s, Array(1079, 1072, 1085, 1103, 1090)) Then
                MapFieldName = "WorkStatus_Raw": Exit Function
            End If
        End If
        ' PosCode: "штатная"
        If MatchCyrPattern(s, Array(1096, 1090, 1072, 1090, 1085)) Then
            MapFieldName = "PosCode_Raw": Exit Function
        End If
        ' PosName: "должность"
        If MatchCyrPattern(s, Array(1076, 1086, 1083, 1078, 1085, 1086, 1089, 1090, 1100)) Then
            MapFieldName = "PosName_Raw": Exit Function
        End If
        ' OrderDate: "дата приказа"
        If MatchCyrPattern(s, Array(1076, 1072, 1090, 1072)) And MatchCyrPattern(s, Array(1087, 1088, 1080, 1082, 1072, 1079)) Then
            MapFieldName = "OrderDate_Raw": Exit Function
        End If
        ' OrderNum: "номер приказа"
        If MatchCyrPattern(s, Array(1085, 1086, 1084, 1077, 1088)) And MatchCyrPattern(s, Array(1087, 1088, 1080, 1082, 1072, 1079)) Then
            MapFieldName = "OrderNum_Raw": Exit Function
        End If
    End If
    
    ' Default: sanitize field name
    MapFieldName = mod_Schema_Manager.SanitizeFieldName(strExcelField)
End Function

' =============================================
' @description Checks if string contains Cyrillic pattern (Unicode codes)
' @param s [String] String to check (lowercased)
' @param pattern [Variant] Array of Unicode code points
' @return [Boolean] True if pattern found
' =============================================
Private Function MatchCyrPattern(s As String, pattern As Variant) As Boolean
    Dim patternStr As String
    Dim i As Long
    
    patternStr = ""
    For i = LBound(pattern) To UBound(pattern)
        patternStr = patternStr & ChrW(pattern(i))
    Next i
    
    MatchCyrPattern = (InStr(LCase(s), LCase(patternStr)) > 0)
End Function

' --- HELPER FUNCTIONS (Same as before, but simplified) ---

Private Function LinkExcelFile(strPath As String) As Boolean
    On Error GoTo ErrorHandler
    Dim db As DAO.Database, tdf As DAO.TableDef
    Dim strConnect As String, strSheet As String

    DeleteExcelLink
    Set db = CurrentDb

    ' Sheet reconnaissance
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
    MsgBox "Link error: " & Err.Description
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
    Dim fd As Object
    Set fd = Application.FileDialog(3)
    With fd
        .Filters.Clear: .Filters.Add "Excel", "*.xls;*.xlsx"
        If .Show = -1 Then SelectExcelFile = .SelectedItems(1)
    End With
End Function

Private Sub DeleteExcelLink()
    On Error Resume Next
    CurrentDb.TableDefs.Delete cstrLinkedTableName
End Sub

' =============================================
' @description Gets file modification date using FileSystemObject (Late Binding).
' @param strFilePath [String] Full path to file.
' @return [Date] File modification date, or Now() if error.
' =============================================
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
    ' Fallback to current date if file is not accessible
    GetFileModificationDate = Now()
End Function

' =============================================
' @description Updates import metadata table with file date and import timestamp.
'              Single-row strategy: DELETE all, then INSERT new.
' @param strFilePath [String] Full path to imported file.
' =============================================
Private Sub UpdateImportMetadata(strFilePath As String)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim dtFileDate As Date
    Dim strSQL As String

    Set db = CurrentDb

    ' Get file modification date
    dtFileDate = GetFileModificationDate(strFilePath)

    ' Clear old metadata (single-row table)
    db.Execute "DELETE FROM tbl_Import_Meta;", dbFailOnError

    ' Insert new metadata with explicit date format (slashes required by Access)
    Dim strFileDate As String, strNowDate As String
    strFileDate = Month(dtFileDate) & "/" & Day(dtFileDate) & "/" & Year(dtFileDate) & " " & _
                  Format(dtFileDate, "hh:nn:ss")
    strNowDate = Month(Now()) & "/" & Day(Now()) & "/" & Year(Now()) & " " & _
                 Format(Now(), "hh:nn:ss")

    strSQL = "INSERT INTO tbl_Import_Meta (ExportFileDate, ImportRunAt, SourceFilePath) " & _
             "VALUES (#" & strFileDate & "#, " & _
             "#" & strNowDate & "#, " & _
             "'" & Replace(strFilePath, "'", "''") & "');"

    db.Execute strSQL, dbFailOnError

    Debug.Print "Import metadata updated. ExportFileDate: " & dtFileDate

    Set db = Nothing
    Exit Sub

ErrorHandler:
    ' Non-critical error, just log
    Debug.Print "Warning: Failed to update import metadata: " & Err.Description
End Sub
