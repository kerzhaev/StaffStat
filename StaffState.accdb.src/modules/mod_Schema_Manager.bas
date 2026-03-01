Attribute VB_Name = "mod_Schema_Manager"
Option Compare Database
Option Explicit

' =============================================
' @module mod_Schema_Manager
' @description On-the-fly table structure and mapping management
' @note 100% English version. Safe for modern IDEs. Cyrillic headers are mapped via ASCII codes.
' =============================================

' =============================================
' @description Converts string to SQL-safe field name
' =============================================
Public Function SanitizeFieldName(ByVal strName As String) As String
    Dim s As String
    s = strName
    s = Replace(s, ".", "_")
    s = Replace(s, ",", "_")
    s = Replace(s, " ", "_")
    s = Replace(s, "-", "_")
    s = Replace(s, "(", "")
    s = Replace(s, ")", "")
    s = Replace(s, "/", "_")

    While InStr(s, "__") > 0
        s = Replace(s, "__", "_")
    Wend

    If Right(s, 1) = "_" Then s = Left(s, Len(s) - 1)
    If Left(s, 1) = "_" Then s = Mid(s, 2)

    SanitizeFieldName = s
End Function

' =============================================
' @description Checks if field exists in table. If not, creates it.
' =============================================
Public Sub EnsureFieldExists(strTableName As String, strFieldName As String, Optional strType As String = "TEXT(255)")
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim blnExists As Boolean

    Set db = CurrentDb
    Set tdf = db.TableDefs(strTableName)

    blnExists = False
    For Each fld In tdf.Fields
        If UCase(fld.Name) = UCase(strFieldName) Then
            blnExists = True
            Exit For
        End If
    Next fld

    If Not blnExists Then
        Dim strSQL As String
        strSQL = "ALTER TABLE [" & strTableName & "] ADD COLUMN [" & strFieldName & "] " & strType & ";"
        db.Execute strSQL, dbFailOnError
        Debug.Print "Schema: field added to " & strTableName & ": " & strFieldName
    End If

    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print "EnsureFieldExists error: " & Err.Description
    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
End Sub

' =============================================
' @description Creates tbl_Settings
' =============================================
Public Sub CreateSettingsTable()
    On Error GoTo ErrorHandler
    If TableExists("tbl_Settings") Then Exit Sub

    Dim db As DAO.Database
    Dim sSQL As String
    Set db = CurrentDb
    sSQL = "CREATE TABLE tbl_Settings (" & _
           "SettingKey TEXT(50) CONSTRAINT PK_Settings PRIMARY KEY, " & _
           "SettingValue TEXT(255), " & _
           "SettingGroup TEXT(50), " & _
           "Description TEXT(255));"
    db.Execute sSQL, dbFailOnError
    Debug.Print "Schema: table tbl_Settings created"
    Set db = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print "CreateSettingsTable error: " & Err.Description
    Set db = Nothing
End Sub

' =============================================
' @description Creates tbl_Validation_Log
' =============================================
Public Sub CreateValidationLogTable()
    On Error GoTo ErrorHandler
    If TableExists("tbl_Validation_Log") Then Exit Sub

    Dim db As DAO.Database
    Dim sSQL As String
    Set db = CurrentDb
    sSQL = "CREATE TABLE tbl_Validation_Log (" & _
           "LogID COUNTER CONSTRAINT PK_ValidationLog PRIMARY KEY, " & _
           "RecordID LONG, TableName TEXT(50), ErrorType TEXT(50), " & _
           "ErrorMessage TEXT(255), CheckDate DATETIME);"
    db.Execute sSQL, dbFailOnError
    Debug.Print "Schema: table tbl_Validation_Log created"
    Set db = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print "CreateValidationLogTable error: " & Err.Description
    Set db = Nothing
End Sub

' =============================================
' @description Returns True if table exists
' =============================================
Public Function TableExists(strTableName As String) As Boolean
    On Error Resume Next
    Dim tdf As DAO.TableDef
    Set tdf = CurrentDb.TableDefs(strTableName)
    TableExists = (Err.Number = 0)
    Err.Clear
    Set tdf = Nothing
End Function

' =============================================
' @description Returns True if field exists
' =============================================
Public Function FieldExists(strTable As String, strField As String) As Boolean
    On Error GoTo ErrorHandler
    Dim db As DAO.Database, tdf As DAO.TableDef, fld As DAO.Field
    FieldExists = False
    If Len(Trim(strTable)) = 0 Or Len(Trim(strField)) = 0 Then Exit Function
    Set db = CurrentDb
    Set tdf = db.TableDefs(strTable)
    For Each fld In tdf.Fields
        If UCase(fld.Name) = UCase(Trim(strField)) Then
            FieldExists = True
            Exit For
        End If
    Next fld
    Set fld = Nothing: Set tdf = Nothing: Set db = Nothing
    Exit Function
ErrorHandler:
    FieldExists = False
    Set fld = Nothing: Set tdf = Nothing: Set db = Nothing
End Function

' =============================================
' @description Creates Import Mapping Tables
' =============================================
Public Sub CreateImportProfilesTable()
    On Error GoTo ErrorHandler
    If TableExists("tbl_Import_Profiles") Then Exit Sub
    CurrentDb.Execute "CREATE TABLE [tbl_Import_Profiles] ([ProfileID] LONG CONSTRAINT [PK_Import_Profiles] PRIMARY KEY, [ProfileName] TEXT(100), [IdStrategy] TEXT(20));", dbFailOnError
    CurrentDb.Execute "INSERT INTO tbl_Import_Profiles (ProfileID, ProfileName, IdStrategy) VALUES (1, 'Default', 'PersonUID');", dbFailOnError
    Exit Sub
ErrorHandler:
    Debug.Print "CreateImportProfilesTable error: " & Err.Description
End Sub

Public Sub CreateImportMappingTable()
    On Error GoTo ErrorHandler
    If TableExists("tbl_Import_Mapping") Then Exit Sub
    CurrentDb.Execute "CREATE TABLE [tbl_Import_Mapping] ([MappingID] COUNTER CONSTRAINT [PK_Import_Mapping] PRIMARY KEY, [ProfileID] LONG NOT NULL, [ExcelHeader] TEXT(255) NOT NULL, [TargetField] TEXT(100) NOT NULL);", dbFailOnError
    Exit Sub
ErrorHandler:
    Debug.Print "CreateImportMappingTable error: " & Err.Description
End Sub

' =============================================
' @description Seeds Profile 1 with ASCII-encoded Cyrillic headers
' =============================================
Public Sub SeedImportMappingProfile1()
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Set db = CurrentDb

    If Not TableExists("tbl_Import_Mapping") Then
        CreateImportMappingTable
        Set db = CurrentDb
    End If

    ' Personal Data
    AddMappingIfNotExists db, 1, CyrStr(1051, 1083, 1094, 1086), "SourceID"
    AddMappingIfNotExists db, 1, CyrStr(1051, 1080, 1095, 1085, 1099, 1081, 32, 1085, 1086, 1084, 1077, 1088), "PersonUID"
    AddMappingIfNotExists db, 1, CyrStr(1042, 1086, 1080, 1085, 1089, 1082, 1086, 1077, 32, 1079, 1074, 1072, 1085, 1080, 1077), "RankName"
    AddMappingIfNotExists db, 1, CyrStr(1060, 1048, 1054), "FullName"
    AddMappingIfNotExists db, 1, CyrStr(1044, 1072, 1090, 1072, 32, 1088, 1086, 1078, 1076, 1077, 1085, 1080, 1103), "BirthDate"
    AddMappingIfNotExists db, 1, CyrStr(1042, 1086, 1079, 1088, 1072, 1089, 1090, 32, 1089, 1086, 1090, 1088, 1091, 1076, 1085, 1080, 1082, 1072), "EmployeeAge"
    AddMappingIfNotExists db, 1, CyrStr(1055, 1086, 1083), "Gender"
    AddMappingIfNotExists db, 1, CyrStr(1057, 1077, 1084, 1077, 1081, 1085, 1086, 1077, 32, 1087, 1086, 1083, 1086, 1078, 1077, 1085, 1080, 1077), "MaritalStatus"
    AddMappingIfNotExists db, 1, CyrStr(1050, 1086, 1083, 1080, 1095, 1077, 1089, 1090, 1074, 1086, 32, 1076, 1077, 1090, 1077, 1081), "ChildrenCount"
    AddMappingIfNotExists db, 1, CyrStr(1053, 1072, 1094, 1080, 1086, 1085, 1072, 1083, 1100, 1085, 1086, 1089, 1090, 1100), "Nationality"
    AddMappingIfNotExists db, 1, CyrStr(1043, 1088, 1072, 1078, 1076, 1072, 1085, 1089, 1090, 1074, 1086), "Citizenship"
    AddMappingIfNotExists db, 1, CyrStr(1044, 1072, 1090, 1072, 32, 1091, 1074, 1086, 1083, 1100, 1085, 1077, 1085, 1080, 1103), "DismissalDate"
    AddMappingIfNotExists db, 1, CyrStr(1043, 1088, 1091, 1087, 1087, 1072, 32, 1089, 1086, 1090, 1088, 1091, 1076, 1085, 1080, 1082, 1086, 1074), "EmployeeGroup"

    ' Contract
    AddMappingIfNotExists db, 1, CyrStr(1042, 1080, 1076, 32, 1082, 1086, 1085, 1090, 1088, 1072, 1082, 1090, 1072), "ContractKind"
    AddMappingIfNotExists db, 1, CyrStr(1058, 1080, 1087, 32, 1082, 1086, 1085, 1090, 1088, 1072, 1082, 1090, 1072), "ContractType"
    AddMappingIfNotExists db, 1, CyrStr(1044, 1072, 1090, 1072, 32, 1085, 1072, 1095, 1072, 1083, 1072, 32, 1082, 1086, 1085, 1090, 1088, 1072, 1082, 1090, 1072), "ContractStartDate"
    AddMappingIfNotExists db, 1, CyrStr(1044, 1072, 1090, 1072, 32, 1086, 1082, 1086, 1085, 1095, 1072, 1085, 1080, 1103, 32, 1082, 1086, 1085, 1090, 1088, 1072, 1082, 1090, 1072), "ContractEndDate"
    AddMappingIfNotExists db, 1, CyrStr(1057, 1088, 1086, 1082, 32, 1076, 1086, 1075, 1086, 1074, 1072, 32, 1075, 1086, 1076), "ContractYears"
    AddMappingIfNotExists db, 1, CyrStr(1057, 1088, 1086, 1082, 32, 1076, 1086, 1075, 1086, 1074, 1072, 32, 1084, 1077, 1089, 1103, 1094), "ContractMonths"

    ' Order
    AddMappingIfNotExists db, 1, CyrStr(1053, 1086, 1084, 1077, 1088, 32, 1087, 1088, 1080, 1082, 1072, 1079, 1072), "OrderNumber"
    AddMappingIfNotExists db, 1, CyrStr(1044, 1072, 1090, 1072, 32, 1087, 1088, 1080, 1082, 1072, 1079, 1072), "OrderDate_Text"
    AddMappingIfNotExists db, 1, CyrStr(1042, 1080, 1076, 32, 1084, 1077, 1088, 1086, 1087, 1088, 1080, 1103, 1090, 1080, 1103), "EventType"
    AddMappingIfNotExists db, 1, CyrStr(1055, 1088, 1080, 1095, 1080, 1085, 1072, 32, 1084, 1077, 1088, 1086, 1087, 1088, 1080, 1103, 1090, 1080, 1103), "EventReason"
    AddMappingIfNotExists db, 1, CyrStr(1053, 1072, 1095, 1072, 1083, 1086, 32, 1089, 1088, 1086, 1082, 1072, 32, 1076, 1077, 1081, 1090, 1074, 1080, 1103), "ValidFromDate"
    AddMappingIfNotExists db, 1, CyrStr(1050, 1086, 1085, 1077, 1094, 32, 1089, 1088, 1086, 1082, 1072, 32, 1076, 1077, 1081, 1090, 1074, 1080, 1103), "ValidToDate"

    ' Position
    AddMappingIfNotExists db, 1, CyrStr(1064, 1090, 1072, 1090, 1085, 1072, 1103, 32, 1076, 1086, 1083, 1078, 1085, 1086, 1089, 1090, 1100), "StaffPosition"
    AddMappingIfNotExists db, 1, CyrStr(1044, 1086, 1083, 1078, 1085, 1086, 1089, 1090, 1100), "Position"
    AddMappingIfNotExists db, 1, CyrStr(1042, 1059, 1057), "VUS"
    AddMappingIfNotExists db, 1, CyrStr(1058, 1072, 1088, 1080, 1092, 1085, 1099, 1081, 32, 1088, 1072, 1079, 1088, 1103, 1076), "SalaryGrade"
    AddMappingIfNotExists db, 1, CyrStr(1044, 1072, 1090, 1072, 32, 1087, 1088, 1080, 1082, 1072, 32, 1051, 1057, 1057), "OrderDate_LS"
    AddMappingIfNotExists db, 1, CyrStr(1063, 1077, 1081, 32, 1087, 1088, 1080, 1082, 1072, 1079, 32, 1076, 1086, 1083, 1078, 1085, 1086, 1089, 1090, 1100), "PositionOrderIssuer"
    AddMappingIfNotExists db, 1, CyrStr(1056, 1072, 1079, 1076, 1077, 1083, 32, 1087, 1077, 1088, 1091, 1086, 1085, 1072, 1083, 1072), "PersonnelDivision"
    AddMappingIfNotExists db, 1, CyrStr(1057, 1090, 1072, 1090, 1091, 1089, 32, 1079, 1072, 1085, 1103, 1090, 1086, 1089, 1090, 1080), "EmploymentStatus"

    ' Banking
    AddMappingIfNotExists db, 1, CyrStr(1050, 1086, 1085, 1090, 1088, 1086, 1083, 1100, 1085, 1099, 1081, 32, 1073, 1072, 1085, 1082, 1086, 1074, 1089, 1082, 1080, 1081, 32, 1082, 1083, 1102, 1095), "BankControlKey"
    AddMappingIfNotExists db, 1, CyrStr(1053, 1086, 1084, 1077, 1088, 32, 1089, 1095, 1077, 1090, 1072, 32, 1074, 32, 1073, 1072, 1085, 1082, 1077), "BankAccountNumber"
    AddMappingIfNotExists db, 1, CyrStr(1055, 1086, 1083, 1091, 1095, 1072, 1090, 1077, 1083, 1100), "Payee"
    AddMappingIfNotExists db, 1, CyrStr(1050, 1083, 1102, 1095, 32, 1073, 1072, 1085, 1082, 1072), "BankKey"

    ' Sizes
    AddMappingIfNotExists db, 1, CyrStr(1056, 1072, 1079, 1084, 1077, 1088, 32, 1057, 1072, 1087, 1086, 1075), "BootSize"
    AddMappingIfNotExists db, 1, CyrStr(1054, 1093, 1074, 1072, 1090, 32, 1075, 1086, 1083, 1086, 1074, 1099), "HeadSize"

    Set db = Nothing
    Debug.Print "SeedImportMappingProfile1: Profile 1 seeded."
    Exit Sub
ErrorHandler:
    Debug.Print "SeedImportMappingProfile1 error: " & Err.Description
    Set db = Nothing
End Sub

Public Sub ReSeedMapping()
    On Error GoTo ErrorHandler
    CurrentDb.Execute "DELETE FROM tbl_Import_Mapping WHERE ProfileID = 1", dbFailOnError
    SeedImportMappingProfile1
    Exit Sub
ErrorHandler:
    Debug.Print "ReSeedMapping error: " & Err.Description
End Sub

' =============================================
' @description Adds mapping row using Parameterized QueryDef
' =============================================
Private Sub AddMapping(db As DAO.Database, lngProfile As Long, strExcel As String, strTarget As String)
    Dim qdf As DAO.QueryDef
    Dim strSQL As String

    strSQL = "PARAMETERS prmProfile Long, prmExcel Text (255), prmTarget Text (100); " & _
             "INSERT INTO tbl_Import_Mapping (ProfileID, ExcelHeader, TargetField) " & _
             "VALUES ([prmProfile], [prmExcel], [prmTarget]);"

    Set qdf = db.CreateQueryDef("", strSQL)
    qdf.Parameters("prmProfile").value = lngProfile
    qdf.Parameters("prmExcel").value = Left$(strExcel, 255)
    qdf.Parameters("prmTarget").value = Left$(strTarget, 100)

    qdf.Execute dbFailOnError
    Set qdf = Nothing
End Sub

Private Sub AddMappingIfNotExists(db As DAO.Database, lngProfile As Long, strExcel As String, strTarget As String)
    Dim strSQL As String
    Dim rs As DAO.Recordset
    strSQL = "SELECT MappingID FROM tbl_Import_Mapping WHERE ProfileID = " & lngProfile & " AND ExcelHeader = '" & Replace(strExcel, "'", "''") & "'"
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If rs.EOF Then AddMapping db, lngProfile, strExcel, strTarget
    rs.Close
    Set rs = Nothing
End Sub

' =============================================
' @description Helper to build string from ASCII codes
' =============================================
Private Function CyrStr(ParamArray codes() As Variant) As String
    Dim i As Long
    Dim s As String
    s = ""
    For i = LBound(codes) To UBound(codes)
        s = s & ChrW(CLng(codes(i)))
    Next i
    CyrStr = s
End Function



' =============================================
' @description Ensures tbl_Personnel_Master has all canonical English fields.
' =============================================
Public Sub SyncMasterStructure()
    Dim colAllowed As Collection
    Dim v As Variant
    Dim strType As String

    If Not TableExists("tbl_Personnel_Master") Then Exit Sub

    Set colAllowed = GetAllowedMasterFields()
    For Each v In colAllowed
        strType = "TEXT(255)"
        If v = "ContractYears" Or v = "ContractMonths" Or v = "SourceID" Then strType = "LONG"
        If v = "PosName" Then strType = "LONGTEXT"
        If v = "LastUpdated" Then strType = "DATETIME"
        If v = "IsActive" Then strType = "BIT"
        EnsureFieldExists "tbl_Personnel_Master", CStr(v), strType
    Next v
    Set colAllowed = Nothing
End Sub

' =============================================
' @description Canonical English fields for Master.
' =============================================
Private Function GetAllowedMasterFields() As Collection
    Dim c As Collection
    Set c = New Collection
    Dim arr As Variant
    Dim i As Long

    arr = Array("PersonUID", "SourceID", "FullName", "RankName", "BirthDate_Text", "WorkStatus", "PosCode", "PosName", "OrderDate_Text", "OrderNumber", "EmployeeAge", "Gender", "MaritalStatus", "ChildrenCount", "Nationality", "Citizenship", "ContractType", "ContractKind", "ContractStartDate", "ContractEndDate", "ContractYears", "ContractMonths", "EventType", "EventReason", "ValidFromDate", "ValidToDate", "StaffPosition", "Position", "VUS", "SalaryGrade", "PersonnelDivision", "BankAccountNumber", "Payee", "BankKey", "BootSize", "HeadSize", "LastUpdated", "IsActive")

    For i = LBound(arr) To UBound(arr)
        On Error Resume Next
        c.Add arr(i), UCase(CStr(arr(i)))
        On Error GoTo 0
    Next i
    Set GetAllowedMasterFields = c
End Function


' =============================================
' @description Gets field type as friendly string (for UI)
' =============================================
Public Function GetFieldTypeFriendly(ByVal strTable As String, ByVal strField As String) As String
    On Error Resume Next
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field

    Set db = CurrentDb
    ' CRITICAL: Force Access to clear cache and look at the real linked table
    db.TableDefs.Refresh

    Set tdf = db.TableDefs(strTable)
    Set fld = tdf.Fields(strField)

    If Err.Number <> 0 Then
        GetFieldTypeFriendly = "Unknown"
        Exit Function
    End If

    Select Case fld.Type
        Case 10: GetFieldTypeFriendly = "Text (255)"
        Case 12: GetFieldTypeFriendly = "Long Text"
        Case 8:  GetFieldTypeFriendly = "Date/Time"
        Case 4:  GetFieldTypeFriendly = "Number"
        Case 1:  GetFieldTypeFriendly = "Yes/No"
        Case Else: GetFieldTypeFriendly = "Type " & fld.Type
    End Select

    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
End Function



' =============================================
' @description Helper: Gets physical file path of a table if linked.
' =============================================
Public Function GetBackendPath(ByVal strTableName As String) As String
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strConnect As String
    Dim iPos As Long

    GetBackendPath = ""
    On Error Resume Next
    Set db = CurrentDb
    db.TableDefs.Refresh
    Set tdf = db.TableDefs(strTableName)

    If Err.Number = 0 Then
        strConnect = tdf.Connect
        If Len(strConnect) > 0 Then
            ' CRITICAL FIX: Case-insensitive search (vbTextCompare) for "DATABASE="
            iPos = InStr(1, strConnect, "DATABASE=", vbTextCompare)
            If iPos > 0 Then
                GetBackendPath = Trim$(Mid$(strConnect, iPos + 9))
            End If
        End If
    End If

    Set tdf = Nothing
    Set db = Nothing
    Err.Clear
End Function

' =============================================
' @description Alters existing column type in Master and Buffer. Adds column if it is missing.
' =============================================
Public Sub AlterFieldType(ByVal strFieldName As String, ByVal strDataType As String)
    Dim dbLocal As DAO.Database
    Dim dbBackend As DAO.Database
    Dim strPath As String
    Dim strSafe As String
    Dim blnLocal As Boolean
    Dim errNum As Long
    Dim errDesc As String
    Dim blnExists As Boolean
    Dim tdfTemp As DAO.TableDef
    Dim fldTemp As DAO.Field

    On Error GoTo ErrorHandler

    strSafe = Trim$(strFieldName)
    If Len(strSafe) = 0 Then Exit Sub

    ' Get path to the Back-End file
    strPath = GetBackendPath("tbl_Personnel_Master")

    If Len(strPath) > 0 Then
        ' Connected to Split Database Back-End
        Set dbBackend = DBEngine.Workspaces(0).OpenDatabase(strPath)
        blnLocal = False
    Else
        ' Safety trigger: If it thinks it's local but is actually linked, STOP!
        If Len(CurrentDb.TableDefs("tbl_Personnel_Master").Connect) > 0 Then
            Err.Raise vbObjectError + 1, "AlterFieldType", "Failed to resolve Back-End path. Connection string: " & CurrentDb.TableDefs("tbl_Personnel_Master").Connect
        End If
        Set dbLocal = CurrentDb
        Set dbBackend = dbLocal
        blnLocal = True
    End If

    ' --- 1. Check and Update MASTER TABLE ---
    blnExists = False
    dbBackend.TableDefs.Refresh
    Set tdfTemp = dbBackend.TableDefs("tbl_Personnel_Master")
    For Each fldTemp In tdfTemp.Fields
        If UCase(fldTemp.Name) = UCase(strSafe) Then
            blnExists = True
            Exit For
        End If
    Next fldTemp

    If blnExists Then
        dbBackend.Execute "ALTER TABLE [tbl_Personnel_Master] ALTER COLUMN [" & strSafe & "] " & strDataType & ";", dbFailOnError
    Else
        dbBackend.Execute "ALTER TABLE [tbl_Personnel_Master] ADD COLUMN [" & strSafe & "] " & strDataType & ";", dbFailOnError
    End If

    ' --- 2. Check and Update BUFFER TABLE ---
    blnExists = False
    Set tdfTemp = dbBackend.TableDefs("tbl_Import_Buffer")
    For Each fldTemp In tdfTemp.Fields
        If UCase(fldTemp.Name) = UCase(strSafe) Then
            blnExists = True
            Exit For
        End If
    Next fldTemp

    If blnExists Then
        dbBackend.Execute "ALTER TABLE [tbl_Import_Buffer] ALTER COLUMN [" & strSafe & "] " & strDataType & ";", dbFailOnError
    Else
        dbBackend.Execute "ALTER TABLE [tbl_Import_Buffer] ADD COLUMN [" & strSafe & "] " & strDataType & ";", dbFailOnError
    End If

    ' --- 3. Refresh Links ---
    If Not blnLocal Then
        dbBackend.Close
        Set dbBackend = Nothing

        Set dbLocal = CurrentDb
        dbLocal.TableDefs.Refresh
        dbLocal.TableDefs("tbl_Personnel_Master").RefreshLink
        dbLocal.TableDefs("tbl_Import_Buffer").RefreshLink
    End If

    Set tdfTemp = Nothing
    Set fldTemp = Nothing
    If Not dbLocal Is Nothing Then Set dbLocal = Nothing
    Exit Sub

ErrorHandler:
    errNum = Err.Number
    errDesc = Err.Description
    If Not dbBackend Is Nothing And Not blnLocal Then
        On Error Resume Next
        dbBackend.Close
    End If
    Set dbBackend = Nothing
    Set dbLocal = Nothing
    Err.Raise errNum, "AlterFieldType", errDesc
End Sub

' =============================================
' @description Wrapper for adding field (Safe for both new and existing fields)
' =============================================
Public Sub AddNewFieldToSchema(ByVal strFieldName As String, Optional ByVal strDataType As String = "VARCHAR(255)")
    AlterFieldType strFieldName, strDataType
End Sub
