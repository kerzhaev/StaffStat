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
' @description Creates Import Mapping Tables & Multiple Profiles
' =============================================
Public Sub CreateImportProfilesTable()
    On Error GoTo ErrorHandler
    If Not TableExists("tbl_Import_Profiles") Then
        CurrentDb.Execute "CREATE TABLE [tbl_Import_Profiles] ([ProfileID] LONG CONSTRAINT [PK_Import_Profiles] PRIMARY KEY, [ProfileName] TEXT(100), [IdStrategy] TEXT(20));", dbFailOnError
    End If

    ' Phase 31: Ensure default profiles exist
    On Error Resume Next
    CurrentDb.Execute "INSERT INTO tbl_Import_Profiles (ProfileID, ProfileName, IdStrategy) VALUES (1, 'Main (Default)', 'PersonUID');"
    CurrentDb.Execute "INSERT INTO tbl_Import_Profiles (ProfileID, ProfileName, IdStrategy) VALUES (2, 'Logistics/Supply', 'PersonUID');"
    CurrentDb.Execute "INSERT INTO tbl_Import_Profiles (ProfileID, ProfileName, IdStrategy) VALUES (3, 'Finance/Banking', 'PersonUID');"
    On Error GoTo ErrorHandler
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
    AddMappingIfNotExists db, 1, CyrStr(1051, 1048, 1094, 1086), "SourceID"
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

' =============================================
' @author
' @description Creates the tbl_Localization table for UI translations
' =============================================
Public Sub CreateLocalizationTable()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    ' Check if table already exists
    Dim td As DAO.TableDef
    Dim bExists As Boolean
    bExists = False

    For Each td In db.TableDefs
        If td.Name = "tbl_Localization" Then
            bExists = True
            Exit For
        End If
    Next td

    If bExists Then
        Debug.Print "Table tbl_Localization already exists. Skipping creation."
        Exit Sub
    End If

    ' Create table
    Dim strSQL As String
    strSQL = "CREATE TABLE tbl_Localization (" & _
             "MsgKey VARCHAR(100) CONSTRAINT PK_Localization PRIMARY KEY, " & _
             "LocalValue LONGTEXT, " & _
             "Category VARCHAR(50));"

    db.Execute strSQL, dbFailOnError
    Debug.Print "Table tbl_Localization created successfully."

    ' Seed default values (English placeholders to protect encoding)
    ' The user will translate these to Russian directly in Access UI
    db.Execute "INSERT INTO tbl_Localization (MsgKey, LocalValue, Category) VALUES ('BTN_SEARCH', 'Search', 'UI')", dbFailOnError
    db.Execute "INSERT INTO tbl_Localization (MsgKey, LocalValue, Category) VALUES ('BTN_CLEAR', 'Clear', 'UI')", dbFailOnError
    db.Execute "INSERT INTO tbl_Localization (MsgKey, LocalValue, Category) VALUES ('BTN_EXPORT', 'Export to Excel', 'UI')", dbFailOnError
    db.Execute "INSERT INTO tbl_Localization (MsgKey, LocalValue, Category) VALUES ('MSG_NO_DUPLICATES', 'No duplicates found.', 'System')", dbFailOnError
    db.Execute "INSERT INTO tbl_Localization (MsgKey, LocalValue, Category) VALUES ('ERR_INVALID_UID', 'Invalid PersonUID format.', 'Error')", dbFailOnError

    Exit Sub

ErrorHandler:
    Debug.Print "CreateLocalizationTable Error " & Err.Number & ": " & Err.Description
End Sub

' =============================================
' @description Helper to safely Insert or Update localization keys
' =============================================
Private Sub UpsertLocKey(db As DAO.Database, strKey As String, strValue As String, strCat As String)
    Dim rs As DAO.Recordset
    ' Ищем существующий ключ
    Set rs = db.OpenRecordset("SELECT * FROM tbl_Localization WHERE MsgKey = '" & strKey & "'", dbOpenDynaset)

    If rs.EOF Then
        rs.AddNew       ' Если нет - добавляем
        rs!MsgKey = strKey
    Else
        rs.Edit         ' Если есть - обновляем текст
    End If

    rs!LocalValue = strValue
    rs!Category = strCat
    rs.Update

    rs.Close
    Set rs = Nothing
End Sub

' =============================================
' @description Seeds tbl_Localization with default Russian UI text
' =============================================
Public Sub SeedLocalizationTable()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    Debug.Print "Starting localization seeding..."

    ' --- Общие элементы UI ---
    UpsertLocKey db, "BTN_SEARCH", "Поиск", "UI"
    UpsertLocKey db, "BTN_CLEAR", "Очистить", "UI"
    UpsertLocKey db, "BTN_EXPORT", "Экспорт в Excel", "UI"
    UpsertLocKey db, "BTN_SAVE", "Сохранить", "UI"
    UpsertLocKey db, "BTN_CANCEL", "Отмена", "UI"

    ' --- Заголовки (Titles) ---
    UpsertLocKey db, "TITLE_INFO", "Информация", "UI"
    UpsertLocKey db, "TITLE_ERROR", "Ошибка", "UI"
    UpsertLocKey db, "TITLE_WARNING", "Предупреждение", "UI"
    UpsertLocKey db, "TITLE_SEARCH", "Поиск", "UI"

    ' --- Сообщения (Messages) ---
    UpsertLocKey db, "MSG_NO_DUPLICATES", "Дубликаты не найдены.", "System"
    UpsertLocKey db, "MSG_DUPLICATES_FOUND", "Найдены дубликаты.", "System"
    UpsertLocKey db, "MSG_SEARCH_PROMPT", "Введите от 2 символов для поиска...", "System"

    ' --- Ошибки (Errors) ---
    UpsertLocKey db, "ERR_INVALID_UID", "Неверный формат личного номера.", "Error"
    UpsertLocKey db, "ERR_DB_LOCKED", "База данных заблокирована другим пользователем.", "Error"

    ' --- Карточка сотрудника (Person Card) ---
    UpsertLocKey db, "LBL_INACTIVE_EMPLOYEE", "СОТРУДНИК НЕАКТИВЕН", "UI"
    UpsertLocKey db, "TAB_SERVICE", "Служба", "UI"
    UpsertLocKey db, "TAB_CONTRACT", "Контракт", "UI"
    UpsertLocKey db, "TAB_PERSONAL", "Личные данные", "UI"
    UpsertLocKey db, "TAB_LOGISTICS", "Снабжение", "UI"
    UpsertLocKey db, "TAB_BANK", "Банк", "UI"
    UpsertLocKey db, "TAB_HISTORY", "Машина времени", "UI"

    ' --- Форма uf_Search ---
    UpsertLocKey db, "TITLE_DUPLICATES", "Дубликаты (ФИО + Дата рождения)", "UI"
    UpsertLocKey db, "LBL_GROUPS", "групп", "UI"
    UpsertLocKey db, "LBL_RECORDS", "записей", "UI"
    UpsertLocKey db, "ERR_EMP_ID_NOT_READ", "Внимание: ID сотрудника не прочитан из списка. Проверьте 'Column Count' (должно быть 4) и 'Bound Column' (должно быть 1).", "Error"
    UpsertLocKey db, "MSG_EXPORT_EMPTY", "Нет данных для экспорта: список пуст.", "System"
    UpsertLocKey db, "ERR_EXPORT_PREP", "Ошибка при подготовке данных к экспорту.", "Error"
    UpsertLocKey db, "MSG_EXPORT_SUCCESS", "Экспорт успешно завершен:", "System"
    UpsertLocKey db, "MSG_NO_EMPLOYEES", "Сотрудники не найдены.", "System"

    ' --- Форма uf_Settings ---
    UpsertLocKey db, "TITLE_SETTINGS", "Настройки системы", "UI"
    UpsertLocKey db, "TAB_GENERAL", "Основные настройки", "UI"
    UpsertLocKey db, "TAB_MAPPING", "Маппинг импорта", "UI"
    UpsertLocKey db, "TAB_MAINTENANCE", "Обслуживание", "UI"
    UpsertLocKey db, "MSG_RESEED_ONLY_MAIN", "Восстановление по умолчанию доступно только для основного профиля (1).", "System"
    UpsertLocKey db, "LBL_EXCEL_HEADER", "Заголовок Excel:", "UI"
    UpsertLocKey db, "LBL_DB_FIELD", "Поле в базе:", "UI"

    UpsertLocKey db, "BTN_ADD_MAPPING", "Добавить связь", "UI"
    UpsertLocKey db, "BTN_DEL_MAPPING", "Удалить связь", "UI"
    UpsertLocKey db, "BTN_RESEED_MAPPING", "Восстановить по умолчанию", "UI"
    UpsertLocKey db, "BTN_CREATE_BACKUP", "Создать резервную копию", "UI"
    UpsertLocKey db, "BTN_CLEAR_LOGS", "Очистить журнал проверки", "UI"
    UpsertLocKey db, "BTN_RUN_HEALTH_CHECK", "Запустить проверку данных", "UI"

    UpsertLocKey db, "MSG_SETTINGS_SAVED", "Настройки успешно сохранены.", "System"
    UpsertLocKey db, "ERR_SAVE_FAILED", "Ошибка сохранения:", "Error"

    UpsertLocKey db, "MSG_FILL_MAPPING_FIELDS", "Пожалуйста, заполните Excel Header и Target Field.", "System"
    UpsertLocKey db, "TITLE_SCHEMA_MANAGER", "Менеджер схемы БД", "UI"
    UpsertLocKey db, "MSG_FIELD_MISSING", "Поле не существует в базе данных:", "System"
    UpsertLocKey db, "PROMPT_SELECT_DATA_TYPE", "Выберите тип данных (введите цифру):" & vbCrLf & "1 - Текст (255 симв.) [По умолчанию]" & vbCrLf & "2 - Дата / Время" & vbCrLf & "3 - Число (Длинное целое)" & vbCrLf & "4 - Длинный текст (Мемо)" & vbCrLf & "5 - Да/Нет (Логическое)", "UI"
    UpsertLocKey db, "MSG_MAPPING_CANCELED", "Создание маппинга отменено.", "System"
    UpsertLocKey db, "MSG_FIELD_CREATED", "Поле успешно создано с типом", "System"
    UpsertLocKey db, "ERR_ADD_MAPPING_FAILED", "Ошибка добавления маппинга:", "Error"

    UpsertLocKey db, "MSG_SELECT_ROW_DELETE", "Пожалуйста, выберите строку для удаления.", "System"
    UpsertLocKey db, "TITLE_CONFIRM", "Подтверждение", "UI"
    UpsertLocKey db, "PROMPT_DEL_PERSONUID", "Удаление маппинга PersonUID может сломать процесс импорта. Продолжить?", "System"
    UpsertLocKey db, "ERR_DELETE_FAILED", "Ошибка удаления:", "Error"

    UpsertLocKey db, "PROMPT_RESEED_MAPPING", "Восстановить маппинг по умолчанию? Все текущие пользовательские связи будут удалены.", "System"
    UpsertLocKey db, "MSG_MAPPING_RESTORED", "Маппинг восстановлен по умолчанию.", "System"
    UpsertLocKey db, "ERR_RESEED_FAILED", "Ошибка восстановления (ReSeed):", "Error"

    UpsertLocKey db, "TITLE_CHANGE_TYPE", "Изменение типа данных", "UI"
    UpsertLocKey db, "PROMPT_CHANGE_TYPE_WARN", "Вы хотите изменить тип данных для поля", "System"
    UpsertLocKey db, "PROMPT_CHANGE_TYPE_WARN2", "ВНИМАНИЕ: Изменение типа может привести к очистке существующих данных в этой колонке.", "System"
    UpsertLocKey db, "PROMPT_CHANGE_TYPE_INPUT", "Изменение типа данных для поля", "System"
    UpsertLocKey db, "MSG_FIELD_TYPE_CHANGED", "Тип поля успешно изменен на", "System"
    UpsertLocKey db, "ERR_CHANGE_TYPE_FAILED", "Ошибка изменения типа:", "Error"

    ' --- Форма uf_PersonCard ---
    UpsertLocKey db, "TITLE_PERSON_CARD", "Карточка сотрудника", "UI"
    UpsertLocKey db, "ERR_EMP_NOT_FOUND", "Сотрудник не найден. Личный номер:", "Error"
    UpsertLocKey db, "STATUS_DISMISSED", "Уволен", "System" ' Заменяет хардкод ChrW
    UpsertLocKey db, "FILTER_ALL", "Все", "UI"
    UpsertLocKey db, "BTN_RESET", "Сброс", "UI"

    ' --- Форма uf_Dashboard ---
    UpsertLocKey db, "TITLE_DASHBOARD", "Панель управления (Дозор)", "UI"
    UpsertLocKey db, "LBL_TOTAL", "Всего: ", "UI"
    UpsertLocKey db, "LBL_ACTIVE", "Активных: ", "UI"
    UpsertLocKey db, "LBL_ERRORS", "Ошибок: ", "UI"
    UpsertLocKey db, "ERR_UI", "Ошибка интерфейса: ", "Error"

    ' --- Обновленные ключи для Dashboard ---
    UpsertLocKey db, "BTN_HEALTH_CHECK", "Проверка данных", "UI"
    UpsertLocKey db, "BTN_OPEN_LOG", "Открыть журнал", "UI"
    UpsertLocKey db, "BTN_CHANGE_REPORT", "Отчет об изменениях", "UI"

    ' Новые ключи для рамок (Frames) и подписей
    UpsertLocKey db, "LBL_MANUAL_CONTROLS", "Ручное управление", "UI"

    ' Сообщения статуса
    UpsertLocKey db, "STATUS_ERROR", "Ошибка.", "UI"
    UpsertLocKey db, "MSG_RUNNING_IMPORT", "Выполнение импорта...", "System"
    UpsertLocKey db, "MSG_IMPORT_SUCCESS", "Импорт успешно завершен.", "System"
    UpsertLocKey db, "MSG_IMPORT_FAILED", "Импорт отменен или завершен с ошибкой.", "Error"

    UpsertLocKey db, "MSG_RUNNING_HEALTH", "Выполнение проверки целостности...", "System"
    UpsertLocKey db, "PROMPT_EXPORT_ERRORS", "Экспортировать отчет об ошибках в Excel?", "System"
    UpsertLocKey db, "TITLE_HEALTH_CHECK", "Проверка данных", "UI"
    UpsertLocKey db, "MSG_HEALTH_DONE", "Проверка завершена. Ошибок: ", "System"

    UpsertLocKey db, "MSG_RUNNING_ANALYSIS", "Выполнение анализа изменений...", "System"
    UpsertLocKey db, "MSG_ANALYSIS_DONE", "Анализ завершен. Новых: ", "System"
    UpsertLocKey db, "LBL_UPDATED", ", Обновлено: ", "System"

    UpsertLocKey db, "MSG_OPENING_LOG", "Открытие журнала...", "System"
    UpsertLocKey db, "MSG_LOG_OPENED", "Журнал открыт.", "System"

    UpsertLocKey db, "MSG_CREATING_INDEXES", "Создание индексов производительности...", "System"
    UpsertLocKey db, "MSG_INDEXES_DONE", "Индексы созданы. Добавлено: ", "System"
    UpsertLocKey db, "LBL_SKIPPED", ", Пропущено: ", "System"

    UpsertLocKey db, "MSG_RUNNING_FULL_SYNC", "Выполнение полного цикла обновления...", "System"
    UpsertLocKey db, "MSG_FULL_SYNC_DONE", "Полное обновление завершено.", "System"

    UpsertLocKey db, "MSG_OPENING_DUPLICATES", "Открытие инструмента поиска дубликатов...", "System"
    UpsertLocKey db, "MSG_DUPLICATES_OPENED", "Поиск дубликатов запущен.", "System"
    UpsertLocKey db, "ERR_OPEN_DUPLICATES", "Не удалось открыть поиск дубликатов: ", "Error"
    UpsertLocKey db, "MSG_SKIPPED_COLS", "Следующие колонки были пропущены (нет маппинга):", "System"

    ' Отчеты
    UpsertLocKey db, "ERR_ENTER_DATES", "Пожалуйста, введите начальную и конечную даты.", "Error"
    UpsertLocKey db, "ERR_INVALID_DATE_FORMAT", "Неверный формат даты. Используйте DD.MM.YYYY.", "Error"
    UpsertLocKey db, "ERR_START_AFTER_END", "Начальная дата должна быть меньше или равна конечной.", "Error"
    UpsertLocKey db, "MSG_GEN_AUDIT", "Генерация аудит-отчета...", "System"
    UpsertLocKey db, "MSG_AUDIT_DONE", "Аудит-отчет готов.", "System"
    UpsertLocKey db, "MSG_GEN_SNAPSHOT", "Генерация штатного среза...", "System"
    UpsertLocKey db, "MSG_SNAPSHOT_DONE", "Штатный срез готов.", "System"

    ' Кнопки
    UpsertLocKey db, "BTN_SETTINGS", "Настройки", "UI"
    UpsertLocKey db, "BTN_IMPORT", "Импорт", "UI"
    UpsertLocKey db, "BTN_HEALTH_CHECK", "Health Check", "UI"
    UpsertLocKey db, "BTN_ANALYZE", "Анализ", "UI"
    UpsertLocKey db, "BTN_OPEN_LOG", "Журнал", "UI"
    UpsertLocKey db, "BTN_CREATE_INDEXES", "Создать индексы", "UI"
    UpsertLocKey db, "BTN_FULL_SYNC", "Полное обновление", "UI"
    UpsertLocKey db, "BTN_FIND_DUPLICATES", "Поиск дубликатов", "UI"
    UpsertLocKey db, "BTN_CHANGE_REPORT", "Changes Report", "UI"
    UpsertLocKey db, "BTN_SNAPSHOT", "Штатный срез", "UI"

    ' --- Новые ключи для интерактивного импорта и редактирования ---
    UpsertLocKey db, "BTN_EDIT_MAPPING", "Изменить заголовок", "UI"
    UpsertLocKey db, "TITLE_NEW_COL", "Новая колонка", "UI"
    UpsertLocKey db, "PROMPT_MAP_NEW_COL", "В файле найдена новая колонка:", "System"
    UpsertLocKey db, "PROMPT_MAP_NEW_COL2", "Добавить ее в базу данных и связать с профилем прямо сейчас?", "System"
    UpsertLocKey db, "PROMPT_ENTER_EN_NAME", "Введите техническое (английское) имя поля для БД (без пробелов):", "System"
    UpsertLocKey db, "TITLE_EDIT_MAPPING", "Редактирование маппинга", "UI"
    UpsertLocKey db, "PROMPT_EDIT_EXCEL_HEADER", "Введите новое точное название заголовка из Excel:", "System"
    UpsertLocKey db, "MSG_MAPPING_UPDATED", "Маппинг успешно обновлен.", "System"

    ' --- Фразы для умного восстановления маппинга ---
    UpsertLocKey db, "TITLE_RESTORE_LINK", "Восстановление связи", "UI"
    UpsertLocKey db, "PROMPT_RESTORE_LINK1", "В файле найдена колонка", "System"
    UpsertLocKey db, "PROMPT_RESTORE_LINK2", "уже существует в базе данных (связь в настройках отсутствует).", "System"
    UpsertLocKey db, "PROMPT_RESTORE_LINK3", "Восстановить связь для текущего профиля?", "System"

    Debug.Print "Localization seeding complete."
    MsgBox "Базовый словарь локализации успешно загружен в таблицу!", vbInformation, "StaffState Init"

    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка при заполнении локализации: " & Err.Description, vbCritical, "Error " & Err.Number
    Set db = Nothing
End Sub
