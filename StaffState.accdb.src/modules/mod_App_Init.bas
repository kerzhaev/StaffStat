Attribute VB_Name = "mod_App_Init"
Option Compare Database
Option Explicit

' =============================================
' @module mod_App_Init
' @author Evgeniy Kerzhaev
' @description Database structure initialization module.
'              Creates Buffer and Master tables from scratch or updates them.
' @note 100% English version. Safe for modern IDEs.
' =============================================

Public Sub InitializeApp()
    On Error GoTo ErrorHandler
    InitDatabaseStructure
    Exit Sub
ErrorHandler:
    Debug.Print "InitializeApp error: " & Err.Description & " (" & Err.Number & ")"
    MsgBox "Initialization failed: " & Err.Description, vbCritical, "StaffState Error"
End Sub

Public Sub InitDatabaseStructure()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    Debug.Print "--- Init database structure start ---"

    ' 1. Create BUFFER table (Raw import)
    ' Delete if exists to clear structure
    DeleteTableIfExists "tbl_Import_Buffer"
    CreateBufferTable db

    ' 2. Create MASTER table (Registry)
    ' In production mode there will be Alter Table logic here.
    ' For now we create from scratch to start.
    If Not TableExists("tbl_Personnel_Master") Then
        CreateMasterTable db
    Else
        Debug.Print "Table 'tbl_Personnel_Master' already exists, skip."
    End If

    ' 3. Create HISTORY table
    If Not TableExists("tbl_History_Log") Then
        CreateHistoryTable db
    End If

    ' 4. Create IMPORT METADATA table
    If Not TableExists("tbl_Import_Meta") Then
        CreateImportMetaTable db
    End If

    ' 5. Create SETTINGS table (Phase 11)
    mod_Schema_Manager.CreateSettingsTable

    ' 6. Create VALIDATION LOG table (Phase 11)
    mod_Schema_Manager.CreateValidationLogTable

    ' 7. Create Import Mapping tables (Phase 19-20) and seed Profile 1
    mod_Schema_Manager.CreateImportProfilesTable
    mod_Schema_Manager.CreateImportMappingTable
    mod_Schema_Manager.SeedImportMappingProfile1

    Debug.Print "--- Init database structure complete ---"
    MsgBox "Database structure created successfully!", vbInformation, "StaffState Init"

    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Initialization error: " & Err.Description, vbCritical, "Error " & Err.Number
    Set db = Nothing
End Sub

Public Sub CreatePerformanceIndexes(ByRef outCreated As Long, ByRef outSkipped As Long, Optional ByVal blnSuppressMsgBox As Boolean = False)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim iCreated As Long
    Dim iSkipped As Long

    Set db = CurrentDb
    iCreated = 0
    iSkipped = 0
    outCreated = 0
    outSkipped = 0

    Debug.Print "--- Creating Performance Indexes ---"

    ' ===== INDEX 1: tbl_Personnel_Master.PersonUID (UNIQUE) =====
    If Not IndexExists("tbl_Personnel_Master", "idx_PersonUID") Then
        Set tdf = db.TableDefs("tbl_Personnel_Master")
        Set idx = tdf.CreateIndex("idx_PersonUID")
        Set fld = idx.CreateField("PersonUID")
        idx.Fields.Append fld
        idx.Primary = False
        idx.Unique = True
        idx.Required = False
        On Error GoTo IndexAppendError
        tdf.Indexes.Append idx
        On Error GoTo ErrorHandler
        Debug.Print "Created: idx_PersonUID on tbl_Personnel_Master (UNIQUE)"
        iCreated = iCreated + 1
    Else
        Debug.Print "Skipped: idx_PersonUID already exists"
        iSkipped = iSkipped + 1
    End If

    ' ===== INDEX 2: tbl_Personnel_Master.FullName =====
    If Not IndexExists("tbl_Personnel_Master", "idx_FullName") Then
        Set tdf = db.TableDefs("tbl_Personnel_Master")
        Set idx = tdf.CreateIndex("idx_FullName")
        Set fld = idx.CreateField("FullName")
        idx.Fields.Append fld
        idx.Primary = False
        idx.Unique = False
        On Error GoTo IndexAppendError
        tdf.Indexes.Append idx
        On Error GoTo ErrorHandler
        Debug.Print "Created: idx_FullName on tbl_Personnel_Master"
        iCreated = iCreated + 1
    Else
        Debug.Print "Skipped: idx_FullName already exists"
        iSkipped = iSkipped + 1
    End If

    ' ===== INDEX 3: tbl_History_Log.PersonUID =====
    If Not IndexExists("tbl_History_Log", "idx_History_PersonUID") Then
        Set tdf = db.TableDefs("tbl_History_Log")
        Set idx = tdf.CreateIndex("idx_History_PersonUID")
        Set fld = idx.CreateField("PersonUID")
        idx.Fields.Append fld
        idx.Primary = False
        idx.Unique = False
        On Error GoTo IndexAppendError
        tdf.Indexes.Append idx
        On Error GoTo ErrorHandler
        Debug.Print "Created: idx_History_PersonUID on tbl_History_Log"
        iCreated = iCreated + 1
    Else
        Debug.Print "Skipped: idx_History_PersonUID already exists"
        iSkipped = iSkipped + 1
    End If

    ' ===== INDEX 4: tbl_History_Log.ChangeDate =====
    If Not IndexExists("tbl_History_Log", "idx_History_ChangeDate") Then
        Set tdf = db.TableDefs("tbl_History_Log")
        Set idx = tdf.CreateIndex("idx_History_ChangeDate")
        Set fld = idx.CreateField("ChangeDate")
        idx.Fields.Append fld
        idx.Primary = False
        idx.Unique = False
        On Error GoTo IndexAppendError
        tdf.Indexes.Append idx
        On Error GoTo ErrorHandler
        Debug.Print "Created: idx_History_ChangeDate on tbl_History_Log"
        iCreated = iCreated + 1
    Else
        Debug.Print "Skipped: idx_History_ChangeDate already exists"
        iSkipped = iSkipped + 1
    End If

    Debug.Print "--- Index Creation Complete ---"
    Debug.Print "Created: " & iCreated & " | Skipped: " & iSkipped

    If Not blnSuppressMsgBox Then
        MsgBox "Performance indexes created!" & vbCrLf & vbCrLf & _
               "Created: " & iCreated & vbCrLf & _
               "Skipped (already exist): " & iSkipped & vbCrLf & vbCrLf & _
               "Search and import should run faster now.", vbInformation, "System Indexes"
    End If

    outCreated = iCreated
    outSkipped = iSkipped

    Set fld = Nothing
    Set idx = Nothing
    Set tdf = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "ERROR: " & Err.Description
    If Not blnSuppressMsgBox Then
        MsgBox "Index creation error: " & Err.Description & vbCrLf & _
               "Error number: " & Err.Number, vbCritical, "System Error"
    End If
    outCreated = iCreated
    outSkipped = iSkipped
    Set fld = Nothing
    Set idx = Nothing
    Set tdf = Nothing
    Set db = Nothing
    Exit Sub

IndexAppendError:
    If Err.Number = 3284 Then
        Debug.Print "Skipped: index already exists (" & idx.Name & ")"
        iSkipped = iSkipped + 1
        Err.Clear
        On Error GoTo ErrorHandler
        Resume Next
    End If
    Resume ErrorHandler
End Sub

Private Function IndexExists(strTableName As String, strIndexName As String) As Boolean
    On Error Resume Next
    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index

    Set tdf = CurrentDb.TableDefs(strTableName)
    Set idx = tdf.Indexes(strIndexName)

    IndexExists = (Err.Number = 0)
    Err.Clear

    Set idx = Nothing
    Set tdf = Nothing
End Function

Private Sub CreateBufferTable(db As DAO.Database)
    Dim sSQL As String
    sSQL = "CREATE TABLE tbl_Import_Buffer (" & _
           "ID COUNTER CONSTRAINT PK_Buffer PRIMARY KEY, " & _
           "PersonUID TEXT(255), SourceID LONG, FullName TEXT(255), RankName TEXT(255), " & _
           "BirthDate_Text TEXT(255), WorkStatus TEXT(255), PosCode TEXT(255), PosName LONGTEXT, " & _
           "OrderDate_Text TEXT(255), OrderNumber TEXT(255), EmployeeAge TEXT(255), Gender TEXT(255), " & _
           "MaritalStatus TEXT(255), ChildrenCount TEXT(255), Nationality TEXT(255), Citizenship TEXT(255), " & _
           "ContractType TEXT(255), ContractKind TEXT(255), ContractStartDate TEXT(255), ContractEndDate TEXT(255), " & _
           "ContractYears LONG, ContractMonths LONG, EventType TEXT(255), EventReason TEXT(255), " & _
           "ValidFromDate TEXT(255), ValidToDate TEXT(255), StaffPosition TEXT(255), Position TEXT(255), " & _
           "VUS TEXT(255), SalaryGrade TEXT(255), PersonnelDivision TEXT(255), BankAccountNumber TEXT(255), " & _
           "Payee TEXT(255), BankKey TEXT(255), BootSize TEXT(255), HeadSize TEXT(255)" & _
           ");"
    db.Execute sSQL, dbFailOnError
    Debug.Print "Created table: tbl_Import_Buffer"
End Sub

Private Sub CreateMasterTable(db As DAO.Database)
    Dim sSQL As String
    sSQL = "CREATE TABLE tbl_Personnel_Master (" & _
           "PersonUID VARCHAR(50) CONSTRAINT PK_Person PRIMARY KEY, " & _
           "SourceID LONG, " & _
           "FullName VARCHAR(150), " & _
           "RankName VARCHAR(100), " & _
           "BirthDate DATETIME, " & _
           "WorkStatus VARCHAR(100), " & _
           "PosCode VARCHAR(50), " & _
           "PosName MEMO, " & _
           "OrderDate DATETIME, " & _
           "OrderNum VARCHAR(50), " & _
           "LastUpdated DATETIME, " & _
           "IsActive BIT " & _
           ");"
    db.Execute sSQL, dbFailOnError
    Debug.Print "Created table: tbl_Personnel_Master"
End Sub

Private Sub CreateHistoryTable(db As DAO.Database)
    Dim sSQL As String
    sSQL = "CREATE TABLE tbl_History_Log (" & _
           "LogID COUNTER CONSTRAINT PK_Log PRIMARY KEY, " & _
           "PersonUID VARCHAR(50), " & _
           "ChangeDate DATETIME, " & _
           "FieldName VARCHAR(100), " & _
           "OldValue MEMO, " & _
           "NewValue MEMO " & _
           ");"
    db.Execute sSQL, dbFailOnError
    Debug.Print "Created table: tbl_History_Log"
End Sub

Private Sub CreateImportMetaTable(db As DAO.Database)
    Dim sSQL As String
    sSQL = "CREATE TABLE tbl_Import_Meta (" & _
           "ID COUNTER CONSTRAINT PK_ImportMeta PRIMARY KEY, " & _
           "ExportFileDate DATETIME, " & _
           "ImportRunAt DATETIME, " & _
           "SourceFilePath TEXT(255) " & _
           ");"
    db.Execute sSQL, dbFailOnError
    Debug.Print "Created table: tbl_Import_Meta"
End Sub

Private Function TableExists(strTableName As String) As Boolean
    Dim tdf As DAO.TableDef
    On Error Resume Next
    Set tdf = CurrentDb.TableDefs(strTableName)
    TableExists = (Err.Number = 0)
    Err.Clear
End Function

Private Sub DeleteTableIfExists(strTableName As String)
    If TableExists(strTableName) Then
        CurrentDb.Execute "DROP TABLE [" & strTableName & "];"
        Debug.Print "Dropped table: " & strTableName
    End If
End Sub
