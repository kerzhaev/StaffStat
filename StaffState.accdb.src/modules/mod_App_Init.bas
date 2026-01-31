Attribute VB_Name = "mod_App_Init"
Option Compare Database

Option Explicit

' =============================================
' @author ??????? ??????? (??? "95 ???" ?? ??)
' @description Database structure initialization module.
'              Creates Buffer and Master tables from scratch or updates them.
' =============================================

' =============================================
' @description Main initial setup procedure.
'              Run once during deployment or to reset structure.
' =============================================
Public Sub InitDatabaseStructure()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    Debug.Print "--- ?????? ????????????? ????????? ---"

    ' 1. Create BUFFER table (Raw import)
    ' Delete if exists to clear structure
    DeleteTableIfExists "tbl_Import_Buffer"
    CreateBufferTable db

    ' 2. Create MASTER table (Registry)
    ' Delete only for testing! In production mode there will be Alter Table logic here.
    ' For now we create from scratch to start.
    If Not TableExists("tbl_Personnel_Master") Then
        CreateMasterTable db
    Else
        Debug.Print "??????? 'tbl_Personnel_Master' ??? ??????????. ???????."
    End If

    ' 3. Create HISTORY table
    If Not TableExists("tbl_History_Log") Then
        CreateHistoryTable db
    End If

    ' 4. Create IMPORT METADATA table
    If Not TableExists("tbl_Import_Meta") Then
        CreateImportMetaTable db
    End If

    Debug.Print "--- ????????????? ??????? ????????? ---"
    MsgBox "Database structure created successfully!", vbInformation, "StaffState Init"

    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Initialization error: " & Err.Description, vbCritical, "Error " & Err.Number
    Set db = Nothing
End Sub

' =============================================
' @description Creates performance indexes for frequently queried fields.
'              Call this once after initial database setup.
'              Safe to run multiple times (will skip existing indexes).
' @author Phase 7 - Performance Improvements
' =============================================
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
        Debug.Print "? Created: idx_PersonUID on tbl_Personnel_Master (UNIQUE)"
        iCreated = iCreated + 1
    Else
        Debug.Print "? Skipped: idx_PersonUID already exists"
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
        Debug.Print "? Created: idx_FullName on tbl_Personnel_Master"
        iCreated = iCreated + 1
    Else
        Debug.Print "? Skipped: idx_FullName already exists"
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
        Debug.Print "? Created: idx_History_PersonUID on tbl_History_Log"
        iCreated = iCreated + 1
    Else
        Debug.Print "? Skipped: idx_History_PersonUID already exists"
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
        Debug.Print "? Created: idx_History_ChangeDate on tbl_History_Log"
        iCreated = iCreated + 1
    Else
        Debug.Print "? Skipped: idx_History_ChangeDate already exists"
        iSkipped = iSkipped + 1
    End If

    Debug.Print "--- Index Creation Complete ---"
    Debug.Print "Created: " & iCreated & " | Skipped: " & iSkipped

    If Not blnSuppressMsgBox Then
        MsgBox "Performance indexes created!" & vbCrLf & vbCrLf & _
               "Created: " & iCreated & vbCrLf & _
               "Skipped (already exist): " & iSkipped & vbCrLf & vbCrLf & _
               "Search and import should run faster now.", vbInformation, "Indexes"
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
               "Error number: " & Err.Number, vbCritical, "Error"
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
        Debug.Print "? Skipped: index already exists (" & idx.Name & ")"
        iSkipped = iSkipped + 1
        Err.Clear
        On Error GoTo ErrorHandler
        Resume Next
    End If
    Resume ErrorHandler
End Sub

' =============================================
' @description Helper function to check if index exists.
' @param strTableName [String] Name of the table.
' @param strIndexName [String] Name of the index.
' @return [Boolean] True if index exists.
' =============================================
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

' =============================================
' @description Creates table for raw import (all fields Text).
' =============================================
Private Sub CreateBufferTable(db As DAO.Database)
    Dim sSQL As String

    ' Note: Use SHORT TEXT (255) for all buffer fields
    ' to avoid type errors during Excel import.
    sSQL = "CREATE TABLE tbl_Import_Buffer (" & _
           "ID COUNTER CONSTRAINT PK_Buffer PRIMARY KEY, " & _
           "SourceID_Raw TEXT(255), " & _
           "PersonUID_Raw TEXT(255), " & _
           "Rank_Raw TEXT(255), " & _
           "FullName_Raw TEXT(255), " & _
           "BirthDate_Raw TEXT(255), " & _
           "WorkStatus_Raw TEXT(255), " & _
           "PosCode_Raw TEXT(255), " & _
           "PosName_Raw MEMO, " & _
           "OrderDate_Raw TEXT(255), " & _
           "OrderNum_Raw TEXT(255) " & _
           ");"

    db.Execute sSQL, dbFailOnError
    Debug.Print "??????? ???????: tbl_Import_Buffer"
End Sub

' =============================================
' @description Creates main personnel table (Typed).
' =============================================
Private Sub CreateMasterTable(db As DAO.Database)
    Dim sSQL As String

    ' Here we use strict data types
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
    Debug.Print "??????? ???????: tbl_Personnel_Master"
End Sub

' =============================================
' @description Creates change log.
'              FIX: Removed DEFAULT Now() as it causes Error 3290 in DAO.
' =============================================
Private Sub CreateHistoryTable(db As DAO.Database)
    Dim sSQL As String

    ' Note: ChangeDate field is now just DATETIME without default.
    ' We will write Now() programmatically when inserting a row.
    sSQL = "CREATE TABLE tbl_History_Log (" & _
           "LogID COUNTER CONSTRAINT PK_Log PRIMARY KEY, " & _
           "PersonUID VARCHAR(50), " & _
           "ChangeDate DATETIME, " & _
           "FieldName VARCHAR(100), " & _
           "OldValue MEMO, " & _
           "NewValue MEMO " & _
           ");"

    db.Execute sSQL, dbFailOnError

    ' Note: Indexes are now created centrally via CreatePerformanceIndexes()
    ' (removed old idx_Log_Person to avoid conflicts)

    Debug.Print "??????? ???????: tbl_History_Log"
End Sub

' =============================================
' @description Creates import metadata table (one row per import).
' =============================================
Private Sub CreateImportMetaTable(db As DAO.Database)
    Dim sSQL As String

    ' Single-row table to store metadata about last import
    sSQL = "CREATE TABLE tbl_Import_Meta (" & _
           "ID COUNTER CONSTRAINT PK_ImportMeta PRIMARY KEY, " & _
           "ExportFileDate DATETIME, " & _
           "ImportRunAt DATETIME, " & _
           "SourceFilePath TEXT(255) " & _
           ");"

    db.Execute sSQL, dbFailOnError
    Debug.Print "??????? ???????: tbl_Import_Meta"
End Sub

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

' =============================================
' @description Helper function for safe table deletion.
' =============================================
Private Sub DeleteTableIfExists(strTableName As String)
    If TableExists(strTableName) Then
        CurrentDb.Execute "DROP TABLE [" & strTableName & "];"
        Debug.Print "??????? ???????: " & strTableName
    End If
End Sub
