Attribute VB_Name = "mod_Schema_Manager"
Option Compare Database
Option Explicit

' =============================================
' @module mod_Schema_Manager
' @description On-the-fly table structure management
' =============================================

' =============================================
' @description Converts "Дата рожд." to "Дата_рожд" (Access doesn't like spaces and dots in SQL)
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

    ' Remove double underscores
    While InStr(s, "__") > 0
        s = Replace(s, "__", "_")
    Wend

    ' Remove underscore at end and beginning
    If Right(s, 1) = "_" Then s = Left(s, Len(s) - 1)
    If Left(s, 1) = "_" Then s = Mid(s, 2)

    SanitizeFieldName = s
End Function

' =============================================
' @description Checks if field exists in table. If not - creates it.
' @param strTableName Table name (Buffer or Master)
' @param strFieldName Field name
' @param strType SQL data type (default TEXT(255))
' =============================================
Public Sub EnsureFieldExists(strTableName As String, strFieldName As String, Optional strType As String = "TEXT(255)")
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

    ' If field doesn't exist - add it via DDL
    If Not blnExists Then
        Dim strSQL As String
        strSQL = "ALTER TABLE [" & strTableName & "] ADD COLUMN [" & strFieldName & "] " & strType & ";"
        db.Execute strSQL, dbFailOnError
        Debug.Print "?? Добавлено новое поле в " & strTableName & ": " & strFieldName
    End If
End Sub

' =============================================
' @description Copies field structure from Buffer to Master (if they don't exist there)
' =============================================
Public Sub SyncMasterStructure()
    Dim db As DAO.Database
    Dim tdfBuffer As DAO.TableDef
    Dim fld As DAO.Field

    Set db = CurrentDb
    Set tdfBuffer = db.TableDefs("tbl_Import_Buffer")

    ' Loop through all Buffer fields
    For Each fld In tdfBuffer.Fields
        ' Skip system fields (ID, SourceID_Raw, etc., if they're already mapped)
        ' But we have dynamics, so just check for existence.

        ' Exclude Access system fields (if any)
        If Left(fld.Name, 4) <> "s_Col" Then
             ' Try to create the same field in Master
             ' By default create as TEXT(255), since everything in Buffer is text
             ' Ideally we could map types, but for MVP text is more reliable.

             ' Note: In Buffer fields are named ..._Raw (e.g., Rank_Raw)
             ' But in Master we want just Rank?
             ' IN CURRENT DYNAMICS LOGIC: We will create fields "one to one"
             ' for new unknown columns.

             Dim strTargetName As String
             strTargetName = fld.Name

             ' If these are our "standard" fields with _Raw suffix, we already processed them manually in CREATE TABLE.
             ' But "Размер_Сапог" doesn't have a suffix.

             EnsureFieldExists "tbl_Personnel_Master", strTargetName, "TEXT(255)"
        End If
    Next fld
End Sub
