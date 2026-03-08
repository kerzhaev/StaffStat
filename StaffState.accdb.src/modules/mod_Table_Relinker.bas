Attribute VB_Name = "mod_Table_Relinker"
Option Compare Database
Option Explicit

' =============================================
' @module mod_Table_Relinker
' @description Auto-relinker for Front-End / Back-End architecture
' @note 100% English version. Safe for modern IDEs.
' =============================================

Private Const cstrBackendName As String = "StaffState_BE.accdb"

' =============================================
' @description Checks and refreshes table links to the Back-End database.
' @return [Boolean] True if links are OK or successfully relinked.
' =============================================
Public Function VerifyAndRelinkTables(Optional ByVal blnForceRelink As Boolean = False) As Boolean
    Dim result As Object

    Set result = VerifyAndRelinkTablesResult(blnForceRelink)
    VerifyAndRelinkTables = CBool(result("Success"))
    Set result = Nothing
End Function

Public Function VerifyAndRelinkTablesResult(Optional ByVal blnForceRelink As Boolean = False) As Object
    On Error GoTo ErrorHandler

    Dim result As Object
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strExpectedBEPath As String
    Dim strCurrentBEPath As String
    Dim blnNeedsRelink As Boolean
    Dim iLinkedCount As Long

    Set result = CreateRelinkerResult()
    Set db = CurrentDb

    ' 1. Determine expected Back-End path (same folder as Front-End)
    strExpectedBEPath = CurrentProject.Path & "\" & cstrBackendName

    ' 2. Verify Back-End file exists
    If Dir(strExpectedBEPath) = "" Then
        result("Status") = "NOT_FOUND"
        result("ErrorMessage") = "Back-End database not found at:" & vbCrLf & strExpectedBEPath
        result("Message") = CStr(result("ErrorMessage"))
        GoTo Cleanup
    End If

    ' 3. Check if we need to relink
    blnNeedsRelink = blnForceRelink
    iLinkedCount = 0

    If Not blnNeedsRelink Then
        For Each tdf In db.TableDefs
            ' If table is linked (Connect string is not empty) and not a system table
            If Len(tdf.Connect) > 0 And Left$(tdf.Name, 4) <> "MSys" And Left$(tdf.Name, 4) <> "USys" Then
                iLinkedCount = iLinkedCount + 1
                ' Extract path from Connect string (Format: ";DATABASE=C:\...\StaffState_BE.accdb")
                strCurrentBEPath = Mid$(tdf.Connect, InStr(tdf.Connect, "DATABASE=") + 9)
                If StrComp(strCurrentBEPath, strExpectedBEPath, vbTextCompare) <> 0 Then
                    blnNeedsRelink = True
                    Exit For
                End If
            End If
        Next tdf
    End If

    ' If no linked tables exist at all, we might be in the monolithic file or early stage
    If iLinkedCount = 0 And Not blnNeedsRelink Then
        Debug.Print "No linked tables found. Running as monolithic DB."
        result("Success") = True
        result("Status") = "MONOLITHIC"
        result("Message") = ""
        GoTo Cleanup
    End If

    ' 4. Perform relinking if needed
    If blnNeedsRelink Then
        Debug.Print "Relinking tables to: " & strExpectedBEPath
        For Each tdf In db.TableDefs
            If Len(tdf.Connect) > 0 And Left$(tdf.Name, 4) <> "MSys" And Left$(tdf.Name, 4) <> "USys" Then
                tdf.Connect = ";DATABASE=" & strExpectedBEPath
                tdf.RefreshLink
                Debug.Print "Relinked: " & tdf.Name
            End If
        Next tdf
        Debug.Print "Relinking complete."
    Else
        Debug.Print "Table links are up to date."
    End If

    result("Success") = True
    If blnNeedsRelink Then
        result("Status") = "RELINKED"
    Else
        result("Status") = "OK"
    End If

Cleanup:
    On Error Resume Next
    Set tdf = Nothing
    Set db = Nothing
    Set VerifyAndRelinkTablesResult = result
    Exit Function

ErrorHandler:
    Debug.Print "VerifyAndRelinkTables error: " & Err.Description & " (" & Err.Number & ")"
    result("Success") = False
    result("Status") = "ERROR"
    result("ErrorNumber") = Err.Number
    result("ErrorMessage") = "Failed to relink tables: " & Err.Description
    result("Message") = CStr(result("ErrorMessage"))
    GoTo Cleanup
End Function

Private Function CreateRelinkerResult() As Object
    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d("Success") = False
    d("Status") = "PENDING"
    d("Message") = ""
    d("ErrorMessage") = ""
    d("ErrorNumber") = 0

    Set CreateRelinkerResult = d
End Function
