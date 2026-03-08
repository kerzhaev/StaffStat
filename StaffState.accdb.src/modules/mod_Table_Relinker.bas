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
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strExpectedBEPath As String
    Dim strCurrentBEPath As String
    Dim blnNeedsRelink As Boolean
    Dim iLinkedCount As Long

    VerifyAndRelinkTables = False
    Set db = CurrentDb

    ' 1. Determine expected Back-End path (same folder as Front-End)
    strExpectedBEPath = CurrentProject.Path & "\" & cstrBackendName

    ' 2. Verify Back-End file exists
    If Dir(strExpectedBEPath) = "" Then
        mod_UI_Helpers.ShowMessage "Back-End database not found at:" & vbCrLf & strExpectedBEPath, vbCritical
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
        VerifyAndRelinkTables = True
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

    VerifyAndRelinkTables = True

Cleanup:
    On Error Resume Next
    Set tdf = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    Debug.Print "VerifyAndRelinkTables error: " & Err.Description & " (" & Err.Number & ")"
    mod_UI_Helpers.ShowMessage "Failed to relink tables: " & Err.Description, vbCritical
    VerifyAndRelinkTables = False
    GoTo Cleanup
End Function
