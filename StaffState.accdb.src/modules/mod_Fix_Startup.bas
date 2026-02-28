Attribute VB_Name = "mod_Fix_Startup"
Option Compare Database

Option Explicit

' =============================================
' @author ??????? ???????
' @description Fix startup form settings
' =============================================

' =============================================
' @description Set startup form to uf_Dashboard
' Run this once if startup settings are broken
' =============================================
Public Sub FixStartupForm()
    On Error GoTo ErrorHandler

    Dim prp As DAO.Property
    Dim db As DAO.Database

    Set db = CurrentDb

    ' Try to get existing property
    On Error Resume Next
    Set prp = db.Properties("StartUpForm")

    If Err.Number <> 0 Then
        ' Property doesn't exist - create it
        Err.Clear
        On Error GoTo ErrorHandler
        Set prp = db.CreateProperty("StartUpForm", dbText, "uf_Dashboard")
        db.Properties.Append prp
        Debug.Print "Created StartUpForm property"
    Else
        ' Property exists - update it
        On Error GoTo ErrorHandler
        prp.value = "uf_Dashboard"
        Debug.Print "Updated StartUpForm property"
    End If

    MsgBox "Startup form set to: uf_Dashboard" & vbCrLf & vbCrLf & _
           "Please close and reopen the database.", _
           vbInformation, "Settings fixed"

    Set prp = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description & " (" & Err.Number & ")", vbCritical
    Set prp = Nothing
    Set db = Nothing
End Sub

' =============================================
' @description Clear startup form (remove setting)
' =============================================
Public Sub ClearStartupForm()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    On Error Resume Next
    db.Properties.Delete "StartUpForm"

    If Err.Number = 0 Then
        MsgBox "Startup form cleared. The database will open with the navigation pane.", vbInformation
    Else
        MsgBox "StartUpForm property not found.", vbInformation
    End If

    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Set db = Nothing
End Sub
