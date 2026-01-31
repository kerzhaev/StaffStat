Attribute VB_Name = "mod_UI_Helpers"
Option Compare Database
Option Explicit

' =============================================
' @module mod_UI_Helpers
' @description UI helper functions (friendly captions, history formatting)
' @note Keep VBA source ASCII-only for stability; build RU strings via ChrW.
' =============================================

' =============================================
' @description Returns "(?????)" in Russian (ASCII-safe).
' =============================================
Private Function RuEmptyToken() As String
    RuEmptyToken = "(" & _
                   ChrW(1087) & ChrW(1091) & ChrW(1089) & ChrW(1090) & ChrW(1086) & _
                   ")"
End Function

' =============================================
' @description Returns Russian caption for internal field name (ASCII-safe).
' @param strInternalName [String] Internal field name (e.g., "RankName")
' @return [String] Russian caption
' =============================================
Public Function GetFieldFriendlyName(ByVal strInternalName As String) As String
    On Error GoTo ErrorHandler

    Dim s As String
    s = Trim$(Nz(strInternalName, ""))

    Select Case s
        Case "_System"
            ' "????"
            GetFieldFriendlyName = ChrW(1059) & ChrW(1095) & ChrW(1077) & ChrW(1090)

        Case "RankName"
            ' "??????"
            GetFieldFriendlyName = ChrW(1047) & ChrW(1074) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1077)

        Case "WorkStatus"
            ' "??????"
            GetFieldFriendlyName = ChrW(1057) & ChrW(1090) & ChrW(1072) & ChrW(1090) & ChrW(1091) & ChrW(1089)

        Case "PosName"
            ' "?????????"
            GetFieldFriendlyName = ChrW(1044) & ChrW(1086) & ChrW(1083) & ChrW(1078) & ChrW(1085) & ChrW(1086) & ChrW(1089) & ChrW(1090) & ChrW(1100)

        Case "PosCode"
            ' "??? ?????????"
            GetFieldFriendlyName = ChrW(1050) & ChrW(1086) & ChrW(1076) & " " & _
                                   ChrW(1076) & ChrW(1086) & ChrW(1083) & ChrW(1078) & ChrW(1085) & ChrW(1086) & ChrW(1089) & ChrW(1090) & ChrW(1080)

        Case "FullName"
            ' "???"
            GetFieldFriendlyName = ChrW(1060) & ChrW(1048) & ChrW(1054)

        Case "PersonUID"
            ' "?????? ?????"
            GetFieldFriendlyName = ChrW(1051) & ChrW(1080) & ChrW(1095) & ChrW(1085) & ChrW(1099) & ChrW(1081) & " " & _
                                   ChrW(1085) & ChrW(1086) & ChrW(1084) & ChrW(1077) & ChrW(1088)

        Case "SourceID"
            ' "????????"
            GetFieldFriendlyName = ChrW(1048) & ChrW(1089) & ChrW(1090) & ChrW(1086) & ChrW(1095) & ChrW(1085) & ChrW(1080) & ChrW(1082)

        Case "OrderDate"
            ' "???? ???????"
            GetFieldFriendlyName = ChrW(1044) & ChrW(1072) & ChrW(1090) & ChrW(1072) & " " & _
                                   ChrW(1087) & ChrW(1088) & ChrW(1080) & ChrW(1082) & ChrW(1072) & ChrW(1079) & ChrW(1072)

        Case "OrderNum"
            ' "????? ???????"
            GetFieldFriendlyName = ChrW(1053) & ChrW(1086) & ChrW(1084) & ChrW(1077) & ChrW(1088) & " " & _
                                   ChrW(1087) & ChrW(1088) & ChrW(1080) & ChrW(1082) & ChrW(1072) & ChrW(1079) & ChrW(1072)

        Case "BirthDate"
            ' "???? ????????"
            GetFieldFriendlyName = ChrW(1044) & ChrW(1072) & ChrW(1090) & ChrW(1072) & " " & _
                                   ChrW(1088) & ChrW(1086) & ChrW(1078) & ChrW(1076) & ChrW(1077) & ChrW(1085) & ChrW(1080) & ChrW(1103)

        Case Else
            GetFieldFriendlyName = Replace(s, "_", " ")
    End Select

    Exit Function
ErrorHandler:
    GetFieldFriendlyName = Trim$(Nz(strInternalName, ""))
End Function

' =============================================
' @description Translates system event token to Russian (ASCII-safe).
' @param sToken [String] Token stored in DB (ASCII)
' @return [String] Russian user-facing text
' =============================================
Private Function TranslateSystemToken(ByVal sToken As String) As String
    Dim s As String
    s = UCase$(Trim$(Nz(sToken, "")))

    Select Case s
        Case "ADDED"
            ' "?????? ?? ????"
            TranslateSystemToken = ChrW(1055) & ChrW(1088) & ChrW(1080) & ChrW(1085) & ChrW(1103) & ChrW(1090) & " " & _
                                   ChrW(1085) & ChrW(1072) & " " & _
                                   ChrW(1091) & ChrW(1095) & ChrW(1077) & ChrW(1090)
        Case "REMOVED"
            ' "???? ? ?????"
            TranslateSystemToken = ChrW(1057) & ChrW(1085) & ChrW(1103) & ChrW(1090) & " " & _
                                   ChrW(1089) & " " & _
                                   ChrW(1091) & ChrW(1095) & ChrW(1077) & ChrW(1090) & ChrW(1072)
        Case Else
            TranslateSystemToken = Trim$(Nz(sToken, ""))
    End Select
End Function

' =============================================
' @description Builds a human-friendly history description line.
' @param strInternalName [String] Field internal name (FieldName from tbl_History_Log)
' @param vOld [Variant] OldValue
' @param vNew [Variant] NewValue
' @return [String] Human-friendly description string
' =============================================
Public Function BuildHistoryDescription(ByVal strInternalName As String, ByVal vOld As Variant, ByVal vNew As Variant) As String
    On Error GoTo ErrorHandler

    Dim sOld As String
    Dim sNew As String
    Dim sNewDisplay As String
    Dim sCaption As String
    Dim sMarker As String

    sOld = Trim$(Nz(vOld, ""))
    sNew = Trim$(Nz(vNew, ""))
    sCaption = GetFieldFriendlyName(strInternalName)

    If Trim$(Nz(strInternalName, "")) = "_System" Then
        sNewDisplay = TranslateSystemToken(sNew)
    Else
        sNewDisplay = sNew
    End If

    If sOld = "" And sNew <> "" Then
        sMarker = "[+]"
        sOld = RuEmptyToken()
    ElseIf sOld <> "" And sNew = "" Then
        sMarker = "[-]"
        sNewDisplay = RuEmptyToken()
    Else
        sMarker = "[*]"
        If sOld = "" Then sOld = RuEmptyToken()
        If sNewDisplay = "" Then sNewDisplay = RuEmptyToken()
    End If

    BuildHistoryDescription = sMarker & " " & sCaption & ": " & sOld & " -> " & sNewDisplay
    Exit Function

ErrorHandler:
    BuildHistoryDescription = "[*] " & Trim$(Nz(strInternalName, "")) & ": " & Trim$(Nz(vOld, "")) & " -> " & Trim$(Nz(vNew, ""))
End Function
