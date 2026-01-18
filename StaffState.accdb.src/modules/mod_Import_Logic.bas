Attribute VB_Name = "mod_Import_Logic"
Option Explicit

' =============================================
' @module mod_Import_Logic
' @author Кержаев Евгений
' @description Логика импорта из Excel с авто-определением имени листа
' =============================================

Private Const cstrLinkedTableName As String = "tmp_Excel_Link"

' =============================================
' @description Запускает процедуру импорта
' =============================================
Public Function ImportExcelData() As Boolean
    On Error GoTo ErrorHandler

    Dim strFilePath As String

    ' 1. Выбор файла
    strFilePath = SelectExcelFile()
    If strFilePath = "" Then Exit Function

    ' 2. Очистка буфера
    If Not ClearImportBuffer() Then Exit Function

    ' 3. Создание связи и импорт
    If Not CreateExcelLinkAndImport(strFilePath) Then Exit Function

    MsgBox "Импорт завершен успешно!", vbInformation
    ImportExcelData = True
    Exit Function

ErrorHandler:
    MsgBox "Критическая ошибка: " & Err.Description, vbCritical
    DeleteExcelLink
End Function

' =============================================
' @description Создает линк и выполняет INSERT
' =============================================
' =============================================
' @description Создает линк и выполняет INSERT (ПОЛНАЯ ВЕРСИЯ)
' =============================================
Private Function CreateExcelLinkAndImport(strPath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strConnect As String
    Dim strSheetName As String
    Dim strSourceRef As String
    Dim fld As DAO.Field

    Set db = CurrentDb
    DeleteExcelLink

    ' --- ШАГ 1: УЗНАЕМ ИМЯ ЛИСТА ---
    strSheetName = GetFirstSheetName(strPath)
    If strSheetName = "" Then
        MsgBox "Не удалось определить имя листа!", vbExclamation
        Exit Function
    End If

    ' --- ШАГ 2: СТРОКА ПОДКЛЮЧЕНИЯ ---
    If Right(strPath, 4) = ".xls" Then
        strConnect = "Excel 8.0;HDR=YES;IMEX=1;DATABASE=" & strPath
    Else
        strConnect = "Excel 12.0 Xml;HDR=YES;IMEX=1;DATABASE=" & strPath
    End If

    ' --- ШАГ 3: ЛИНК ---
    Set tdf = db.CreateTableDef(cstrLinkedTableName)
    tdf.Connect = strConnect
    If Right(strSheetName, 1) <> "$" Then
        strSourceRef = strSheetName & "$"
    Else
        strSourceRef = strSheetName
    End If
    tdf.SourceTableName = strSourceRef
    db.TableDefs.Append tdf

    ' --- ДИАГНОСТИКА ПОЛЕЙ (Смотри в Ctrl+G) ---
    Debug.Print "--- НАЙДЕННЫЕ ПОЛЯ В EXCEL ---"
    For Each fld In db.TableDefs(cstrLinkedTableName).Fields
        Debug.Print "Поле: " & fld.Name
    Next fld
    Debug.Print "------------------------------"

' --- ШАГ 4: ИМПОРТ (ПОЛНЫЙ SQL) ---
    Dim strSQL As String

    ' ЛОГИКА ДУБЛИКАТОВ:
    ' Access обычно переименовывает второе поле "Лицо" в "Лицо1".
    ' Мы пробуем взять ФИО из [Лицо1].

    strSQL = "INSERT INTO tbl_Import_Buffer (" & _
             "SourceID_Raw, PersonUID_Raw, Rank_Raw, FullName_Raw, " & _
             "BirthDate_Raw, WorkStatus_Raw, PosCode_Raw, PosName_Raw, " & _
             "OrderDate_Raw, OrderNum_Raw) " & _
             "SELECT " & _
             "T.[Лицо], " & _
             "T.[Личный номер], " & _
             "T.[Воинское звание], " & _
             "T.[ФИО], " & _
             "T.[Дата рождения], " & _
             "T.[Статус занятости], " & _
             "T.[Штатная должность], " & _
             "T.[Должность], " & _
             "T.[Дата приказа], " & _
             "T.[Номер приказа] " & _
             "FROM [" & cstrLinkedTableName & "] AS T " & _
             "WHERE T.[Личный номер] IS NOT NULL;"

    Debug.Print "SQL: " & strSQL
    db.Execute strSQL, dbFailOnError

    DeleteExcelLink
    CreateExcelLinkAndImport = True
    Exit Function

ErrorHandler:
    ' Если ошибка "Слишком мало параметров" (3061) - значит имя поля неправильное
    If Err.Number = 3061 Then
        MsgBox "ОШИБКА ПОЛЕЙ: Access не нашел одно из полей." & vbCrLf & _
               "Скорее всего, поле ФИО называется не 'Лицо1'." & vbCrLf & _
               "Нажмите Ctrl+G и пришлите список полей разработчику.", vbCritical
    Else
        MsgBox "Ошибка Import (Line " & Erl & "): " & Err.Description
    End If
    DeleteExcelLink
End Function

' =============================================
' @description Функция-разведчик. Открывает Excel и смотрит имя 1-го листа.
'              Использует Late Binding (CreateObject).
' =============================================
Private Function GetFirstSheetName(strPath As String) As String
    Dim xlApp As Object ' Excel.Application
    Dim xlWb As Object  ' Excel.Workbook

    On Error Resume Next
    Set xlApp = CreateObject("Excel.Application")
    If xlApp Is Nothing Then
        MsgBox "Не удалось запустить Excel для проверки файла.", vbCritical
        Exit Function
    End If

    ' Открываем скрытно и только для чтения
    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    Set xlWb = xlApp.Workbooks.Open(strPath, False, True) ' UpdateLinks=False, ReadOnly=True

    If Err.Number <> 0 Then
        Debug.Print "Ошибка открытия книги: " & Err.Description
        xlApp.Quit
        Set xlApp = Nothing
        Exit Function
    End If

    ' Берем имя первого листа
    GetFirstSheetName = xlWb.Sheets(1).Name

    ' Закрываем
    xlWb.Close False
    xlApp.Quit

    Set xlWb = Nothing
    Set xlApp = Nothing
End Function

' =============================================
' @description Диалог выбора файла
' =============================================
Public Function SelectExcelFile() As String
    Dim fd As Object
    Set fd = Application.FileDialog(3)
    With fd
        .Title = "Выберите файл выгрузки"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls"
        If .Show = -1 Then SelectExcelFile = .SelectedItems(1)
    End With
End Function

Private Function ClearImportBuffer() As Boolean
    On Error Resume Next
    CurrentDb.Execute "DELETE FROM tbl_Import_Buffer;", dbFailOnError
    ClearImportBuffer = True
End Function

Private Sub DeleteExcelLink()
    On Error Resume Next
    CurrentDb.TableDefs.Delete cstrLinkedTableName
End Sub
