Attribute VB_Name = "mod_App_Init"
Option Compare Database

Option Explicit

' =============================================
' @author Кержаев Евгений (ФКУ "95 ФЭС" МО РФ)
' @description Модуль инициализации структуры БД.
'              Создает таблицы Buffer и Master с нуля или обновляет их.
' =============================================

' =============================================
' @description Главная процедура первичной настройки.
'              Запускать один раз при развертывании или для сброса структуры.
' =============================================
Public Sub InitDatabaseStructure()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    Debug.Print "--- Начало инициализации структуры ---"

    ' 1. Создаем таблицу БУФЕРА (Сырой импорт)
    ' Удаляем, если есть, чтобы очистить структуру
    DeleteTableIfExists "tbl_Import_Buffer"
    CreateBufferTable db

    ' 2. Создаем таблицу МАСТЕРА (Реестр)
    ' Удаляем только для теста! В боевом режиме тут будет логика Alter Table.
    ' Сейчас для старта создаем с нуля.
    If Not TableExists("tbl_Personnel_Master") Then
        CreateMasterTable db
    Else
        Debug.Print "Таблица 'tbl_Personnel_Master' уже существует. Пропуск."
    End If

    ' 3. Создаем таблицу ИСТОРИИ
    If Not TableExists("tbl_History_Log") Then
        CreateHistoryTable db
    End If

    Debug.Print "--- Инициализация успешно завершена ---"
    MsgBox "Структура базы данных успешно создана!", vbInformation, "StaffState Init"

    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка при инициализации: " & Err.Description, vbCritical, "Error " & Err.Number
    Set db = Nothing
End Sub

' =============================================
' @description Создает таблицу для сырого импорта (все поля Text).
' =============================================
Private Sub CreateBufferTable(db As DAO.Database)
    Dim sSQL As String

    ' Обрати внимание: Используем SHORT TEXT (255) для всех полей буфера,
    ' чтобы избежать ошибок типов при импорте Excel.
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
    Debug.Print "Создана таблица: tbl_Import_Buffer"
End Sub

' =============================================
' @description Создает главную таблицу персонала (Типизированную).
' =============================================
Private Sub CreateMasterTable(db As DAO.Database)
    Dim sSQL As String

    ' Тут уже строгие типы данных
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
    Debug.Print "Создана таблица: tbl_Personnel_Master"
End Sub

' =============================================
' @description Создает журнал изменений.
'              FIX: Убрано DEFAULT Now(), так как оно вызывает Error 3290 в DAO.
' =============================================
Private Sub CreateHistoryTable(db As DAO.Database)
    Dim sSQL As String

    ' Обрати внимание: поле ChangeDate теперь просто DATETIME без дефолта.
    ' Мы будем писать туда Now() программно при вставке строки.
    sSQL = "CREATE TABLE tbl_History_Log (" & _
           "LogID COUNTER CONSTRAINT PK_Log PRIMARY KEY, " & _
           "PersonUID VARCHAR(50), " & _
           "ChangeDate DATETIME, " & _
           "FieldName VARCHAR(100), " & _
           "OldValue MEMO, " & _
           "NewValue MEMO " & _
           ");"

    db.Execute sSQL, dbFailOnError

    ' Создаем индекс для ускорения поиска по человеку
    db.Execute "CREATE INDEX idx_Log_Person ON tbl_History_Log (PersonUID);", dbFailOnError

    Debug.Print "Создана таблица: tbl_History_Log"
End Sub

' =============================================
' @description Вспомогательная функция проверки существования таблицы.
' =============================================
Private Function TableExists(strTableName As String) As Boolean
    Dim tdf As DAO.TableDef
    On Error Resume Next
    Set tdf = CurrentDb.TableDefs(strTableName)
    TableExists = (Err.Number = 0)
    Err.Clear
End Function

' =============================================
' @description Вспомогательная функция безопасного удаления таблицы.
' =============================================
Private Sub DeleteTableIfExists(strTableName As String)
    If TableExists(strTableName) Then
        CurrentDb.Execute "DROP TABLE [" & strTableName & "];"
        Debug.Print "Удалена таблица: " & strTableName
    End If
End Sub
