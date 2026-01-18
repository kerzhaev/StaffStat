Attribute VB_Name = "mod_Analysis_Logic"
Option Explicit

' =============================================
' @module mod_Analysis_Logic
' @author Кержаев Евгений (ФКУ "95 ФЭС" МО РФ)
' @description Логика анализа и синхронизации данных из Buffer в Master.
'              Отслеживает изменения (Звание, Должность, Статус) и фиксирует в History.
' =============================================

' =============================================
' @description Главная процедура синхронизации.
'              Переносит данные из tbl_Import_Buffer в tbl_Personnel_Master.
'              Новых добавляет, существующих обновляет с проверкой изменений.
' =============================================
Public Sub SyncBufferToMaster()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rsBuffer As DAO.Recordset
    Dim rsMaster As DAO.Recordset
    Dim strPersonUID As String
    Dim lngNewCount As Long
    Dim lngUpdatedCount As Long
    Dim lngChangedCount As Long

    Set db = CurrentDb

    Debug.Print "--- Начало синхронизации Buffer -> Master ---"

    ' Открываем Recordset буфера (все записи)
    Set rsBuffer = db.OpenRecordset("SELECT * FROM tbl_Import_Buffer ORDER BY PersonUID_Raw;", dbOpenDynaset)

    ' Открываем Recordset мастера для поиска и обновления
    Set rsMaster = db.OpenRecordset("SELECT * FROM tbl_Personnel_Master;", dbOpenDynaset)

    If rsBuffer.EOF And rsBuffer.BOF Then
        Debug.Print "Буфер пуст. Нет данных для синхронизации."
        MsgBox "Буфер пуст. Сначала выполните импорт данных.", vbInformation
        GoTo CleanExit
    End If

    ' Цикл по всем записям буфера
    Do While Not rsBuffer.EOF
        strPersonUID = Nz(rsBuffer!PersonUID_Raw, "")

        ' Пропускаем записи без PersonUID
        If strPersonUID = "" Then
            rsBuffer.MoveNext
            GoTo NextRecord
        End If

        ' Ищем человека в Мастере по PersonUID
        rsMaster.FindFirst "PersonUID = '" & Replace(strPersonUID, "'", "''") & "'"

        If rsMaster.NoMatch Then
            ' --- НОВЫЙ СОТРУДНИК ---
            Call AddNewPersonToMaster(rsBuffer, rsMaster)
            Call LogChange(strPersonUID, "Статус", "", "Принят на учет")
            lngNewCount = lngNewCount + 1
            Debug.Print "Добавлен новый: " & strPersonUID
        Else
            ' --- СУЩЕСТВУЮЩИЙ СОТРУДНИК ---
            Dim lngChangesInRecord As Long
            lngChangesInRecord = UpdateExistingPerson(rsBuffer, rsMaster, strPersonUID)
            If lngChangesInRecord > 0 Then
                lngChangedCount = lngChangedCount + lngChangesInRecord
                lngUpdatedCount = lngUpdatedCount + 1
            End If
        End If

NextRecord:
        rsBuffer.MoveNext
    Loop

    Debug.Print "--- Синхронизация завершена ---"
    Debug.Print "  Новых записей: " & lngNewCount
    Debug.Print "  Обновлено записей: " & lngUpdatedCount
    Debug.Print "  Всего изменений: " & lngChangedCount

    MsgBox "Синхронизация завершена!" & vbCrLf & _
           "Новых: " & lngNewCount & vbCrLf & _
           "Обновлено: " & lngUpdatedCount & vbCrLf & _
           "Изменений: " & lngChangedCount, vbInformation, "Sync Complete"

CleanExit:
    rsBuffer.Close
    rsMaster.Close
    Set rsBuffer = Nothing
    Set rsMaster = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка SyncBufferToMaster (Line " & Erl & "): " & Err.Description
    MsgBox "Критическая ошибка синхронизации: " & Err.Description & vbCrLf & _
           "Номер: " & Err.Number, vbCritical
    Resume CleanExit
End Sub

' =============================================
' @description Добавляет нового сотрудника в Master из Buffer.
' @param   rsBuffer [DAO.Recordset] Запись из буфера (текущая).
' @param   rsMaster [DAO.Recordset] Recordset мастера (открыт для добавления).
' =============================================
Private Sub AddNewPersonToMaster(rsBuffer As DAO.Recordset, rsMaster As DAO.Recordset)
    On Error GoTo ErrorHandler

    rsMaster.AddNew

    ' Ключевое поле (обязательное)
    rsMaster!PersonUID = Nz(rsBuffer!PersonUID_Raw, "")

    ' Преобразование типов из текста
    rsMaster!SourceID = ConvertToLong(rsBuffer!SourceID_Raw)
    rsMaster!FullName = Nz(rsBuffer!FullName_Raw, "")
    rsMaster!RankName = Nz(rsBuffer!Rank_Raw, "")
    rsMaster!BirthDate = ConvertToDate(rsBuffer!BirthDate_Raw)
    rsMaster!WorkStatus = Nz(rsBuffer!WorkStatus_Raw, "")
    rsMaster!PosCode = Nz(rsBuffer!PosCode_Raw, "")
    rsMaster!PosName = Nz(rsBuffer!PosName_Raw, "")
    rsMaster!OrderDate = ConvertToDate(rsBuffer!OrderDate_Raw)
    rsMaster!OrderNum = Nz(rsBuffer!OrderNum_Raw, "")

    ' Служебные поля
    rsMaster!LastUpdated = Now()
    rsMaster!IsActive = True

    rsMaster.Update

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка AddNewPersonToMaster: " & Err.Description
    If rsMaster.EditMode <> dbEditNone Then rsMaster.CancelUpdate
    Err.Raise Err.Number, , "AddNewPersonToMaster: " & Err.Description
End Sub

' =============================================
' @description Обновляет существующего сотрудника, сравнивая поля.
'              Записывает изменения в History.
' @param   rsBuffer [DAO.Recordset] Запись из буфера (текущая).
' @param   rsMaster [DAO.Recordset] Запись из мастера (найдена, редактируется).
' @param   strPersonUID [String] Личный номер для логирования.
' @return  [Long] Количество измененных полей.
' =============================================
Private Function UpdateExistingPerson(rsBuffer As DAO.Recordset, rsMaster As DAO.Recordset, strPersonUID As String) As Long
    On Error GoTo ErrorHandler

    Dim lngChanges As Long
    Dim varOldVal As Variant
    Dim varNewVal As Variant

    lngChanges = 0

    rsMaster.Edit

    ' --- СРАВНЕНИЕ RankName (Звание) ---
    varOldVal = Nz(rsMaster!RankName, "")
    varNewVal = Nz(rsBuffer!Rank_Raw, "")
    If varOldVal <> varNewVal Then
        rsMaster!RankName = varNewVal
        Call LogChange(strPersonUID, "RankName", CStr(varOldVal), CStr(varNewVal))
        lngChanges = lngChanges + 1
    End If

    ' --- СРАВНЕНИЕ WorkStatus (Статус) ---
    varOldVal = Nz(rsMaster!WorkStatus, "")
    varNewVal = Nz(rsBuffer!WorkStatus_Raw, "")
    If varOldVal <> varNewVal Then
        rsMaster!WorkStatus = varNewVal
        Call LogChange(strPersonUID, "WorkStatus", CStr(varOldVal), CStr(varNewVal))
        lngChanges = lngChanges + 1
    End If

    ' --- СРАВНЕНИЕ PosCode (Код должности) ---
    varOldVal = Nz(rsMaster!PosCode, "")
    varNewVal = Nz(rsBuffer!PosCode_Raw, "")
    If varOldVal <> varNewVal Then
        rsMaster!PosCode = varNewVal
        Call LogChange(strPersonUID, "PosCode", CStr(varOldVal), CStr(varNewVal))
        lngChanges = lngChanges + 1
    End If

    ' Обновляем также другие поля (без логирования, т.к. они не критичны для отслеживания)
    If Nz(rsMaster!FullName, "") <> Nz(rsBuffer!FullName_Raw, "") Then
        rsMaster!FullName = Nz(rsBuffer!FullName_Raw, "")
    End If
    If Nz(rsMaster!PosName, "") <> Nz(rsBuffer!PosName_Raw, "") Then
        rsMaster!PosName = Nz(rsBuffer!PosName_Raw, "")
    End If

    ' Обновляем дату актуальности
    rsMaster!LastUpdated = Now()

    rsMaster.Update

    UpdateExistingPerson = lngChanges

    Exit Function

ErrorHandler:
    Debug.Print "Ошибка UpdateExistingPerson: " & Err.Description
    If rsMaster.EditMode <> dbEditNone Then rsMaster.CancelUpdate
    UpdateExistingPerson = 0
End Function

' =============================================
' @description Записывает изменение в tbl_History_Log.
' @param   strPersonUID [String] Личный номер сотрудника.
' @param   strFieldName [String] Название поля (RankName, WorkStatus, PosCode и т.д.).
' @param   strOldValue [String] Старое значение.
' @param   strNewValue [String] Новое значение.
' =============================================
Public Sub LogChange(strPersonUID As String, strFieldName As String, strOldValue As String, strNewValue As String)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rsLog As DAO.Recordset

    Set db = CurrentDb
    Set rsLog = db.OpenRecordset("SELECT * FROM tbl_History_Log;", dbOpenDynaset)

    rsLog.AddNew
    rsLog!PersonUID = strPersonUID
    rsLog!ChangeDate = Now()
    rsLog!FieldName = strFieldName
    rsLog!OldValue = strOldValue
    rsLog!NewValue = strNewValue
    rsLog.Update

    rsLog.Close
    Set rsLog = Nothing
    Set db = Nothing

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка LogChange: " & Err.Description
    If Not rsLog Is Nothing Then
        If rsLog.EditMode <> dbEditNone Then rsLog.CancelUpdate
        rsLog.Close
        Set rsLog = Nothing
    End If
    Set db = Nothing
End Sub

' =============================================
' @description Преобразует текстовое значение в Long с обработкой ошибок.
' @param   varInput [Variant] Входное значение (обычно Text из Buffer).
' @return  [Long] Число или 0 при ошибке.
' =============================================
Private Function ConvertToLong(varInput As Variant) As Long
    On Error Resume Next

    If IsNull(varInput) Or varInput = "" Then
        ConvertToLong = 0
        Exit Function
    End If

    ConvertToLong = CLng(varInput)

    If Err.Number <> 0 Then
        Debug.Print "Ошибка ConvertToLong для значения: " & CStr(varInput)
        ConvertToLong = 0
        Err.Clear
    End If
End Function

' =============================================
' @description Преобразует текстовое значение в Date с обработкой ошибок.
' @param   varInput [Variant] Входное значение (обычно Text из Buffer).
' @return  [Date] Дата или Null при ошибке.
' =============================================
Private Function ConvertToDate(varInput As Variant) As Variant
    On Error Resume Next

    If IsNull(varInput) Or varInput = "" Then
        ConvertToDate = Null
        Exit Function
    End If

    ConvertToDate = CDate(varInput)

    If Err.Number <> 0 Then
        Debug.Print "Ошибка ConvertToDate для значения: " & CStr(varInput)
        ConvertToDate = Null
        Err.Clear
    End If
End Function
