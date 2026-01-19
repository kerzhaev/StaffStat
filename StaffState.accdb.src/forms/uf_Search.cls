Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Option Explicit

' =============================================
' @form uf_Search
' @author Кержаев Евгений
' @description Форма поиска сотрудников
' =============================================

' --- КНОПКА ПОИСК ---
Private Sub btnSearch_Click()
    PerformSearch
End Sub

' --- КНОПКА СБРОС ---
Private Sub btnClear_Click()
    On Error Resume Next
    Me.txtFilter.Value = ""
    Me.lstResults.RowSource = ""
    Me.txtFilter.SetFocus
End Sub

' --- НАЖАТИЕ ENTER В ПОЛЕ ВВОДА ---
Private Sub txtFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Если нажат Enter (Код 13)
    If KeyCode = 13 Then
        PerformSearch
        KeyCode = 0 ' Глушим стандартный звук Enter
    End If
End Sub

' --- ДВОЙНОЙ КЛИК ПО СПИСКУ ---
Private Sub lstResults_DblClick(Cancel As Integer)
    On Error GoTo ErrorHandler

    Dim strUID As Variant ' Используем Variant, чтобы поймать Null

    ' Проверяем, есть ли выделение
    If Me.lstResults.ItemsSelected.Count = 0 Then Exit Sub

    ' Берем значение (оно берется из Присоединенного столбца = 1)
    strUID = Me.lstResults.Value

    ' ДИАГНОСТИКА: Если пусто - скажем об этом
    If IsNull(strUID) Or strUID = "" Then
        MsgBox "Внимание: ID сотрудника не считался из списка." & vbCrLf & _
               "Проверьте свойство 'Число столбцов' (должно быть 4) и 'Присоединенный столбец' (должно быть 1).", vbExclamation
        Exit Sub
    End If

    ' Если все ок - открываем карточку
    ' Передаем ID в OpenArgs (последний параметр)
    DoCmd.OpenForm "uf_PersonCard", , , , , acDialog, CStr(strUID)

    Exit Sub
ErrorHandler:
    MsgBox "Ошибка открытия: " & Err.Description
End Sub

' --- ЛОГИКА ПОИСКА ---
Private Sub PerformSearch()
    On Error GoTo ErrorHandler

    Dim strSQL As String
    Dim strFilter As String

    strFilter = Trim(Nz(Me.txtFilter.Value, ""))

    ' Формируем базовый SQL запрос
    ' ВАЖНО: В lstResults свойство "Тип источника строк" должно быть "Таблица/Запрос"
    strSQL = "SELECT PersonUID, FullName, RankName, PosName " & _
             "FROM tbl_Personnel_Master "

    ' Добавляем условия фильтрации
    If strFilter <> "" Then
        strSQL = strSQL & "WHERE (FullName LIKE '*" & strFilter & "*') " & _
                 "OR (PersonUID LIKE '*" & strFilter & "*') " & _
                 "OR (CStr(SourceID) LIKE '*" & strFilter & "*') "
    End If

    ' Сортировка
    strSQL = strSQL & "ORDER BY FullName;"

    ' Присваиваем SQL списку
    Me.lstResults.RowSource = strSQL

    ' Если ничего не найдено
    If Me.lstResults.ListCount = 0 Then
        MsgBox "Сотрудники не найдены.", vbExclamation
    End If

    Exit Sub
ErrorHandler:
    MsgBox "Ошибка поиска: " & Err.Description, vbCritical
End Sub
