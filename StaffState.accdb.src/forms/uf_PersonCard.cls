Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Option Explicit

' =============================================
' @form uf_PersonCard
' @description Карточка сотрудника с историей изменений
' =============================================

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    Dim strUID As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String

    ' 1. Получаем ID сотрудника из аргументов открытия
    ' (Мы передадим его из формы Поиска)
    If IsNull(Me.OpenArgs) Then
        MsgBox "Ошибка: Не передан ID сотрудника.", vbCritical
        DoCmd.Close
        Exit Sub
    End If

    strUID = Me.OpenArgs

    ' 2. Загружаем текущие данные (Шапка)
    Set db = CurrentDb
    strSQL = "SELECT * FROM tbl_Personnel_Master WHERE PersonUID = '" & strUID & "'"
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)

    If Not rs.EOF Then
        Me.txtFullName.Value = rs!FullName
        Me.txtPersonUID.Value = rs!PersonUID
        Me.txtRank.Value = rs!RankName
        Me.txtPosition.Value = rs!PosName
        Me.txtStatus.Value = rs!WorkStatus
    Else
        MsgBox "Сотрудник не найден в базе!", vbExclamation
        DoCmd.Close
        Exit Sub
    End If
    rs.Close

    ' 3. Загружаем Историю (Список)
    ' Показываем: Дата, Что менялось, Старое, Новое
    strSQL = "SELECT ChangeDate, FieldName, OldValue, NewValue " & _
             "FROM tbl_History_Log " & _
             "WHERE PersonUID = '" & strUID & "' " & _
             "ORDER BY ChangeDate DESC;"

    Me.lstHistory.RowSource = strSQL

    Exit Sub
ErrorHandler:
    MsgBox "Ошибка загрузки карточки: " & Err.Description, vbCritical
End Sub

' --- КНОПКА ЗАКРЫТЬ (Опционально) ---
Private Sub btnClose_Click()
    DoCmd.Close
End Sub
