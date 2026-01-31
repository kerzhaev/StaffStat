Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Option Explicit

' =============================================
' @form uf_Dashboard
' @author Кержаев Евгений
' @description Главная панель управления системой "Дозор"
' =============================================

' --- КНОПКА 1: ЗАГРУЗКА ---
Private Sub btnImport_Click()
    On Error GoTo ErrorHandler

    Me.lblStatus.Caption = "Выполняется импорт..."
    DoEvents ' Даем интерфейсу перерисоваться

    ' Вызов логики из модуля
    If mod_Import_Logic.ImportExcelData() Then
        Me.lblStatus.Caption = "Импорт завершен успешно."
    Else
        Me.lblStatus.Caption = "Импорт отменен или прошел с ошибкой."
    End If

    Exit Sub
ErrorHandler:
    MsgBox "UI error: " & Err.Description, vbCritical
    Me.lblStatus.Caption = "Ошибка."
End Sub

' --- КНОПКА 2: АНАЛИЗ ---
Private Sub btnAnalyze_Click()
    On Error GoTo ErrorHandler

    Me.lblStatus.Caption = "Выполняется анализ данных..."
    DoEvents

    ' Вызов логики анализа
    mod_Analysis_Logic.SyncBufferToMaster

    Me.lblStatus.Caption = "Анализ завершен. База актуальна."

    Exit Sub
ErrorHandler:
    MsgBox "UI error: " & Err.Description, vbCritical
    Me.lblStatus.Caption = "Ошибка."
End Sub

' --- КНОПКА 3: ЖУРНАЛ ---
Private Sub btnOpenLog_Click()
    On Error GoTo ErrorHandler

    Me.lblStatus.Caption = "Открываю журнал..."

    ' Открываем таблицу истории
    DoCmd.OpenTable "tbl_History_Log"

    Me.lblStatus.Caption = "Журнал открыт."

    Exit Sub
ErrorHandler:
    MsgBox "Unable to open log: " & Err.Description, vbCritical
End Sub

' --- КНОПКА 4: СОЗДАТЬ ИНДЕКСЫ (Phase 7) ---
Private Sub btnCreateIndexes_Click()
    On Error GoTo ErrorHandler
    
    Me.lblStatus.Caption = "Создание индексов производительности..."
    DoEvents
    
    ' Call index creation routine
    mod_App_Init.CreatePerformanceIndexes
    
    Me.lblStatus.Caption = "Индексы созданы. Система работает быстрее."
    
    Exit Sub
ErrorHandler:
    MsgBox "Index creation error: " & Err.Description, vbCritical
    Me.lblStatus.Caption = "Ошибка."
End Sub
