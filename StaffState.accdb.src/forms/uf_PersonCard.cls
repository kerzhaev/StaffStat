Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =============================================
' @form uf_PersonCard
' @description Employee card with "Time Machine" (History filters)
' =============================================

Private m_strUID As String ' Form-level variable to store ID

' =============================================
' @description Translate user-friendly field names to technical database names
' @param strUserInput [String] User-entered field name (e.g., "звание", "должность")
' @return [String] Technical field name or original input if no match
' =============================================
Private Function TranslateFieldName(ByVal strUserInput As String) As String
    Dim strLower As String
    
    ' Convert to lowercase for case-insensitive comparison
    strLower = LCase(Trim(strUserInput))
    
    ' Return "Все" as-is (special case)
    If strLower = "все" Or strLower = "" Then
        TranslateFieldName = strUserInput
        Exit Function
    End If
    
    ' === COMPREHENSIVE FIELD NAME MAPPING ===
    Select Case strLower
        ' --- CORE FIELDS ---
        Case "фио", "имя", "полное имя"
            TranslateFieldName = "FullName"
        
        Case "звание", "воинское", "воинское звание", "ранг"
            TranslateFieldName = "RankName"
        
        Case "должность", "позиция", "пост"
            TranslateFieldName = "PosName"
        
        Case "код должности", "код"
            TranslateFieldName = "PosCode"
        
        Case "статус", "состояние", "рабочий статус"
            TranslateFieldName = "WorkStatus"
        
        Case "личный", "личный номер", "uid", "номер"
            TranslateFieldName = "PersonUID"
        
        Case "активный", "активен"
            TranslateFieldName = "IsActive"
        
        ' --- DATES ---
        Case "дата рождения", "рождение", "др"
            TranslateFieldName = "Дата_рождения"
        
        Case "дата приказа", "приказ дата"
            TranslateFieldName = "Дата_приказа"
        
        Case "номер приказа", "приказ номер"
            TranslateFieldName = "Номер_приказа"
        
        Case "дата увольнения", "увольнение"
            TranslateFieldName = "Дата_увольнения"
        
        Case "начало контракта", "дата начала контракта"
            TranslateFieldName = "Дата_начала_контракта"
        
        Case "конец контракта", "окончание контракта", "дата окончания контракта"
            TranslateFieldName = "Дата_окончания_контракта"
        
        Case "начало срока", "начало срока действия"
            TranslateFieldName = "Начало_срока_действия"
        
        Case "конец срока", "конец срока действия"
            TranslateFieldName = "Конец_срока_действия"
        
        ' --- PERSONAL INFO ---
        Case "возраст"
            TranslateFieldName = "Возраст_сотрудника"
        
        Case "пол"
            TranslateFieldName = "Пол"
        
        Case "семья", "семейное", "семейное положение", "брак"
            TranslateFieldName = "Семейное_положение"
        
        Case "дети", "количество детей"
            TranslateFieldName = "Количество_детей"
        
        Case "национальность", "нация"
            TranslateFieldName = "Национальность"
        
        Case "гражданство"
            TranslateFieldName = "Гражданство"
        
        ' --- CONTRACT INFO ---
        Case "вид контракта", "контракт вид"
            TranslateFieldName = "Вид_контракта"
        
        Case "тип контракта", "контракт тип"
            TranslateFieldName = "Тип_контракта"
        
        Case "срок договора год", "договор год"
            TranslateFieldName = "Срок_договора_год"
        
        Case "срок договора месяц", "договор месяц"
            TranslateFieldName = "Срок_договора_месяц"
        
        Case "группа сотрудников", "группа"
            TranslateFieldName = "Группа_сотрудников"
        
        ' --- EVENT INFO ---
        Case "вид мероприятия", "мероприятие"
            TranslateFieldName = "Вид_мероприятия"
        
        Case "причина мероприятия", "причина"
            TranslateFieldName = "Причина_мероприятия"
        
        ' --- POSITION DETAILS ---
        Case "штатная должность", "штатка"
            TranslateFieldName = "Штатная_должность"
        
        Case "вус", "военно-учетная специальность"
            TranslateFieldName = "ВУС"
        
        Case "тарифный разряд", "разряд"
            TranslateFieldName = "Тарифный_разряд"
        
        Case "раздел персонала", "раздел"
            TranslateFieldName = "Раздел_персонала"
        
        Case "статус занятости", "занятость"
            TranslateFieldName = "Статус_занятости"
        
        Case "единица расчета"
            TranslateFieldName = "Единица_расчета"
        
        Case "рассчитано по"
            TranslateFieldName = "Рассчитано_по"
        
        ' --- BANK INFO ---
        Case "банк", "банковский счет", "счет в банке", "номер счета"
            TranslateFieldName = "Номер_счета_в_банке"
        
        Case "ключ банка", "банковский ключ"
            TranslateFieldName = "Ключ_банка"
        
        Case "контрольный ключ", "контрольный банковский ключ"
            TranslateFieldName = "Контрольный_банковский_ключ"
        
        Case "получатель"
            TranslateFieldName = "Получатель"
        
        ' --- CLOTHING SIZES ---
        Case "сапоги", "размер сапог", "обувь", "размер обуви"
            TranslateFieldName = "Размер_Сапог"
        
        Case "голова", "охват головы", "размер головы", "головной убор"
            TranslateFieldName = "Охват_головы"
        
        Case "куртка", "размер куртки"
            TranslateFieldName = "Размер_Куртки"
        
        Case "брюки", "размер брюк", "штаны"
            TranslateFieldName = "Размер_Брюк"
        
        Case "рост", "размер по росту"
            TranslateFieldName = "Размер_по_Росту"
        
        ' --- RAW FIELDS (for technical users) ---
        Case "sourceid", "ид источника"
            TranslateFieldName = "SourceID"
        
        Case "orderdate", "дата приказа ориг"
            TranslateFieldName = "OrderDate"
        
        Case "ordernum", "номер приказа ориг"
            TranslateFieldName = "OrderNum"
        
        ' --- EXACT TECHNICAL NAMES (pass through as-is) ---
        Case "rankname", "posname", "workstatus", "fullname", "personuid", _
             "poscode", "birthdate", "isactive", "lastupdated"
            TranslateFieldName = strUserInput
        
        ' --- DEFAULT: Return original input (for partial matches or unrecognized) ---
        Case Else
            TranslateFieldName = strUserInput
    End Select
End Function

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String

    ' 1. Check incoming ID
    If IsNull(Me.OpenArgs) Then
        MsgBox "Error: Employee ID not provided.", vbCritical
        DoCmd.Close
        Exit Sub
    End If

    m_strUID = Me.OpenArgs

    ' 2. Load header (Master)
    Set db = CurrentDb
    strSQL = "SELECT * FROM tbl_Personnel_Master WHERE PersonUID = '" & m_strUID & "'"
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)

    If Not rs.EOF Then
        ' Remove ControlSource before assignment (if exists)
        On Error Resume Next
        Me.txtFullName.ControlSource = ""
        Me.txtPersonUID.ControlSource = ""
        Me.txtRank.ControlSource = ""
        Me.txtPosition.ControlSource = ""
        Me.txtStatus.ControlSource = ""
        On Error GoTo ErrorHandler

        ' Assign values (use simple assignment)
        Me.txtFullName = Nz(rs!FullName, "")
        Me.txtPersonUID = Nz(rs!PersonUID, "")
        Me.txtRank = Nz(rs!RankName, "")
        Me.txtPosition = Nz(rs!PosName, "")
        Me.txtStatus = Nz(rs!WorkStatus, "")
    Else
        MsgBox "Employee not found in database!", vbExclamation
        DoCmd.Close
        Exit Sub
    End If
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    ' 3. Set default filters
    On Error Resume Next
    Me.cboFilterHistory.ControlSource = ""
    ' Force correct RowSourceType (fix for Access caching bug)
    Me.cboFilterHistory.RowSourceType = "Value List"
    ' Provide comprehensive list of user-friendly field names
    Me.cboFilterHistory.RowSource = """Все"";" & _
                                    """Звание"";""Должность"";""Статус"";" & _
                                    """ФИО"";""Дата рождения"";""Возраст"";" & _
                                    """Пол"";""Семейное положение"";""Дети"";" & _
                                    """Национальность"";""Гражданство"";" & _
                                    """Вид контракта"";""Начало контракта"";""Конец контракта"";" & _
                                    """Дата увольнения"";""Группа сотрудников"";" & _
                                    """Штатная должность"";""ВУС"";""Тарифный разряд"";" & _
                                    """Банковский счет"";""Получатель"";" & _
                                    """Сапоги"";""Охват головы"";""Куртка"";""Брюки"";""Рост"";" & _
                                    """RankName"";""PosName"";""WorkStatus"""
    On Error GoTo ErrorHandler
    Me.cboFilterHistory = "Все"
    Me.txtDateFrom = Null
    Me.txtDateTo = Null

    ' 4. First history load
    ApplyHistoryFilter

    Exit Sub
ErrorHandler:
    MsgBox "Card load error: " & Err.Description & vbCrLf & "Error number: " & Err.Number, vbCritical
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not db Is Nothing Then Set db = Nothing
End Sub

' =============================================
' @description MAIN FILTERING ENGINE
' =============================================
Private Sub ApplyHistoryFilter()
    On Error GoTo ErrorHandler

    Dim strSQL As String
    Dim strWhere As String
    Dim strFieldFilter As String

    ' Base query
    ' NOTE: lstHistory is a ListBox; native grouping is not available.
    ' We display Date in a separate column, and include time in the description.
    strSQL = "SELECT " & _
             "ChangeDate AS SortChangeDate, " & _
             "FORMAT(ChangeDate, ""dd.mm.yyyy"") AS [Дата], " & _
             "FORMAT(ChangeDate, ""hh:nn"") & "" "" & BuildHistoryDescription([FieldName], [OldValue], [NewValue]) AS [Описание изменений] " & _
             "FROM tbl_History_Log"
    strWhere = " WHERE PersonUID = '" & m_strUID & "'"

    ' 1. Filter by Field Type (ComboBox) - allow custom input with partial match
    ' Try multiple methods to get current text
    On Error Resume Next
    strFieldFilter = ""
    
    ' Method 1: Try .Text (works during typing if cursor is in the field)
    strFieldFilter = Me.cboFilterHistory.Text
    
    ' Method 2: If failed, try .Value (works after selection or update)
    If Err.Number <> 0 Or IsNull(strFieldFilter) Or strFieldFilter = "" Then
        Err.Clear
        strFieldFilter = Nz(Me.cboFilterHistory.Value, "")
    End If
    
    ' Method 3: Check if control has focus and try Screen.ActiveControl
    If strFieldFilter = "" And Not Me.cboFilterHistory Is Nothing Then
        If TypeName(Screen.ActiveControl) = "ComboBox" Then
            If Screen.ActiveControl.Name = "cboFilterHistory" Then
                strFieldFilter = Screen.ActiveControl.Text
            End If
        End If
    End If
    
    On Error GoTo ErrorHandler
    
    strFieldFilter = Trim(strFieldFilter)
    
    ' TRANSLATE USER INPUT: Convert user-friendly names to technical names
    ' E.g., "Rank" -> "RankName", "Position" -> "PosName"
    strFieldFilter = TranslateFieldName(strFieldFilter)
    
    ' Escape single quotes for SQL safety
    strFieldFilter = Replace(strFieldFilter, "'", "''")
    
    ' Use LIKE for partial matching (e.g., "семе" will find "Семейное_положение")
    If strFieldFilter <> "" And strFieldFilter <> "Все" Then
        strWhere = strWhere & " AND FieldName LIKE '*" & strFieldFilter & "*'"
    End If

    ' 2. Filter by Date From
    If IsDate(Me.txtDateFrom) Then
        ' Access SQL requires date format #mm/dd/yyyy#
        strWhere = strWhere & " AND ChangeDate >= #" & Format(Me.txtDateFrom, "mm\/dd\/yyyy") & "#"
    End If

    ' 3. Filter by Date To
    ' Add 23:59:59 to capture entire end of day
    If IsDate(Me.txtDateTo) Then
        strWhere = strWhere & " AND ChangeDate <= #" & Format(Me.txtDateTo, "mm\/dd\/yyyy") & " 23:59:59#"
    End If

    ' Build and sort
    strSQL = strSQL & strWhere & " ORDER BY ChangeDate DESC;"

    ' Apply to list
    Me.lstHistory.RowSource = strSQL
    
    ' Force refresh (fix for Access caching)
    Me.lstHistory.Requery

    Exit Sub
ErrorHandler:
    ' Suppress errors in silent mode (during typing)
    If Err.Number <> 2467 Then ' Ignore "The expression you entered refers to an object that is closed"
        MsgBox "Filter error: " & Err.Description & " (Error: " & Err.Number & ")"
    End If
End Sub

' --- INTERFACE EVENTS ---

' Real-time filter as user types in ComboBox
Private Sub cboFilterHistory_Change()
    ' Apply filter immediately as user types custom field name
    ApplyHistoryFilter
End Sub

' Filter change in dropdown list
Private Sub cboFilterHistory_AfterUpdate()
    ApplyHistoryFilter
End Sub

' Date From input
Private Sub txtDateFrom_AfterUpdate()
    ApplyHistoryFilter
End Sub

' Date To input
Private Sub txtDateTo_AfterUpdate()
    ApplyHistoryFilter
End Sub

' Reset button (X)
Private Sub btnResetDates_Click()
    Me.txtDateFrom = Null
    Me.txtDateTo = Null
    Me.cboFilterHistory = "Все"
    ApplyHistoryFilter
End Sub

' Close button (if exists)
Private Sub btnClose_Click()
    DoCmd.Close
End Sub
