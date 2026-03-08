# PROJECT CONTEXT: StaffState (Дозор)

**Последнее обновление:** 2026-03-08 (Tech Debt Refactor — Safe Init, UI Decoupling, Import Self-Heal)
**Разработчик:** Кержаев Евгений
**Тип проекта:** MS Access VBA Personnel Management System

### Текущее состояние (Current State)
- **Version 0.15 (Tech Debt Refactor Wave 1) — COMPLETED:** Выполнен первый пакет техдолг-рефакторинга поверх существующих Phase 19–31. `mod_App_Init.InitializeApp` теперь вызывает `InitDatabaseStructure True`; инициализация больше не делает destructive recreate для `tbl_Import_Buffer`, а выравнивает схему через `mod_Schema_Manager.EnsureBufferStructure` и `SyncMasterStructure`. Каноническая модель полей Master/Buffer расширена и централизована в `mod_Schema_Manager` (`GetAllowedMasterFields`, `GetAllowedBufferFields`, `GetCanonicalFieldType`). В `uf_Search` debug-лог поиска переведён на режим `LogLevel=DEBUG`, чтобы убрать постоянную hot-path запись в `debug_search.log`.
- **Version 0.15a (Service/UI Decoupling — Stage 1) — COMPLETED:** `mod_Maintenance_Logic` и верхний уровень `mod_Import_Logic` больше не показывают UI-сообщения напрямую для backup / validation export / health check / clear log / import summary. Добавлены result-based API (`RunDataHealthCheckResult`, `ExportValidationLogResult`, `CreateDatabaseBackupResult`, `ClearValidationLogResult`, `ImportExcelDataResult`), а формы `uf_Dashboard` и `uf_Settings` берут на себя показ сообщений и confirmation-диалогов.
- **Version 0.15b (Import Recovery & Diagnostics) — COMPLETED:** `mod_Analysis_Logic.RunFullSyncProcess` теперь показывает точную причину сбоя импорта в блоке `Import details`, вместо общего `FAILED or CANCELED`. `SelectExcelFile` трактует `ImportFolderPath = N/A` как пустое значение без ложного warning. Для регрессии с пропавшими профилями добавлен self-heal: `mod_Schema_Manager.CreateImportProfilesTable` вызывается перед auto-detect профиля в импорте и при загрузке `uf_Settings`, что автоматически восстанавливает default profiles 1/2/3.
- **Version 0.15c (Import Wizard UI Decoupling — Stage 2) — COMPLETED:** Интерактивный import wizard больше не вызывает `AskUserYesNo`/`InputBox` внутри `mod_Import_Logic`. Сервис возвращает `RequiresUserAction`, `ActionType`, `ActionMessage`, `ExcelField`, `SuggestedFieldName`, `ProfileID`, а `uf_Dashboard` управляет циклом решений пользователя и повторным запуском импорта с тем же Excel-файлом. Добавлены decision helpers: `CreateImportDecisionStore`, `SetSkipImportDecision`, `SetRestoreImportDecision`, `SetMapImportFieldDecision`.
- **Phase 24 (uf_Settings Tab UI & Logic Sync) — COMPLETED:** Форма настроек переведена на Tab Control (`tabSettings` в `uf_Settings.bas`): **pgGeneral** («Основные настройки») — название организации, папка импорта, автопроверка, уровень логов; **pgMapping** («Маппинг импорта») — список маппинга, добавление/удаление связей, «Восстановить по умолчанию»; **pgMaintenance** («Обслуживание») — резервная копия, очистка журнала проверки, «Запустить проверку данных». В `uf_Settings.cls`: загрузка настроек обёрнута в `Nz()` для всех вызовов `GetSetting` в `LoadSettingsFromStorage` (устранена ошибка «Save failed: Invalid use of Null»); при переключении на вкладку «Маппинг импорта» вызывается `tabSettings_Change` → `RefreshMappingList` для актуального списка.
- **Phase 23 (Person Card UI Overhaul — Tab-Based) — COMPLETED:** Рефакторинг `uf_PersonCard` с использованием Tab Control (`tabMain`) для лучшей организации данных. **Header (всегда видим):** txtFullName, txtPersonUID, txtRank, txtPosition, txtStatus + lblInactiveWarning. **Tab 1 — pgService (Служба):** PersonnelDivision, PersonnelDivision1, StaffPosition, VUS, SalaryGrade, EmployeeGroup. **Tab 2 — pgContract (Контракт):** ContractType, ContractKind, ContractStartDate, ContractEndDate, DismissalDate, EmploymentStatus + txtContractRemains (автоматический расчёт оставшегося времени). **Tab 3 — pgPersonal (Личные данные):** BirthDate, EmployeeAge, Gender, MaritalStatus, ChildrenCount, Nationality, Citizenship, Address. **Tab 4 — pgLogistics (Снабжение):** BootSize, HeadSize, SizeJacket, SizePants, SizeByHeight. **Tab 5 — pgBanking (Банк):** BankAccountNumber, BankKey, BankControlKey, Payee. **Tab 6 — pgHistory (Машина времени):** все фильтры истории (cboFilterHistory, txtDateFrom, txtDateTo, btnResetDates) и lstHistory перенесены на отдельную вкладку, lstHistory расширен на полную высоту/ширину вкладки. **Status-Based Styling:** при WorkStatus содержит "Dismissed" или "Уволен" — txtFullName.ForeColor = vbRed; при IsActive = False — lblInactiveWarning показывает "СОТРУДНИК НЕАКТИВЕН" красным цветом. DAO для загрузки данных; Option Explicit; английские комментарии; Cyrillic строки через ChrW для совместимости с Windows-1251. **Phase 23-Polish:** исправлена кодировка в `SetupHistoryFilters` (использованы английские технические названия: "All", "RankName" и т.д.); улучшен SQL в `ApplyHistoryFilter` для читаемого отображения даты и описания изменений (Old -> New); настроены `lstHistory.ColumnCount = 4` и `ColumnWidths = "2cm;3cm;4cm;4cm"`; удалён временный модуль `mod_Tests_Phase22.bas`.
- **Phase 22 (Advanced Export with Dynamic Headers) — COMPLETED:** Экспорт результатов поиска с динамическими русскими заголовками из `tbl_Import_Mapping`. Рефакторинг `uf_Search.cls`: создана публичная функция `GetSearchSQL(bAddTop50)` для получения SQL запроса; `GetCurrentFilterText()` сделан публичным для доступа из модуля экспорта; `PerformSearch()` переписан для вызова `GetSearchSQL(True)`; `btnExportExcel_Click()` упрощён для вызова `mod_Export_Logic.ExportFullSearchToExcel()`. В `mod_Export_Logic.bas`: реализована `ExportFullSearchToExcel()` — экспорт всех результатов поиска без TOP 50; `GetHeaderFromMapping()` загружает маппинг из `tbl_Import_Mapping` (ProfileID=1) в static Scripting.Dictionary (lazy load, кешируется на сессию); цепочка fallback для заголовков: Mapping → `GetFieldFriendlyName()` → `Replace("_", " ")`. Excel форматирование: Bold headers, AutoFilter, FreezePanes, Borders, AutoFit, date columns в формате dd.mm.yyyy. Интеграция с MCP AccessDB для проверки структуры БД при разработке.
- **Phase 21 (Universal Mapping Engine & UI) — COMPLETED:** Dynamic Mapping UI implemented in uf_Settings, English core schema expanded with Address and BirthDate_Text fields, Partial Import Protection (skipping empty fields) is active. Универсальный маппинг через `tbl_Import_Mapping` + `tbl_Import_Profiles`; заголовки Excel сопоставляются с английскими полями; UI маппинга (список по ExcelHeader ASC, «Поле в базе» с SourceID, сортировка A–Z, LimitToList=No); опечатки в сиде исправлены.
- **Phase 21-3 (COMPLETED):** В сиде маппинга исправлена опечатка «Лииуинный номер»→«Личный номер» (PersonUID). Удаление маппинга PersonUID — предупреждение и подтверждение через `AskUserYesNo` (при «Да» удаление выполняется). `ReSeedMapping` явно очищает маппинг профиля 1 перед повторным сидированием. Расширение схемы из формы настроек: при добавлении связи в маппинг проверяется наличие поля в `tbl_Personnel_Master` (`FieldExists`); при отсутствии вызывается `AddNewFieldToSchema` (добавление колонки LONGTEXT в Master и Buffer) и выводится сообщение «Поле [имя] создано в структуре базы данных.»; комбобокс «Поле в базе» — `LimitToList = No` для ввода новых имён полей.
- **Системные настройки:** таблица `tbl_Settings` реализована и создаётся при инициализации (идемпотентно). API: `GetSetting` / `SetSetting` в `mod_Maintenance_Logic`. Пользовательские настройки (например, `AutoCheckEnabled`) сохраняются в `tbl_Settings`.
- **Settings Manager UI (uf_Settings):** форма настроек полностью реализована (Phase 13, Phase 24). **Tab-based layout (Phase 24):** Tab Control `tabSettings` с тремя вкладками — «Основные настройки», «Маппинг импорта», «Обслуживание». Постоянные настройки (Название организации, Папка импорта, Автопроверка, Уровень логирования) — на вкладке «Основные настройки»; чтение/запись через `mod_Maintenance_Logic.GetSetting`/`SetSetting`; загрузка с защитой от Null (`Nz()` в `LoadSettingsFromStorage`); при переключении на вкладку «Маппинг импорта» список маппинга обновляется (`tabSettings_Change`).
- **Проверка целостности данных:** таблица `tbl_Validation_Log` реализована; проверки (дубликаты, сироты, будущие даты) выполняются через `RunDataHealthCheck` в `mod_Maintenance_Logic`.
- **Центральный модуль:** `mod_Maintenance_Logic` — единая точка входа для settings / health-check / maintenance logic; UI-диалоги для этих операций постепенно выносятся в формы через result-based API.
- **Dashboard (Phase 12, Phase 15, Phase 17, Phase 18):** на панели управления — кнопка Health Check с экспортом отчёта об ошибках в Excel; кнопка "Changes Report" формирует аудит-отчёт за период; кнопка **"Поиск дубликатов"** открывает uf_Search в режиме дубликатов; **Phase 18:** метрики (Total/Active/Errors) в lblTotalCount, lblActiveCount, lblErrorCount; кнопка **"Штатный срез"** формирует отчёт текущего состава в Excel.
- **Аудит-отчёт (Phase 15):** модуль `mod_Reports_Logic` с `GenerateAuditReport` — экспорт истории изменений в Excel (JOIN с `tbl_Personnel_Master`, форматирование, автофильтр, закрепление областей).
- **Автоматическая проверка после синхронизации:** при включённой настройке `AutoCheckEnabled` проверка целостности данных выполняется автоматически после завершения процесса Full Update (импорт → синхронизация).

---

## 🎯 Назначение проекта

**StaffState (Дозор)** — система мониторинга и учета персонала на базе MS Access.

### Основные задачи:
- Учет личного состава (ФИО, звания, должности, личные данные)
- История изменений всех полей сотрудников (аудит-лог)
- Импорт данных из Excel
- Интерактивная "Машина времени" для просмотра истории изменений

---

## 📊 Структура базы данных

**Ядро схемы (Phase 20):** все поля в `tbl_Personnel_Master` и `tbl_Import_Buffer` используют **английские имена**; импорт из Excel с русскими заголовками выполняется через систему маппинга `tbl_Import_Mapping` (см. ниже).

### Таблицы:

#### 1. `tbl_Personnel_Master` (Основная таблица персонала)
**Ключевое поле:** `PersonUID` (Личный номер). **Все поля — английские имена.**

**Основные поля:** PersonUID, FullName, RankName, PosName, PosCode, WorkStatus, BirthDate, IsActive, SourceID, LastUpdated.

**Личные данные (English):** EmployeeAge, Gender, MaritalStatus, ChildrenCount, Nationality, Citizenship, DismissalDate.

**Контракт и сроки:** ContractType, ContractKind, ContractStartDate, ContractEndDate, ContractYears, ContractMonths.

**Должность (детали):** StaffPosition, StaffPosition1, Position, VUS, SalaryGrade, PersonnelDivision, PersonnelDivision1, EmploymentStatus, EmployeeGroup.

**Мероприятия:** EventType, EventReason, ValidFromDate, ValidToDate, OrderNumber, OrderDate_Text, OrderNumber1, OrderDate_LS, PositionOrderIssuer.

**Банковские данные:** BankAccountNumber, BankKey, BankControlKey, Payee.

**Размеры одежды:** BootSize, HeadSize (и др. при расширении).

**Прочие (расширяемые из UI):** Address и любые поля, добавленные через форму настроек (маппинг импорта → «Поле в базе» с новым именем → автоматическое добавление колонки в Master и Buffer).

#### 2. `tbl_History_Log` (Журнал изменений)
**Ключевое поле:** `LogID` (AUTOINCREMENT)

- `PersonUID` — личный номер сотрудника
- `ChangeDate` — дата/время изменения
- `FieldName` — название поля (например, "RankName")
- `OldValue` — старое значение
- `NewValue` — новое значение

#### 3. `tbl_Import_Buffer` (Буфер импорта)
Временная таблица для хранения данных из Excel перед синхронизацией. Заполняется **только** колонками, сопоставленными в `tbl_Import_Mapping` (английские имена полей).

#### 4. `tbl_Import_Mapping` (Маппинг импорта) ✅ Phase 19
- **ProfileID** — профиль (например 1).
- **ExcelHeader** — заголовок из Excel (кириллица или латиница).
- **TargetField** — целевое английское поле в Buffer/Master.
Импорт в `mod_Import_Logic.RunDynamicImport` сопоставляет заголовки Excel с `ExcelHeader` (нормализация UCase+Trim) и пишет данные в соответствующие `TargetField`. Создание/заполнение: `mod_Schema_Manager.CreateImportMappingTable`, `SeedImportMappingProfile1`.

#### 5. `tbl_Import_Profiles` (Профили импорта) ✅ Phase 19
- **ProfileID** (PK), **ProfileName**, **IdStrategy** (например 'UID'). Создаётся через `mod_Schema_Manager.CreateImportProfilesTable`; по умолчанию профиль 1.

#### 6. `tbl_Import_Meta` (Метаданные импорта)
История импортов (дата, файл, количество записей).

#### 7. `tbl_Settings` (Системные настройки) ✅ Phase 11
- `SettingKey` — ключ (PK, Text 50)
- `SettingValue` — значение (Text 255)
- `SettingGroup` — группа (Text 50)
- `Description` — описание (Text 255)  
Создаётся через `mod_Schema_Manager.CreateSettingsTable`; доступ — `GetSetting` / `SetSetting` в `mod_Maintenance_Logic`.

#### 8. `tbl_Validation_Log` (Журнал проверки целостности) ✅ Phase 11
- `LogID` — автоинкремент (PK)
- `RecordID` — Long
- `TableName` — Text 50
- `ErrorType` — Text 50 (Duplicate, Orphan, FutureDate)
- `ErrorMessage` — Text 255
- `CheckDate` — Date/Time  
Создаётся через `mod_Schema_Manager.CreateValidationLogTable`; заполняется при `RunDataHealthCheck`.

---

## 🖥️ Формы (User Interface)

### 1. `uf_Dashboard` (Панель управления)
**Статус:** Базовая реализация + Phase 12 (Health Check) + Phase 15 (Audit Report)

Главная форма приложения с навигацией:
- Кнопка "Поиск сотрудника" → открывает `uf_Search`
- Кнопка **"Поиск дубликатов"** (Phase 17) → открывает `uf_Search` в режиме дубликатов (список записей с одинаковыми FullName и BirthDate); при повторном нажатии форма переоткрывается с актуальными данными
- Кнопка "Health Check" → ручная проверка целостности данных; при наличии ошибок — запрос на экспорт отчёта в Excel
- Поля **Start Date** и **End Date** + кнопка "Changes Report" (Phase 15) → генерация аудит-отчёта за выбранный период (даты берутся с формы, без дублирующих диалогов)
- Кнопки "Import", "Analyze", "Full Update", "Create Indexes", "Open Log"
- Кнопка "Отчеты" → (планируется)

### 1a. `uf_Settings` (Настройки системы) ✅ Phase 13, Phase 21, Phase 24
**Статус:** Полностью реализована (Tab-based UI)

**Tab Control `tabSettings` (Phase 24):** три вкладки — «Основные настройки», «Маппинг импорта», «Обслуживание».

- **Вкладка «Основные настройки» (pgGeneral):** **txtOrganizationName** (`OrganizationName`), **txtImportFolderPath** (`ImportFolderPath`), **chkAutoCheckEnabled** (`AutoCheckEnabled`), **cboLogLevel** (`LogLevel`: DEBUG/INFO/ERROR). Кнопки "Сохранить" и "Отмена" внизу формы (запись через `SetSetting`). Загрузка с `Nz()` для всех `GetSetting` — устранена ошибка «Invalid use of Null».
- **Вкладка «Маппинг импорта» (pgMapping):** список **lstMapping** по `ExcelHeader ASC`; **txtExcelHeader**, **cboTargetField** (Value List, LimitToList=No); кнопки «Добавить связь», «Удалить связь», «Восстановить по умолчанию». При переключении на эту вкладку вызывается `tabSettings_Change` → `RefreshMappingList`. Phase 21-3: при добавлении связи — проверка `FieldExists`, при отсутствии поля — `AddNewFieldToSchema`; удаление маппинга PersonUID — предупреждение через `AskUserYesNo`.
- **Вкладка «Обслуживание» (pgMaintenance):** кнопки «Создать резервную копию», «Очистить журнал проверки», «Запустить проверку данных» (Phase 14 + Health Check).
- Стабильный экспорт формы в VCS (без ControlSource у контролов).

---

### 2. `uf_Search` (Поиск сотрудников) ✅ Phase 17
**Статус:** Полностью реализована и оптимизирована

#### Функционал:
- **Умный поиск** по ФИО, личному номеру, SourceID
- **Автопоиск** при вводе (события: Change, KeyUp)
- **Минимальная длина:** начинает поиск после ввода ≥2 символов
- **Ограничение результатов:** TOP 50 записей
- **Двойной клик** на результате → открывает карточку сотрудника
- **Режим дубликатов (Phase 17):** при открытии с OpenArgs `MODE=DUPLICATES` — список строится напрямую из `tbl_Personnel_Master` (группы по FullName + BirthDate с Count>1); показывается сообщение «Дубликаты не найдены.» или «Найдены дубликаты: групп = X, записей = Y.»; кнопка «Export to Excel» включается при наличии строк и выгружает дубликаты в .xlsx (PersonUID, FullName, BirthDate, RankName, PosName, IsActive)

#### Элементы управления:
- `txtFilter` — поле ввода поискового запроса (в режиме дубликатов отключено)
- `lstResults` — список результатов (4 колонки: PersonUID, FullName, RankName, PosName)
- `btnSearch` — кнопка поиска (Enter также работает)
- `btnClear` — очистить результаты
- `btnExportExcel` — экспорт в Excel (в обычном режиме — результаты поиска; в режиме дубликатов — список дубликатов в отдельный .xlsx)

#### Технические особенности:
- SQL-запрос с LIKE для частичного совпадения
- Защита от SQL-инъекций (экранирование апострофов)
- Принудительное обновление списка (Requery)
- Обработка событий: Change, KeyUp, KeyDown

#### Файлы:
- `forms/uf_Search.bas` — дизайн формы
- `forms/uf_Search.cls` — код формы

---

### 3. `uf_PersonCard` (Карточка сотрудника) ✅✅
**Статус:** Полностью реализована с Tab Control, "Машиной времени" и системой синонимов (Phase 23)

#### Функционал:
- **Tab-Based UI:** 6 вкладок для лучшей организации данных
- **Header (всегда видим):** ФИО, личный номер, звание, должность, статус + предупреждение о неактивности
- **Status-Based Styling:** визуальное выделение уволенных и неактивных сотрудников
- **Машина времени (Tab 6):** фильтрация истории изменений по типу поля и диапазону дат
- **Автоматический расчёт:** оставшегося времени контракта (Tab 2)
- **Система умных синонимов:** 170+ вариантов русских названий полей
- **Автофильтрация при вводе** в комбобоксе

#### Элементы управления:
**Хедер (всегда видим):**
- `txtFullName` — ФИО (крупный шрифт, цвет меняется по статусу)
- `txtPersonUID` — личный номер
- `txtRank` — звание
- `txtPosition` — должность
- `txtStatus` — статус работы
- `lblInactiveWarning` — предупреждение "СОТРУДНИК НЕАКТИВЕН" (красный, жирный)

**Tab 1 — pgService (Служба):**
- `txtPersonnelDivision` — Раздел персонала
- `txtPersonnelDivision1` — Раздел персонала 1
- `txtStaffPosition` — Штатная должность
- `txtVUS` — ВУС
- `txtSalaryGrade` — Тарифный разряд
- `txtEmployeeGroup` — Группа сотрудников

**Tab 2 — pgContract (Контракт):**
- `txtContractType` — Вид контракта
- `txtContractKind` — Тип контракта
- `txtContractStartDate` — Дата начала контракта
- `txtContractEndDate` — Дата окончания контракта
- `txtDismissalDate` — Дата увольнения
- `txtEmploymentStatus` — Статус занятости
- `txtContractRemains` — Оставшееся время (автоматический расчёт)

**Tab 3 — pgPersonal (Личные данные):**
- `txtBirthDate` — Дата рождения
- `txtEmployeeAge` — Возраст
- `txtGender` — Пол
- `txtMaritalStatus` — Семейное положение
- `txtChildrenCount` — Количество детей
- `txtNationality` — Национальность
- `txtCitizenship` — Гражданство
- `txtAddress` — Адрес

**Tab 4 — pgLogistics (Снабжение):**
- `txtBootSize` — Размер сапог
- `txtHeadSize` — Охват головы
- `txtSizeJacket` — Размер куртки
- `txtSizePants` — Размер брюк
- `txtSizeByHeight` — Размер по росту

**Tab 5 — pgBanking (Банк):**
- `txtBankAccountNumber` — Номер счёта в банке
- `txtBankKey` — Ключ банка
- `txtBankControlKey` — Контрольный ключ
- `txtPayee` — Получатель

**Tab 6 — pgHistory (Машина времени):**
- `cboFilterHistory` — комбобокс с выбором типа поля (редактируемый!)
- `txtDateFrom` — дата начала периода
- `txtDateTo` — дата окончания периода
- `btnResetDates` — кнопка сброса фильтров (Х)
- `lstHistory` — список изменений (расширен на всю вкладку)

#### 🎯 Система умных синонимов `TranslateFieldName()`

**Функция автоматически переводит русские названия в технические:**

| Вводите по-русски | → Ищет в базе |
|------------------|---------------|
| звание, воинское, ранг | RankName |
| должность, позиция | PosName |
| статус, состояние | WorkStatus |
| фио, имя | FullName |
| дата рождения, др | Дата_рождения |
| возраст | Возраст_сотрудника |
| пол | Пол |
| семья, семейное, брак | Семейное_положение |
| дети, количество детей | Количество_детей |
| национальность, нация | Национальность |
| гражданство | Гражданство |
| вид контракта | Вид_контракта |
| тип контракта | Тип_контракта |
| начало контракта | Дата_начала_контракта |
| конец контракта | Дата_окончания_контракта |
| увольнение | Дата_увольнения |
| штатная должность, штатка | Штатная_должность |
| вус | ВУС |
| разряд, тарифный разряд | Тарифный_разряд |
| группа | Группа_сотрудников |
| банк, банковский счет | Номер_счета_в_банке |
| получатель | Получатель |
| сапоги, обувь | Размер_Сапог |
| голова, головной убор | Охват_головы |
| куртка | Размер_Куртки |
| брюки, штаны | Размер_Брюк |
| рост | Размер_по_Росту |
| приказ, дата приказа | Дата_приказа |
| номер приказа | Номер_приказа |
| мероприятие | Вид_мероприятия |
| причина | Причина_мероприятия |

**Всего поддерживается 170+ синонимов для всех полей базы!**

#### Примеры использования:
1. Введите **"звание"** → покажет историю изменений поля `RankName`
2. Введите **"семья"** → покажет историю `Семейное_положение`
3. Введите **"банк"** → покажет историю `Номер_счета_в_банке`
4. Введите **"сапоги"** → покажет историю `Размер_Сапог`

#### Технические особенности:
- **LimitToList = No** — разрешает вводить произвольные значения
- **Частичный поиск** — `LIKE '*input*'` вместо точного совпадения
- **Автофильтрация** — событие `Change` обновляет список при каждом вводе
- **Защита от SQL-инъекций** — экранирование апострофов
- **Принудительное обновление** — Requery для обхода кэширования Access
- **Обработка ошибок** — подавление ошибки 2467 при закрытии формы

#### Файлы:
- `forms/uf_PersonCard.bas` — дизайн формы
- `forms/uf_PersonCard.cls` — код формы (включая функцию TranslateFieldName)

---

## 📦 Модули

### 1. `mod_App_Init`
Инициализация приложения.

### 2. `mod_App_Logger`
Централизованное логирование ошибок и событий.

### 3. `mod_Import_Logic` ✅ Phase 19
Динамический импорт из Excel в `tbl_Import_Buffer` по маппингу `tbl_Import_Mapping` (ProfileID=1): сопоставление заголовков Excel (ExcelHeader) с английскими полями (TargetField), валидация наличия колонки, маппируемой на PersonUID; ранняя проверка существования и заполненности маппинга.

### 4. `mod_Export_Logic` ✅ Phase 22
Экспорт результатов поиска в Excel (Late Binding). **Phase 22:** `ExportFullSearchToExcel()` — экспорт всех результатов поиска без TOP 50 с динамическими русскими заголовками из `tbl_Import_Mapping`. Использует публичную функцию `GetSearchSQL(False)` из `uf_Search`. Маппинг загружается в static Scripting.Dictionary (lazy load). Fallback на `GetFieldFriendlyName()` → `Replace("_", " ")`. Excel форматирование: Bold headers, AutoFilter, FreezePanes, Borders, AutoFit, date columns → dd.mm.yyyy. Legacy: `ExportSearchToExcel(rs)` — выгрузка recordset в новую книгу; используется в `uf_Search` в обычном режиме поиска. В режиме дубликатов экспорт выполняется через временный QueryDef и `DoCmd.TransferSpreadsheet`. Заголовки через `GetFieldFriendlyName`, стиль: Bold, AutoFit, FreezePanes.

### 5. `mod_Schema_Manager` ✅ Phase 19, Phase 21-3
Управление структурой БД (проверка/создание таблиц). Phase 19: `CreateImportProfilesTable`, `CreateImportMappingTable`, `SeedImportMappingProfile1` (ChrW/CyrStr для кириллических заголовков, опечатка «Личный номер» для PersonUID исправлена), `CleanupImportBuffer`. **Phase 21-3:** `FieldExists(strTable, strField)` — проверка наличия поля в таблице (TableDefs.Fields); `AddNewFieldToSchema(strFieldName)` — добавление колонки LONGTEXT в `tbl_Personnel_Master` и `tbl_Import_Buffer` (имя в квадратных скобках); `ReSeedMapping` — очистка маппинга профиля 1 и повторное сидирование.

### 6. `mod_Analysis_Logic` ✅ Phase 16
Анализ и сравнение данных (определение изменений для `tbl_History_Log`).
- **Phase 16 — Sequential History Import (Smart Sync):** буфер читается с сортировкой `ORDER BY PersonUID_Raw, Nz(OrderDate_Raw,''), ID`. Для каждой записи буфера: проверка наличия UID в Master; при наличии — сравнение всех полей (Null/"" не считаются изменением), логирование отличий с контекстом даты приказа, немедленное обновление Master; при отсутствии — INSERT + лог "Added". Debug.Print: "Processing record X of Y for UID: [UID]".

### 7. `mod_Maintenance_Logic` ✅ Phase 11–14, Phase 18
Центральный модуль: **весь I/O настроек** и проверка целостности данных.
- **Phase 18:** `IsValidPersonUID(strUID)` — проверка формата (1–2 русские буквы + дефис + 6 цифр). `LogValidationError` — Public для вызова из mod_Analysis_Logic. `GetDashboardStats()` — возвращает Dictionary с TotalCount, ActiveCount, ErrorCount (DAO).
- **Настройки (единственная точка входа):** `GetSetting(key, [defaultValue])`, `SetSetting(key, value, [group])` — работа с `tbl_Settings`. Используются формой `uf_Settings` для OrganizationName, ImportFolderPath, AutoCheckEnabled, LogLevel и прочих ключей.
- **Проверка данных:** `RunDataHealthCheck([bSilentIfNoErrors])` — возвращает число ошибок; дубликаты (PersonUID, FullName+BirthDate), сироты в `tbl_History_Log`, даты в будущем; результаты пишутся в `tbl_Validation_Log`, итог — через `mod_UI_Helpers.ShowMessage`.
- **Экспорт (Phase 12):** `ExportValidationLogToExcel()` — экспорт `tbl_Validation_Log` в новую книгу Excel (Late Binding, заголовки: ID, RecordID, Table, Error Type, Message, Date; Bold, AutoFit, FreezePanes).
- **Phase 14:** `CreateDatabaseBackup()` — копирование БД в папку \Backups с меткой времени; `ClearValidationLog()` — очистка журнала проверки с подтверждением. Вызываются из uf_Settings (кнопки «Резервная копия», «Очистить логи»).

### 8. `mod_Reports_Logic` ✅ Phase 15, Phase 18
Модуль отчётов: централизованная генерация отчётов.
- **Аудит-отчёт:** `GenerateAuditReport([dtStart], [dtEnd])` — экспорт истории изменений за период в Excel. При вызове с датами (например, с Dashboard по полям Start/End Date) использует их; без параметров — запрос через InputBox. SQL JOIN `tbl_History_Log` + `tbl_Personnel_Master` (ФИО в отчёте). Late Binding Excel: заголовки, закрепление областей, автофильтр, формат дат, границы. Русский UI через ChrW().
- **Штатный срез (Phase 18):** `GenerateCurrentStaffReport()` — экспорт текущего состава из `tbl_Personnel_Master` в Excel (ORDER BY FullName). Колонки: PersonUID, FullName, RankName, PosName, WorkStatus, IsActive. Стиль как в аудит-отчёте (Bold, AutoFilter, FreezePanes, Borders).

### 9. `mod_UI_Helpers`
Вспомогательные UI-функции: `ShowMessage(msg, [msgType])`, `AskUserYesNo(msg, [title])`. Используются в Dashboard (Health Check — запрос на экспорт отчёта), Maintenance (итоги проверки), Settings и др.

### 10. `mod_Fix_Startup`
Утилита: `FixStartupForm()` — установка стартовой формы в uf_Dashboard (при сбое настроек запуска).

---

## 🔧 Технические стандарты

### Кодирование:
- **VBA файлы (.bas, .cls):** Windows-1251 (CP1251)
- **Документация (.md, .json):** UTF-8
- **Комментарии в коде:** ТОЛЬКО на английском (ASCII)
- **Пользовательские строки:** на русском (Cyrillic CP1251)

### Соглашения о коде:
- `Option Explicit` — обязательно
- Обработка ошибок: `On Error GoTo ErrorHandler` в каждой процедуре
- Типы данных: `Long` (не Integer), `LongPtr` для WinAPI
- Объекты: `Set obj = Nothing` в блоке выхода
- Позднее связывание: `CreateObject()` для внешних библиотек
- **Техническое логирование (Phase 11):** все сообщения в `Debug.Print` и логере — только на английском (избежание проблем с кодировкой в Immediate Window)
- **DAO:** в процедурах с DAO использовать `Dim db As DAO.Database: Set db = CurrentDb` для стабильности

### Комментарии JSDoc-стиль:
```vba
' =============================================
' @author Кержаев Евгений
' @description Short description of function
' @param varName [Type] Description
' @return [Type] Description
' =============================================
```

---

## 📝 История версий

### Версия 0.14 (2026-02-08) — uf_Settings Tab UI & Null-Safe Load (Phase 24) ✅
**Phase 24 — FORCED UI ALIGNMENT (Settings Form)**
- ✅ **uf_Settings.bas:** Форма экспортирована с Tab Control `tabSettings` (Begin Tab / Begin Page). Три страницы: **pgGeneral** («Основные настройки»), **pgMapping** («Маппинг импорта»), **pgMaintenance** («Обслуживание»). Элементы размещены по вкладкам; кнопки Сохранить/Отмена общие для формы.
- ✅ **uf_Settings.cls — LoadSettingsFromStorage:** Все присвоения из `GetSetting` обёрнуты в `Nz(..., default)` для полей OrganizationName, ImportFolderPath, AutoCheckEnabled, LogLevel — устранена ошибка «Save failed: Invalid use of Null» при пустых/отсутствующих значениях в `tbl_Settings`.
- ✅ **tabSettings_Change:** При переключении на вкладку «Маппинг импорта» (Value = 1) вызывается `RefreshMappingList` для актуального списка маппинга.
- ✅ **Form_Load:** Без изменений логики — загрузка настроек и первичное обновление списка маппинга при открытии формы.

### Версия 0.15 (2026-03-08) — Tech Debt Refactor Wave 1 ✅
**Safe Init + Service/UI Decoupling + Import Self-Heal**
- ✅ `mod_App_Init.InitializeApp` вызывает `InitDatabaseStructure True`; startup path больше не поднимает лишние MsgBox при нормальной инициализации.
- ✅ `InitDatabaseStructure` больше не удаляет `tbl_Import_Buffer`; вместо этого вызывает `mod_Schema_Manager.EnsureBufferStructure`, а Master выравнивается через `SyncMasterStructure`.
- ✅ `mod_Schema_Manager` стал источником канонической модели для schema evolution: добавлены `EnsureBufferStructure`, расширенные списки allowed fields для Buffer/Master и `GetCanonicalFieldType`.
- ✅ `uf_Search` пишет debug search log только при `LogLevel=DEBUG`, что убрало постоянную запись в hot path поиска.
- ✅ `mod_Maintenance_Logic` переведён на result-based maintenance API: `RunDataHealthCheckResult`, `ExportValidationLogResult`, `CreateDatabaseBackupResult`, `ClearValidationLogResult`, `FactoryResetDataResult`.
- ✅ `mod_Import_Logic` получил `ImportExcelDataResult` и перестал показывать итоговые import-error/import-success диалоги из сервисного слоя.
- ✅ `uf_Dashboard` и `uf_Settings` теперь сами показывают user-facing сообщения для import, health check, validation export, backup и clear log.
- ✅ `RunFullSyncProcess` выводит `Import details` при падении импорта, что резко упростило диагностику поломки маппинга/профилей.
- ✅ Добавлен self-heal для `tbl_Import_Profiles`: default profiles восстанавливаются автоматически перед auto-detect профиля и при открытии `uf_Settings`.
- ✅ Интерактивный import wizard отвязан от UI: `mod_Import_Logic` возвращает `RequiresUserAction` и action payload, а `uf_Dashboard` принимает решение пользователя и повторяет импорт с тем же Excel-файлом.
- ✅ Decision helpers для import wizard: `CreateImportDecisionStore`, `SetSkipImportDecision`, `SetRestoreImportDecision`, `SetMapImportFieldDecision`.

### Версия 0.13 (2026-02-08) — Person Card UI Overhaul + Encoding Fix (Phase 23) ✅
**Phase 23 — Tab-Based Person Card (FINAL)**
- ✅ **uf_PersonCard.cls refactoring:** Создана таб-ориентированная структура с 6 вкладками: pgService (Служба), pgContract (Контракт), pgPersonal (Личные данные), pgLogistics (Снабжение), pgBanking (Банк), pgHistory (История/Машина времени).
- ✅ **Header Section (Always Visible):** txtFullName, txtPersonUID, txtRank, txtPosition, txtStatus остаются в заголовке формы (не на вкладках).
- ✅ **Status-Based Styling:** при WorkStatus содержит "Dismissed" или "Уволен" — txtFullName.ForeColor = vbRed; при IsActive = False — lblInactiveWarning показывает "СОТРУДНИК НЕАКТИВЕН" красным цветом.
- ✅ **Tab 1 — pgService (Служба):** PersonnelDivision, PersonnelDivision1, StaffPosition, VUS, SalaryGrade, EmployeeGroup.
- ✅ **Tab 2 — pgContract (Контракт):** ContractType, ContractKind, ContractStartDate, ContractEndDate, DismissalDate, EmploymentStatus + txtContractRemains (автоматический расчёт оставшегося времени контракта).
- ✅ **Tab 3 — pgPersonal (Личные данные):** BirthDate, EmployeeAge, Gender, MaritalStatus, ChildrenCount, Nationality, Citizenship, Address.
- ✅ **Tab 4 — pgLogistics (Снабжение):** BootSize, HeadSize, SizeJacket, SizePants, SizeByHeight.
- ✅ **Tab 5 — pgBanking (Банк):** BankAccountNumber, BankKey, BankControlKey, Payee.
- ✅ **Tab 6 — pgHistory (Машина времени):** все фильтры истории (cboFilterHistory, txtDateFrom, txtDateTo, btnResetDates) и lstHistory перенесены на отдельную вкладку; lstHistory расширен на полную высоту/ширину вкладки.
- ✅ **Phase 23-Polish:** Исправлена кодировка в `SetupHistoryFilters` (использованы английские технические названия: "All", "RankName", "PosName", "WorkStatus", "FullName", "BirthDate"); улучшен SQL в `ApplyHistoryFilter` для читаемого отображения даты (Format dd.mm.yyyy) и описания изменений (OldValue -> NewValue); настроены `lstHistory.ColumnCount = 4` и `ColumnWidths = "2cm;3cm;4cm;4cm"`; удалён временный модуль `mod_Tests_Phase22.bas`.
- ✅ **Code Quality:** Option Explicit, DAO для загрузки данных, английские комментарии, Cyrillic строки через ChrW для совместимости с Windows-1251.

### Версия 0.12 (2026-02-07) — Advanced Export with Dynamic Headers (Phase 22) ✅
**Phase 22 — MCP-Enhanced Advanced Export**
- ✅ **uf_Search.cls refactoring:** Создана публичная функция `GetSearchSQL(bAddTop50 As Boolean)` для получения SQL запроса. `GetCurrentFilterText()` сделан публичным для доступа из модуля экспорта. `PerformSearch()` переписан для вызова `GetSearchSQL(True)`. `btnExportExcel_Click()` упрощён для вызова `mod_Export_Logic.ExportFullSearchToExcel()`.
- ✅ **mod_Export_Logic.bas — ExportFullSearchToExcel():** Экспорт всех результатов поиска (без TOP 50) с динамическими русскими заголовками. Заголовки берутся из `tbl_Import_Mapping` через Late Binding Scripting.Dictionary. Fallback: `mod_UI_Helpers.GetFieldFriendlyName()` → `Replace("_", " ")`. Excel форматирование: Bold headers, AutoFilter, FreezePanes, Borders, AutoFit, date columns в формате dd.mm.yyyy.
- ✅ **Маппинг из базы данных:** Функция `GetHeaderFromMapping()` создаёт Static dictionary из `tbl_Import_Mapping` (ProfileID=1) — загружается только один раз за сессию для оптимизации.
- ✅ **Интеграция с MCP:** Использование MCP AccessDB сервера для проверки структуры таблиц и содержимого маппинга при разработке.

### Версия 0.11 (2026-02-07) — Mapping Engine & English Core (Phase 19–20) ✅
**Phase 19 — Universal Mapping Engine**
- ✅ Таблицы `tbl_Import_Profiles` (ProfileID, ProfileName, IdStrategy) и `tbl_Import_Mapping` (ProfileID, ExcelHeader, TargetField). Создание и сидирование в `mod_Schema_Manager`: `CreateImportProfilesTable`, `CreateImportMappingTable`, `SeedImportMappingProfile1` (русские заголовки через CyrStr/ChrW).
- ✅ Импорт в `mod_Import_Logic.RunDynamicImport` только по маппингу: сопоставление заголовков Excel с `ExcelHeader`, запись в английские `TargetField`; ранняя проверка наличия таблицы и строк маппинга; понятные сообщения при отсутствии PersonUID с перечислением заголовков Excel.

**Phase 20 — English Core**
- ✅ Ядро схемы переведено на английские имена полей в `tbl_Personnel_Master` и `tbl_Import_Buffer` (EmployeeAge, Gender, MaritalStatus, ContractType, OrderNumber, BootSize, HeadSize и др.). Импорт из Excel с русскими заголовками работает через `tbl_Import_Mapping`.

**Phase 21 — Mapping UI & polish (100% COMPLETED)**
- ✅ uf_Settings: маппинг — список по `ORDER BY ExcelHeader ASC`; cboTargetField с **SourceID**, список полей по алфавиту (BUFFER_FIELDS_SORTED).
- ✅ mod_Schema_Manager: в `SeedImportMappingProfile1` опечатка «Штатус»→«Статус» (WorkStatus) исправлена (CyrStr 1064→1057). Существующую ошибочную запись пользователь удаляет вручную через UI.
- ✅ uf_Search: из `PerformSearch` удалены все `Debug.Print`.

**Phase 21-3 (2026-02-07) — Mapping typo, UI flexibility, schema expansion**
- ✅ mod_Schema_Manager: в `SeedImportMappingProfile1` опечатка «Лииуинный номер»→«Личный номер» (PersonUID, CyrStr). `ReSeedMapping` — явная очистка таблицы маппинга перед сидированием. Добавлены `FieldExists(strTable, strField)` и `AddNewFieldToSchema(strFieldName)` (ALTER TABLE … ADD COLUMN […] LONGTEXT в Master и Buffer).
- ✅ uf_Settings: удаление маппинга PersonUID — предупреждение и подтверждение через `AskUserYesNo`; при добавлении связи проверка `FieldExists("tbl_Personnel_Master", TargetField)`; при отсутствии поля — `AddNewFieldToSchema` и сообщение «Поле [имя] создано в структуре базы данных.»; cboTargetField — `LimitToList = No` для ввода новых имён полей.

### Версия 0.10 (2026-02-02) — Data Validation Hardening & Dashboard Metrics (Phase 18)
**Phase 18 — Strict PersonUID validation and dashboard analytics**
- ✅ **IsValidPersonUID** в `mod_Maintenance_Logic`: формат 1 или 2 русские буквы + дефис + ровно 6 цифр (Ю-111114, МТ-123114). Регистронезависимо.
- ✅ **Интеграция в синхронизацию:** в `mod_Analysis_Logic.SyncBufferToMaster` при невалидном UID — запись в `tbl_Validation_Log` (InvalidPersonUID), запись пропускается.
- ✅ **GetDashboardStats():** возвращает Dictionary с TotalCount, ActiveCount, ErrorCount (DAO, tbl_Personnel_Master, tbl_Validation_Log).
- ✅ **Dashboard:** три метки lblTotalCount, lblActiveCount, lblErrorCount; подпрограмма RefreshMetrics(); вызов в Form_Load, Form_Activate, после Health Check и Full Update. Ошибки — красный цвет при ErrorCount > 0, иначе зелёный.
- ✅ **GenerateCurrentStaffReport()** в `mod_Reports_Logic`: экспорт tbl_Personnel_Master в Excel (ORDER BY FullName), колонки PersonUID, FullName, RankName, PosName, WorkStatus, IsActive; стиль как в аудит-отчёте (Bold, AutoFilter, FreezePanes, Borders).
- ✅ Кнопка **"Штатный срез"** (btnSnapshotReport) на uf_Dashboard.

### Версия 0.9 (2026-02-01) — Duplicate Finder (Phase 17)
**Phase 17 — Find duplicates**
- ✅ **Кнопка на Dashboard:** «Поиск дубликатов» открывает uf_Search с OpenArgs `MODE=DUPLICATES`; при уже открытой uf_Search форма закрывается и открывается заново, чтобы Form_Load применил режим.
- ✅ **Режим дубликатов в uf_Search:** список строится только из `tbl_Personnel_Master` (подзапрос по FullName + BirthDate, HAVING Count(*)>1); Requery списка; сообщение пользователю: «Дубликаты не найдены.» или «Найдены дубликаты: групп = X, записей = Y.».
- ✅ **Экспорт дубликатов:** кнопка Export to Excel в режиме дубликатов включается при непустом списке; экспорт в .xlsx через временный QueryDef и DoCmd.TransferSpreadsheet (колонки: PersonUID, FullName, BirthDate, RankName, PosName, IsActive). Обычный режим поиска и его экспорт без изменений.

### Версия 0.8 (2026-02-01) — Smart Synchronization
**Phase 16 — Sequential History Import**
- ✅ **Умная синхронизация:** Реорганизован алгоритм импорта в `mod_Analysis_Logic`. Теперь система обрабатывает записи в хронологическом порядке (`PersonUID` -> `Дата_приказа` -> `ID`).
- ✅ **Полная реконструкция истории:** Если в одном файле импорта у сотрудника несколько изменений, система последовательно фиксирует каждое из них в `tbl_History_Log`, не пропуская промежуточные состояния.
- ✅ **Стабильность данных:** Улучшена обработка `Null` значений и логирование процесса синхронизации в Immediate Window.

### Версия 0.7 (2026-02-01)
**Phase 15 — Reporting & Auditing (Audit Report)**
- ✅ **Модуль отчётов:** Создан `mod_Reports_Logic` для централизованной генерации отчётов.
- ✅ **Экспорт аудита:** Реализована функция `GenerateAuditReport`. Позволяет выгружать историю изменений за любой период в Excel.
- ✅ **Интеграция данных:** Использование SQL JOIN для объединения `tbl_History_Log` с `tbl_Personnel_Master` (вывод ФИО вместо сухих UID).
- ✅ **Профессиональный экспорт:** Автоматическое форматирование Excel (закрепление областей, автофильтр, стили, форматы дат).
- ✅ **Интеграция с Dashboard:** Кнопка "Changes Report" использует даты из полей Start Date и End Date на форме (без дублирующих InputBox). Поддержка вызова с параметрами или без (fallback на InputBox).

### Версия 0.6 (2026-02-01)
**Phase 14 — Backup & Cleanup Tools**
- ✅ **Система бэкапов:** Добавлен `CreateDatabaseBackup` в `mod_Maintenance_Logic`. Автоматическое создание папки `\Backups` и копирование текущей БД с меткой времени (`YYYYMMDD_HHMMSS`).
- ✅ **Очистка данных:** Добавлена процедура `ClearValidationLog` с подтверждением пользователя.
- ✅ **UI Обслуживание:** В форму `uf_Settings` добавлен раздел "Обслуживание базы данных" с кнопками для запуска бэкапа и очистки логов.
- ✅ **Late Binding:** Использование `FileSystemObject` через Late Binding для исключения проблем с референсами.

### Версия 0.5 (2026-02-01)
**Phase 11 Reboot - Data Integrity & System Settings**

- ✅ **Settings API**
  - Таблица `tbl_Settings` (SettingKey PK, SettingValue, SettingGroup, Description).
  - `mod_Maintenance_Logic.GetSetting(key, [defaultValue])` — чтение настройки.
  - `mod_Maintenance_Logic.SetSetting(key, value, [group])` — запись/обновление (Insert or Update).
- ✅ **Data Integrity (Health Check)**
  - Таблица `tbl_Validation_Log` (LogID, RecordID, TableName, ErrorType, ErrorMessage, CheckDate).
  - `RunDataHealthCheck()`: проверки дубликатов (PersonUID, FullName+BirthDate в `tbl_Personnel_Master`), сирот (PersonUID в `tbl_History_Log` без записи в Master), дат в будущем (ChangeDate > Now()); все находки пишутся в `tbl_Validation_Log`, итог — через `ShowMessage`.
- ✅ **Reboot-подход:** надёжная работа с DAO (`CurrentDb` в каждой процедуре), идемпотентное создание таблиц (`CreateSettingsTable` / `CreateValidationLogTable` с ранним выходом при существующей таблице), единая точка входа `InitializeApp` → `InitDatabaseStructure`.
- ✅ **Техническое логирование:** все `Debug.Print` и сообщения логера — только на английском.

**Phase 12 (2026-02-01) — Dashboard Integration & Automation**

- ✅ **Health Check на Dashboard:** кнопка "Health Check" для ручной проверки целостности; при наличии ошибок — запрос (через `mod_UI_Helpers.AskUserYesNo`) на экспорт отчёта в Excel.
- ✅ **Экспорт журнала проверки:** `ExportValidationLogToExcel()` в `mod_Maintenance_Logic` — экспорт `tbl_Validation_Log` в новую книгу Excel (Late Binding, стиль как в `mod_Export_Logic`: Bold headers, AutoFit, FreezePanes).
- ✅ **Автопроверка после Full Update:** в `RunFullSyncProcess` после синхронизации при настройке `AutoCheckEnabled = True` вызывается `RunDataHealthCheck(bSilentIfNoErrors:=True)`; сообщение показывается только при наличии ошибок.
- ✅ **Настройка AutoCheckEnabled:** сохраняется в `tbl_Settings`; управляет автоматическим запуском проверки целостности после процесса импорта/синхронизации.

**Phase 13 (2026-02-01) — Settings Manager UI Final**

- ✅ **Форма uf_Settings:** финализирован UI менеджера настроек; постоянные настройки (Organization Name, Import Path, Auto-check, Log Level) управляются через форму.
- ✅ **Исправления ControlSource:** привязка к данным убрана; значения загружаются и сохраняются программно через `GetSetting`/`SetSetting` в `mod_Maintenance_Logic`, стабильный экспорт формы в VCS.
- ✅ **Единая точка I/O настроек:** `mod_Maintenance_Logic` подтверждён как модуль, обрабатывающий весь ввод/вывод настроек (форма только вызывает его API).

### Версия 0.4 (2026-01-23)
**Phase 7 - Performance & Usability Improvements**

#### Session 1 - Critical Performance:
- ✅ **Task 1: Performance Indexes** (COMPLETED)
  - Создана функция `mod_App_Init.CreatePerformanceIndexes()`
  - 4 индекса для ускорения операций:
    - `idx_PersonUID` на `tbl_Personnel_Master` (UNIQUE)
    - `idx_FullName` на `tbl_Personnel_Master`
    - `idx_History_PersonUID` на `tbl_History_Log`
    - `idx_History_ChangeDate` на `tbl_History_Log`
  - Добавлена кнопка "Создать индексы" на форму `uf_Dashboard`
  - **Результат:** Импорт 20,000 записей ускорен в ~30 раз, поиск мгновенный (<0.1 сек)
  
- ✅ **Task 2: PersonUID Validation** (COMPLETED in Phase 18) — см. раздел «🔍 Валидация PersonUID» ниже.

### Версия 0.3 (2026-01-22)
**Реализовано:**
- ✅ Полная оптимизация формы `uf_Search`:
  - Минимальная длина поиска: ≥2 символа
  - Ограничение результатов: TOP 50
  - Автопоиск на событиях Change и KeyUp
  - SQL-защита от инъекций
  
- ✅ Улучшена форма `uf_PersonCard`:
  - Редактируемый комбобокс для фильтра истории
  - Частичный поиск по названиям полей (LIKE)
  - Автофильтрация при вводе
  
- ✅ **Система умных синонимов `TranslateFieldName()`:**
  - 170+ русских синонимов для всех полей базы
  - Автоматический перевод: "звание" → "RankName"
  - Поддержка всех категорий: личные данные, контракты, одежда, банк
  - Регистронезависимый поиск
  
- ✅ Обновлен список в комбобоксе с популярными русскими вариантами

### Версия 0.2 (2026-01-22)
**Реализовано:**
- ✅ Форма `uf_Search` с автопоиском
- ✅ Форма `uf_PersonCard` с фильтрацией истории
- ✅ Базовая "Машина времени" (фильтры по дате и типу поля)

### Версия 0.1 (2026-01-20)
**Реализовано:**
- ✅ Структура таблиц: `tbl_Personnel_Master`, `tbl_History_Log`
- ✅ Базовые модули: Logger, Init, Import, Schema Manager
- ✅ Форма `uf_Dashboard`

---

## 🎯 Планы развития

### Phase 24 — 100% COMPLETED (2026-02-08)
- ✅ **uf_Settings Tab UI:** Tab Control `tabSettings` с вкладками pgGeneral, pgMapping, pgMaintenance. Загрузка настроек с `Nz()` в `LoadSettingsFromStorage`; обновление списка маппинга при переключении на вкладку «Маппинг импорта» (`tabSettings_Change`).

### Phase 23 — 100% COMPLETED (2026-02-08)
- ✅ **Phase 23 (Person Card UI Overhaul — Tab-Based):** Рефакторинг `uf_PersonCard` с использованием Tab Control (`tabMain`) для лучшей организации данных. Header (всегда видим): txtFullName, txtPersonUID, txtRank, txtPosition, txtStatus + lblInactiveWarning. Tab 1 — pgService (Служба): PersonnelDivision, PersonnelDivision1, StaffPosition, VUS, SalaryGrade, EmployeeGroup. Tab 2 — pgContract (Контракт): ContractType, ContractKind, ContractStartDate, ContractEndDate, DismissalDate, EmploymentStatus + txtContractRemains (автоматический расчёт оставшегося времени). Tab 3 — pgPersonal (Личные данные): BirthDate, EmployeeAge, Gender, MaritalStatus, ChildrenCount, Nationality, Citizenship, Address. Tab 4 — pgLogistics (Снабжение): BootSize, HeadSize, SizeJacket, SizePants, SizeByHeight. Tab 5 — pgBanking (Банк): BankAccountNumber, BankKey, BankControlKey, Payee. Tab 6 — pgHistory (Машина времени): все фильтры истории и lstHistory на отдельной вкладке. Status-Based Styling: txtFullName.ForeColor = vbRed при увольнении, lblInactiveWarning для неактивных сотрудников.
- ✅ **Phase 23-Polish (Encoding & History Tab):** Исправлена кодировка в `SetupHistoryFilters` (английские технические названия: "All", "RankName", "PosName", "WorkStatus", "FullName", "BirthDate"); улучшен SQL в `ApplyHistoryFilter` для читаемого отображения даты (Format dd.mm.yyyy) и описания изменений (OldValue -> NewValue); настроены `lstHistory.ColumnCount = 4` и `ColumnWidths = "2cm;3cm;4cm;4cm"`; удалён временный модуль `mod_Tests_Phase22.bas`.
- **Key Deliverables:** Tab-based UI organization; Status-based visual feedback; Automatic contract remaining time calculation; Expanded history view with encoding-safe implementation; Improved code structure with separate loading functions for each tab.

### Phase 22 — 100% COMPLETED (2026-02-08)
- ✅ **Phase 22 (Advanced Export with Dynamic Headers):** Рефакторинг `uf_Search.cls` — публичные функции `GetSearchSQL()` и `GetCurrentFilterText()` для доступа из модуля экспорта. `ExportFullSearchToExcel()` в `mod_Export_Logic` — экспорт всех результатов поиска с динамическими русскими заголовками из `tbl_Import_Mapping`. Late Binding Scripting.Dictionary для маппинга (static, lazy load). Fallback на `GetFieldFriendlyName()`. Excel форматирование: Bold, AutoFilter, FreezePanes, Borders, AutoFit, date columns → dd.mm.yyyy.
- ✅ **Task 5: Search Results Export (COMPLETED):** Полный экспорт результатов поиска в Excel без TOP 50, с русскими заголовками из маппинга.
- **Key Deliverables:** Dynamic Mapping for Exports (using `tbl_Import_Mapping`); Decoupled SQL generation from uf_Search logic; UI Testing infrastructure (Python 3.14 + Selenium 3.141.0 + WinAppDriver); Automatic directory creation for Exports.

### Phase 21 — 100% COMPLETED (2026-02-07)
- ✅ **Phase 21 (Mapping Engine & UI):** `tbl_Import_Profiles`, `tbl_Import_Mapping`; импорт только по маппингу; сидирование профиля 1 (русские заголовки → английские поля). UI маппинга: список по ExcelHeader ASC; комбобокс «Поле в базе» с SourceID, сортировка A–Z; опечатка «Штатус»→«Статус» в SeedImportMappingProfile1 исправлена; Debug.Print убраны из PerformSearch (uf_Search).
- ✅ **Phase 21-3 (Mapping fixes & schema expansion):** опечатка «Личный номер» в сиде; предупреждение при удалении маппинга PersonUID; ReSeedMapping очищает таблицу; FieldExists, AddNewFieldToSchema; расширение схемы из формы настроек (LimitToList=No, автосоздание поля при добавлении связи).

### Phase 7 - В процессе реализации:
- ✅ Task 1: Performance Indexes (COMPLETED)
- ✅ Task 2: PersonUID Validation (COMPLETED in Phase 18) — проверка наличия колонки в Excel, фильтрация пустых значений, проверка дубликатов через Health Check, UNIQUE constraint на уровне БД; Phase 18: форматная валидация (1–2 русские буквы + дефис + 6 цифр) при синхронизации буфер→мастер
- ✅ Task 3: Changes Report (Excel export) — Phase 15
- ⏳ Task 4: Duplicate Import Detection
- ⏳ Task 5: Search Results Export
- 🔵 Task 6: History Archiving (LOW PRIORITY - Year 15+)
- 🔵 Task 7: Search Column Configuration (OPTIONAL)

### Будущие задачи:
1. **Отчеты и аналитика:**
   - Статистика по изменениям
   - ✅ Отчет "Кто изменился за период" (Task 3 — Phase 15)
   - Экспорт данных в Excel (Tasks 5 и др.)

2. **Улучшения интерфейса:**
   - Редактирование данных сотрудника
   - Массовые операции
   - ✅ Поиск дубликатов (Phase 17 — завершён; слияние/удаление дубликатов — возможное развитие в будущем)

3. **Долгосрочная оптимизация:**
   - Архивация старых записей (Task 6 - после 10+ лет эксплуатации)

---

## 🐛 Известные проблемы

### Решенные:
- ✅ Access кэширует RowSource списков → решено через `Requery`
- ✅ Событие Change не всегда срабатывает → добавлено KeyUp
- ✅ `.Text` не доступен вне фокуса → fallback на `.Value`
- ✅ Точный поиск по полю → заменен на LIKE для частичного совпадения
- ✅ Неудобные технические названия → система синонимов TranslateFieldName

### Текущие:
- Нет валидации данных при ручном редактировании
- Нет защиты от одновременного редактирования

---

## 📚 Ключевые функции и процедуры

### Форма uf_PersonCard

#### `TranslateFieldName(strUserInput As String) As String`
**Назначение:** Переводит русские синонимы в технические названия полей.

**Примеры:**
- `TranslateFieldName("звание")` → "RankName"
- `TranslateFieldName("семья")` → "Семейное_положение"
- `TranslateFieldName("банк")` → "Номер_счета_в_банке"

**Использование:**
```vba
strFieldFilter = TranslateFieldName(Me.cboFilterHistory.Text)
```

#### `ApplyHistoryFilter()`
**Назначение:** Применяет фильтры к таблице истории изменений.

**Фильтры:**
- По типу поля (с поддержкой синонимов)
- По диапазону дат (От/До)
- Комбинация фильтров

#### `Form_Load()`
**Назначение:** Загружает данные сотрудника и инициализирует фильтры.

**Параметры:** `OpenArgs` — PersonUID сотрудника

### Форма uf_Search

#### `PerformSearch(Optional bSilent As Boolean)`
**Назначение:** Выполняет поиск сотрудников по введенному запросу.

**Параметры:**
- `bSilent` — если True, не показывать сообщения об ошибках

**Особенности:**
- Минимальная длина: 2 символа
- TOP 50 результатов
- Поиск по ФИО, PersonUID, SourceID
- Защита от SQL-инъекций

### mod_Maintenance_Logic (Phase 11, Phase 12, Phase 13)

#### `GetSetting(key, [defaultValue]) As Variant`
**Назначение:** Чтение значения настройки из `tbl_Settings` по ключу. При отсутствии ключа возвращает `defaultValue`. Перед обращением вызывается `CreateSettingsTable` при необходимости.

#### `SetSetting(key, value, [group])`
**Назначение:** Запись или обновление настройки в `tbl_Settings`. При отсутствии записи — INSERT, иначе UPDATE. Группа по умолчанию — "General". Настройки вроде `AutoCheckEnabled` сохраняются в `tbl_Settings`.

#### `RunDataHealthCheck([bSilentIfNoErrors]) As Long`
**Назначение:** Проверка целостности данных. Возвращает число найденных ошибок. Дубликаты (PersonUID, FullName+BirthDate в `tbl_Personnel_Master`), сироты (PersonUID в `tbl_History_Log` без записи в Master), даты в будущем (ChangeDate > Now()). Результаты пишутся в `tbl_Validation_Log`, итог выводится через `ShowMessage` (если `bSilentIfNoErrors = True` и ошибок 0 — сообщение не показывается).

#### `ExportValidationLogToExcel() As Boolean`
**Назначение (Phase 12):** Экспорт записей из `tbl_Validation_Log` в новую книгу Excel (Late Binding). Колонки: ID, RecordID, Table, Error Type, Message, Date. Стиль: Bold headers, AutoFit, FreezePanes.

### mod_Reports_Logic (Phase 15)

#### `GenerateAuditReport([dtStart], [dtEnd])`
**Назначение:** Экспорт истории изменений персонала за период в новую книгу Excel. Опциональные параметры — даты начала и конца периода; при вызове с Dashboard передаются из полей Start Date / End Date; без параметров — запрос через InputBox. SQL: JOIN `tbl_History_Log` и `tbl_Personnel_Master` (в отчёте ФИО, личный номер, поле, дата изменения, старое/новое значение). Форматирование: заголовок и подзаголовок периода, закрепление первых 3 строк, автофильтр по заголовкам, формат даты dd.mm.yyyy hh:mm, границы и выравнивание.

---

## 🔍 Валидация PersonUID (Task 2 - Phase 7 PARTIAL)

### Что уже реализовано:

#### 1. **Проверка при импорте** (`mod_Import_Logic.bas`)
```130:142:mod_Import_Logic.bas
' --- VALIDATE: PersonUID field must exist ---
If Len(strPersonUIDExcelName) = 0 Then
    MsgBox "CRITICAL: Excel file has no PersonUID column!" & vbCrLf & vbCrLf & _
           "Found columns:" & vbCrLf & strAllColumns & vbCrLf & vbCrLf & _
           "Looking for column containing: 'number', 'UID', 'PersonUID'", _
           vbCritical, "Import Error"
    RunDynamicImport = False
    Exit Function
End If
```
- ✅ Проверяется наличие колонки PersonUID в Excel-файле
- ✅ При отсутствии колонки импорт прерывается с критическим сообщением
- ✅ SQL-запрос фильтрует пустые значения: `WHERE [PersonUID] IS NOT NULL`

#### 2. **Фильтрация пустых значений** (`mod_Analysis_Logic.bas`)
```58:58:mod_Analysis_Logic.bas
If strUID <> "" Then
```
- ✅ Пропускаются записи с пустым PersonUID при синхронизации буфера с Master

#### 3. **Проверка дубликатов** (`mod_Maintenance_Logic.bas`)
```181:188:mod_Maintenance_Logic.bas
strSQL = "SELECT PersonUID, COUNT(*) AS Cnt FROM tbl_Personnel_Master GROUP BY PersonUID HAVING COUNT(*)>1;"
Do While Not rs.EOF
    LogValidationError 0, "tbl_Personnel_Master", "Duplicate", "Duplicate PersonUID: " & Nz(rs!PersonUID, "")
    lngDup = lngDup + 1
    rs.MoveNext
Loop
```
- ✅ Проверка дубликатов через Health Check
- ✅ Дубликаты логируются в `tbl_Validation_Log`
- ✅ Автоматический запуск после импорта (если включена настройка `AutoCheckEnabled`)

#### 4. **UNIQUE constraint на уровне БД** (`tbl_Personnel_Master.sql`)
```2:2:tbldefs/tbl_Personnel_Master.sql
[PersonUID] VARCHAR (50) CONSTRAINT [idx_PersonUID] UNIQUE CONSTRAINT [PK_Person] PRIMARY KEY UNIQUE NOT NULL,
```
- ✅ Защита на уровне БД от дубликатов
- ✅ Автоматический отклонение INSERT с дублирующим PersonUID

### Phase 18 — форматная валидация PersonUID (реализовано)

- ✅ **IsValidPersonUID(strUID)** в `mod_Maintenance_Logic`: паттерн 1 или 2 русские буквы (А-Я, Ё) + дефис + ровно 6 цифр. Регистронезависимо.
- ✅ Валидация при синхронизации буфер→мастер: невалидные UID логируются в `tbl_Validation_Log` (ErrorType InvalidPersonUID), запись не переносится в Master.

---

## 📂 Структура файлов проекта

```
StaffState.accdb.src/
├── .gitattributes
├── .gitignore
├── .spec/
│   ├── PROJECT_CONTEXT.md          ← Этот файл
│   ├── 007-performance-improvements.md
│   └── phase-19-20-universal-mapping-and-english-schema.md
├── forms/
│   ├── uf_Dashboard.bas/.cls       ← Панель управления
│   ├── uf_Settings.bas/.cls        ← Настройки системы, табы: Основные/Маппинг/Обслуживание (Phase 13, 24)
│   ├── uf_Search.bas/.cls          ← Поиск сотрудников
│   └── uf_PersonCard.bas/.cls      ← Карточка сотрудника + Машина времени
├── modules/
│   ├── mod_App_Init.bas
│   ├── mod_App_Logger.bas
│   ├── mod_Import_Logic.bas
│   ├── mod_Export_Logic.bas
│   ├── mod_Fix_Startup.bas
│   ├── mod_Schema_Manager.bas
│   ├── mod_Analysis_Logic.bas
│   ├── mod_Maintenance_Logic.bas   ← Phase 11–14, 18: Settings, Health Check, IsValidPersonUID, GetDashboardStats, Export, Backup
│   ├── mod_Reports_Logic.bas       ← Phase 15, 18: Audit Report, GenerateCurrentStaffReport (штатный срез)
│   └── mod_UI_Helpers.bas
├── tbldefs/
│   ├── tbl_Personnel_Master.sql/.xml
│   ├── tbl_History_Log.sql/.xml
│   ├── tbl_Import_Buffer.sql/.xml
│   ├── tbl_Import_Mapping.sql/.xml   ← Phase 19
│   ├── tbl_Import_Meta.sql/.xml
│   ├── tbl_Settings.sql/.xml
│   └── tbl_Validation_Log.sql/.xml
├── themes/
│   └── Тема Office.thmx
├── dbs-properties.json
├── documents.json
├── nav-pane-groups.json
├── project.json
├── vbe-project.json
├── vbe-references.json
└── vcs-options.json
```

---

## 💡 Советы по работе

### Для разработчиков:
1. **Всегда читайте .cursorrules** перед началом работы
2. **Используйте UTF-8 для .md файлов**, Windows-1251 для VBA
3. **Комментарии на английском**, строки пользователю на русском
4. **Тестируйте в Access** после каждого импорта

### Для пользователей:
1. **В поиске (uf_Search):** минимум 2 символа для начала поиска
2. **В фильтре истории:** можно вводить русские названия ("звание", "семья")
3. **Двойной клик** на результате поиска открывает карточку
4. **Кнопка X** сбрасывает все фильтры истории

---

## 🔗 Связанные ресурсы

- **.cursorrules** — конституция проекта (в корне)
- **Build.log / Export.log** — логи экспорта/импорта VCS
- **vbe-references.json** — используемые библиотеки

---

**Конец документа**  
Последнее обновление: 2026-02-08 (Phase 24 — uf_Settings tab UI & Null-safe load)
