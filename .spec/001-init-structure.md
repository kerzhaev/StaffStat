# Spec: Инициализация структуры БД и импорт

## 1. Описание задачи
Создать структуру таблиц в MS Access для приема данных из файла выгрузки. Реализовать SQL-скрипты (DDL) для генерации схемы.

## 2. Анализ данных (Mapping)
На основе файла "Фейковая выгрузка.csv" определены целевые поля:

| Excel Header (CSV) | Access Field | Type | Комментарий |
|---|---|---|---|
| Лицо (ID) | SourceID | Long | Внутренний ID источника (163676) |
| Личный номер | PersonUID | Text(20) | **PRIMARY KEY** (Ю-111111) |
| Воинское звание | RankName | Text(50) | Лейтенант |
| Лицо (ФИО) | FullName | Text(150) | Иванов А.А. |
| Дата рождения | BirthDate | Date | 04.04.1985 |
| Статус занятости | WorkStatus | Text(50) | активный |
| Штатная должность | PosCode | Text(20) | Код (16807987) |
| Должность | PosName | Memo/LongText | Длинное название подразделения |
| Дата приказа | OrderDate | Date | Дата приказа Л/С |
| Номер приказа | OrderNum | Text(50) | Номер приказа Л/С |

## 3. План реализации (Implementation Plan)

### Шаг 1: Модуль `mod_App_Init`
Создать процедуру `InitDatabaseStructure`, которая выполнит DDL запросы:

1. **Удаление/Пересоздание `tbl_Import_Buffer`**:
   - Поля создаются как `Short Text` (255), чтобы избежать ошибок импорта (Type Mismatch).
   - Индексов нет (для скорости вставки).
   - Все поля из CSV заголовка.

2. **Создание `tbl_Personnel_Master` (если нет)**:
   - `PersonUID` (PK, Text(20)).
   - Типизированные поля (Date, Long).
   - Поля мета-данных: `LastUpdated` (Date), `IsActive` (Boolean).

3. **Создание `tbl_History_Log`**:
   - `LogID` (AutoNumber, PK).
   - `PersonUID` (Text(20), FK).
   - `ChangeDate` (Date).
   - `FieldName` (Text(50)), `OldValue` (Memo), `NewValue` (Memo).

### Шаг 2: SQL DDL (Предварительный код)
```sql
CREATE TABLE tbl_Personnel_Master (
    PersonUID VARCHAR(20) CONSTRAINT PK_Person PRIMARY KEY,
    SourceID LONG,
    FullName VARCHAR(150),
    RankName VARCHAR(50),
    BirthDate DATETIME,
    WorkStatus VARCHAR(50),
    PosCode VARCHAR(20),
    PosName MEMO,
    OrderDate DATETIME,
    OrderNum VARCHAR(50),
    LastUpdated DATETIME,
    IsActive YESNO
);
```

## 4. Критерии приемки
- Процедура `InitDatabaseStructure` запускается без ошибок.
- Таблицы появляются в Access.
- Типы данных соответствуют ожидаемым.
- Все поля из CSV могут быть импортированы в буфер.
