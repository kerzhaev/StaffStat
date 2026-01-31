# Spec: Дата выгрузки из файла (Export File Date)

## 1. Context & Goal

**Problem:** Сейчас дата изменений в `tbl_History_Log` и `LastUpdated` в Master задаётся как **момент загрузки файла в систему** (`Now()`). Файлы выгрузки из учётной программы могли быть сохранены давно (полгода, год, два года назад). Фактическая дата «на когда» данные актуальны — это **дата сохранения файла выгрузки**, т.е. когда файл был выгружен из учётной системы.

**Goal:** Использовать **дату модификации файла выгрузки** (дата сохранения) как дату изменений при логировании в History и при обновлении `LastUpdated` в Master. Загрузка в Access может происходить в любой момент позже.

## 2. User Stories

- [ ] Как пользователь, я загружаю старый файл выгрузки (например, годовой давности), запускаю анализ — чтобы в журнале изменений и в «Машине времени» фигурировала **дата выгрузки файла**, а не сегодняшняя дата.

## 3. Functional Requirements

### 3.1. Input
- Путь к выбранному Excel-файлу при импорте (уже есть в `ImportExcelData`).
- При синхронизации — данные из `tbl_Import_Meta` (см. ниже).

### 3.2. Process

1. **При импорте (mod_Import_Logic):**
   - После выбора файла (`SelectExcelFile`) и до очистки буфера:
     - Получить **дату последней модификации** файла через `FileSystemObject` (`.DateLastModified`).
   - После успешного `RunDynamicImport`:
     - Записать в **`tbl_Import_Meta`** одну строку: `ExportFileDate`, `ImportRunAt` (и при желании `SourceFilePath` для аудита).
   - Таблица `tbl_Import_Meta` — одна строка на «последний импорт»; при каждом новом импорте — UPDATE этой строки (или DELETE + INSERT).

2. **При синхронизации (mod_Analysis_Logic):**
   - В начале `SyncBufferToMaster`:
     - Прочитать `ExportFileDate` из `tbl_Import_Meta`.
     - Если записи нет или `ExportFileDate` пусто — **fallback на `Now()`** (обратная совместимость, старые импорты без метаданных).
   - Использовать `ExportFileDate` вместо `Now()`:
     - В `LogChange` — для `ChangeDate` в `tbl_History_Log`.
     - В Master — для `LastUpdated` при добавлении и обновлении записей.

3. **Получение даты файла:**
   - Late binding: `CreateObject("Scripting.FileSystemObject")`, затем `GetFile(path).DateLastModified`.
   - Обработка ошибок: если файл недоступен или путь неверный — считать дату неизвестной, использовать `Now()`.

### 3.3. Output
- В `tbl_History_Log.ChangeDate` — дата сохранения файла выгрузки (когда имеется).
- В `tbl_Personnel_Master.LastUpdated` — та же дата.
- Новая таблица `tbl_Import_Meta` с метаданными последнего импорта.

## 4. Data Model

### 4.1. `tbl_Import_Meta` (новая)

| Column         | Type     | Description                                    |
|----------------|----------|------------------------------------------------|
| ID             | Long     | PK, автоинкремент (одна строка, ID=1)         |
| ExportFileDate | DateTime | Дата модификации файла выгрузки               |
| ImportRunAt    | DateTime | Когда выполнен импорт в Access                |
| SourceFilePath | Text(255)| Путь к файлу (опционально, для аудита)        |

Стратегия: одна строка. При импорте — `UPDATE tbl_Import_Meta SET ...` или, если пусто, `INSERT`.

## 5. UI/UX

- Без изменений в формах. Импорт и анализ вызываются как раньше (Dashboard).
- При желании: в статусной строке Dashboard после импорта можно показывать «Дата выгрузки файла: DD.MM.YYYY» — опционально, не в scope первой итерации.

## 6. Files to Create / Modify

| File                       | Action                                                                 |
|----------------------------|------------------------------------------------------------------------|
| `mod_App_Init.bas`         | Добавить создание `tbl_Import_Meta` в `InitDatabaseStructure` (или в `mod_Schema_Manager` при первой миграции). |
| `mod_Import_Logic.bas`     | Получить `DateLastModified` по пути файла; после импорта — запись в `tbl_Import_Meta`. |
| `mod_Analysis_Logic.bas`   | Читать `ExportFileDate` из `tbl_Import_Meta`; передавать в `LogChange` и использовать для `LastUpdated`. |
| `tbldefs/tbl_Import_Meta.xml` | Добавить при использовании msaccess-vcs (структура таблицы).          |
| `.spec/PROJECT_CONTEXT.md` | Обновить «Текущее состояние» и историю после реализации.              |

## 7. Edge Cases & Fallbacks

- **Нет `tbl_Import_Meta` или нет строки:** использовать `Now()`.
- **Файл удалён/перемещён до импорта:** маловероятно (импорт идёт по выбранному файлу). Если при линковке ошибка — импорт прерывается, мета не обновляется.
- **Повторный «Анализ» без нового импорта:** используется та же `ExportFileDate` из последнего импорта — корректно.

## 8. Constraints

- MS Access 2010+, VBA 7.0. `Scripting.FileSystemObject` доступен по умолчанию.
- Late binding для FSO.
- Encoding: Windows-1251 для VBA, UTF-8 для spec.

## 9. Acceptance Criteria

- [ ] После импорта в `tbl_Import_Meta` есть строка с `ExportFileDate` = дата модификации выбранного файла.
- [ ] После «Анализ» записи в `tbl_History_Log` имеют `ChangeDate` = дата файла выгрузки (а не текущая дата).
- [ ] `LastUpdated` в Master для новых/обновлённых записей = дата файла выгрузки.
- [ ] Если `tbl_Import_Meta` пуста (старая БД или сбой) — используется `Now()`, без ошибок.
