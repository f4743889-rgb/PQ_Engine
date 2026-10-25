# TEST_PQ_ENGINE

> Файл створює макрос ExportPowerQueriesToRepo з таблиць книги (hlpProject, tblTableCatalog, hlp-таблиці, cfgQueryGroups). Вручну не редагувати: зміни вносяться в книгу.

## Правила структури книги

### Розміщення

1. Шапка кожної таблиці - не вище рядка 3. Рядок над шапкою порожній; ще вище (рядок 1) - підпис таблиці, який ставить макрос UpdateHelp.
2. Таблиці ростуть униз і вправо, тому кілька таблиць на одному аркуші ставлять лише поруч, ніколи одна під одною.
3. Таблиця з запиту Power Query (q_, pq_) - одна на аркуші. Аркуш бажано називати так само, як запит.
4. Кожна таблиця займає щонайменше 3 колонки (місце для підпису). Між сусідніми таблицями - щонайменше одна порожня колонка.
5. Нижче і правіше таблиці нічого не розміщують: ні формул, ні приміток.

### Назви

1. Префікс задає роль таблиці: tbl - ручне введення і налаштування; q_ - дані призначення (результат запиту); pq_ - діагностика і звіти Engine; hlp - довідка; test_ - тести (лише тестова книга).
2. Назви таблиць унікальні і завжди з префіксом. Довідка до tblX за замовчуванням - hlpX.

### Довідка

1. Аркуш HELP: tblTableCatalog, hlpProject (ця таблиця) і всі hlp-таблиці - поруч, шапки в рядку 3.
2. Призначення таблиці описують лише в tblTableCatalog[Призначення]. Правила заповнення колонок - у hlp-таблиці, вказаній у tblTableCatalog[Довідка].

### Каталог

1. Колонки каталогу, які підлягають заповненню: Призначення; Підпис (TRUE - показувати підпис над таблицею, FALSE - ні); Довідка (назва hlp-таблиці). Решту заповнює UpdateHelp, введені раніше значення зберігаються між запусками.
2. Колонка «Порушення» показує відхилення від цих правил. Макрос відхилення не виправляє; єдина автоматична зміна - вставка рядків над таблицею, шапка якої вище рядка 3.

### Підпис

1. Підпис - три клітинки над першими колонками таблиці: назва \| «ℹ Призначення» \| «? Як заповнювати» (для hlp-таблиці - «← до tblX»). Вручну не редагують: UpdateHelp перезаписує їх.

### Нова таблиця

1. Нова таблиця з запиту: завантажити на новий аркуш → запустити UpdateHelp (додасть у каталог, зсуне шапку в рядок 3, поставить підпис) → заповнити «Призначення», за потреби створити hlp-таблицю і вписати її в «Довідку».

### Після змін

1. Після вставлення рядків чи колонок перед таблицею, перейменування таблиці чи аркуша - запустити UpdateHelp. Інакше підпис покаже «⚠ немає в каталозі».

### Службові

1. Без підпису за замовчуванням: tblTableCatalog, hlpProject, tblRunLog, test_… і tblTest….

### Технічне

1. Книга має бути збережена: формула назви в підписі використовує CELL("filename"). Потрібен Excel 365 (LET, XLOOKUP, TEXTAFTER).

## Таблиці книги

| Таблиця | Аркуш | Тип | Запит | Призначення | Довідка |
| --- | --- | --- | --- | --- | --- |
| tbl_test_All | Тести | Результат запиту | test_All | Результат тесту test_All. |  |
| tblSchemaColumns | pq_ExtractedColumns | Результат запиту | pq_ExtractedColumns | Каталог колонок файлів по призначеннях (результат pq_ExtractedColumns); джерело списків Data Validation. |  |
| tblSourceRegistry | CONFIG | Ручна |  | Реєстр джерел: один рядок = один файл (SourceID, призначення, шлях, період, політика помилок). | [hlpSourceRegistry](#hlpsourceregistry) |
| tblSchemaRegistry | SCHEMA | Ручна |  | Схема даних: колонки кожного призначення, цільові назви, типи, обов'язковість, перетворення. | [hlpSchemaRegistry](#hlpschemaregistry) |
| tblDefaultHeaders | TEMPLATES | Ручна |  | Шаблон заголовків призначення на випадок, коли шлях до файла не задано. |  |
| tblDataType | SysParameters | Ручна |  | Довідник типів даних для колонки DataType схеми (Data Validation). |  |
| tblPeriodRules | SysParameters | Ручна |  | Довідник типів періодів і максимальних номерів (PeriodType, MaxNumber). |  |
| tblNums | SysParameters | Ручна |  | Номери періодів (N) для Data Validation колонки PeriodNumber. |  |
| tblLogical | SysParameters | Ручна |  | Довідник логічних значень (TRUE/FALSE тощо) для Data Validation. |  |
| tblPeriod | SysParameters | Ручна |  | Межі звітного періоду (StartDate, EndDate) для календаря Calendar. |  |
| tblOperator | SysParameters | Ручна |  | Довідник операторів фільтрів службових рядків (Data Validation). |  |
| tblTargetType | SysParameters | Ручна |  | Довідник призначень (TargetType); список для Data Validation. |  |
| tblAnalytics | UsrParametrs | Ручна |  | Елемент бізнес-даних, довідник аналітичного рівня 1 |  |
| tblRowFilters | UsrParametrs | Ручна |  | Правила відсікання службових рядків (Всього, х тощо) до типізації. | [hlpRowFilters](#hlprowfilters) |
| pq_QualitySummary | pq_QualitySummary | Результат запиту | pq_QualitySummary | Зведення якості по джерелах (Status, QualityFlag, кількість рядків і помилок). |  |
| tblBusinessRules | RULES | Ручна |  | Елемент бізнес-даних, довідник правил опрацювання господарських операцій (на початковій стадії заповнення…) |  |
| hlpProject | HELP | Ручна |  | Загальна довідка проєкту: правила структури книги, каталогу і довідки. |  |
| hlpSourceRegistry | HELP | Ручна |  | Довідник з детальним описом та призначенням колонок tblSourceRegistry |  |
| tblErrorPolicy | HELP | Ручна |  | Довідник політик помилок (Strict, Report, Lenient) з описами. |  |
| hlpErrorThreshold | HELP | Ручна |  | Довідник-опис допустимих значень колонки ErrorThresholdта результату, який вони повертають |  |
| hlpRowFilters | HELP | Ручна |  | Довідник з детальним описом та призначенням колонок tblRowFilters |  |
| hlpSchemaRegistry | HELP | Ручна |  | Довідник з детальним описом та призначенням колонок tblSchemaRegistry |  |
| tblCustomTransform | HELP | Ручна |  | Довідник дій CustomTransform схеми (Data Validation). |  |
| tblTestExpected | Очікування | Ручна |  | Тестова книга: очікувані значення зведення якості (test_Summary). |  |
| tblTestLoad | Навантаження | Ручна |  | Тестова книга: сценарії завантаження fn_ReadPurpose (test_Load). |  |

## Довідки до таблиць

### hlpSourceRegistry

Як заповнювати таблицю **tblSourceRegistry**.

| Колонка | Тип даних / Валідація | Призначення та Логіка | Інфо |
| --- | --- | --- | --- |
| SourceID | Text / Auto | Унікальний код джерела (наприклад, SALES_MON, BANK_STMT). Зручно для прив'язки до схем. |  |
| TargetType (Призначення) | Data Validation (Список) | Категорія/призначення файлу. Прив'язує файл до відповідної бізнес-логіки та схеми. | tblTargetType |
| FullPath (ПовнийШлях) | Text (Ручно / Макрос) | Повна адреса файлу. Якщо заповнено — Path і FileName вираховуються формулами. |  |
| Path (Шлях) | Text | Каталог розміщення файлу. |  |
| FileName (Ім'я файлу) | Text | Назва файлу з розширенням (.xlsx, .csv). |  |
| TableSheetName (TableName) | Text (Опційно) | Назва листа або таблиці всередині Excel-файлу/бази даних. |  |
| PeriodType (Тип періоду) | Data Validation (Список) | Визначення періодичності звіту. | tblPeriodRules |
| PeriodNumber (Номер періоду) | Data Validation (Список) | Динамічний список, залежний від PeriodType. Визначається іменованою формулою GetPeriodNumbers. |  |
| IsRequired (Required) | Logical (TRUE / FALSE) | Якщо порожнє або TRUE, і файл відсутній — Power Query видає критичну помилку/стоп. Якщо FALSE — файл пропускається. |  |
| Enabled | Logical (TRUE / FALSE) | Дозволяє тимчасово відключати обробку окремого джерела без видалення рядка з таблиці. |  |
| ErrorPolicy | Text | Визначає рівень суворості обробки помилок. Опис - у tblErrorPolicy | tblErrorPolicy |
| ErrorThreshold | Number | Визначає поріг помилок у даних, додаткові відомості у hlpErrorThreshold | hlpErrorThreshold |

### hlpErrorThreshold

| Що вписано | Тип у клітинці | Результат |
| --- | --- | --- |
| 0,05 | відсотковий формат (Excel зберігає 0,05) | 0,05 |
| 0,05 | число | 0,05 |
| 1 | число | 100%, а не 1% |
| 5%, 0,05, 0.05, 12,5 % | текст | 5%, 5%, 5%, 12,5% |
| порожньо | — | значення з cfgSettings[ErrorThreshold], а якщо там null, поріг вимкнено |
| off, ні, no, -, false, FALSE | текст або логічне | поріг вимкнено для цього джерела, навіть якщо в cfgSettings він заданий |
| від'ємне, понад 100%, abc, TRUE | будь-який | помилка НекоректнийПоріг, рядок у pq_RegistryCheck, QualityFlag = BAD_THRESHOLD |

### hlpRowFilters

Як заповнювати таблицю **tblRowFilters**.

| Колонка | Тип даних / Валідація | Призначення та Логіка | Інфо |
| --- | --- | --- | --- |
| TargetType (Призначення) | Data Validation (Список) | Категорія/призначення файлу. Прив'язує файл до відповідної бізнес-логіки та схеми. | tblTargetType |
| SourceColumnName | Text | Оригінальна назва стовпця у вихідному файлі Excel/CSV. |  |
| Operator | Data Validation (Список) | Підтримує оператори, вказані у довідковій таблиці tblOperator | tblOperator |
| Value | Text | Текст для використання з Operator. Кілька варіантів у Value розділяються символом \|. |  |
| Use | Logical (TRUE / FALSE) | Дозволяє тимчасово відключати обробку окремого правила без видалення рядка з таблиці. |  |

### hlpSchemaRegistry

Як заповнювати таблицю **tblSchemaRegistry**.

| Колонка | Тип даних / Валідація | Призначення та Опис | Інфо |
| --- | --- | --- | --- |
| TargetType | Data Validation (Список) | Зв'язок із джерелом даних. Вказує, до якого саме призначення/шаблону належить цей рядок (наприклад, БанківськаВиписка, Продажі). | tblTargetType |
| SourceColumnName | Text | Оригінальна назва стовпця у вихідному файлі Excel/CSV. |  |
| TargetColumnName | Text | Канонічна (цільова) назва стовпця, яка буде використовуватися у підсумковій моделі після перейменування. Можна залишити порожньою, тоді буде використано значення з SourceColumnName |  |
| DataType | Data Validation (Список) | Цільовий тип даних Power Query. | tblDataType |
| IsRequired | Logical (TRUE / FALSE) | Чи є колонка обов'язковою. Якщо TRUE і вона відсутня у файлі, Schema Enforcer згенерує помилку структура звітного файлу. |  |
| TrimFlag | Logical (TRUE / FALSE) | Якщо TRUE, Power Query автоматично виконає Text.Trim +  Text.Clean для цієї колонки після завантаження. |  |
| CustomTransform | Text | Резервне поле для додаткових специфічних операцій (опційно). Кілька варіантів у Value розділяються символом ;. | tblCustomTransform |
| Use | Logical (TRUE / FALSE) | Позначка, чи використовувати стовпець. Порожне або TRUE значення означає "так", FALSE - ні (дозволяє тимчасово вимкнути колонку) |  |

## Запити Power Query

Тека — як у редакторі Power Query (cfgQueryGroups); тип і залежності — у metadata/queries.yml.

### 00 Конфігурація

- [cfgQueryGroups](<queries/cfgQueryGroups.pq>) — Config
- [cfgSettings](<queries/cfgSettings.pq>) — Config
- [cfgTypeMap](<queries/cfgTypeMap.pq>) — Config
- [cfgTransforms](<queries/cfgTransforms.pq>) — Standalone
- [cfgReportTypes](<queries/cfgReportTypes.pq>) — Config
- [cfgStatuses](<queries/cfgStatuses.pq>) — Config
- [fn_LoadConfig](<queries/fn_LoadConfig.pq>) — Function
- [fn_NormalizeSchema](<queries/fn_NormalizeSchema.pq>) — Function
- [fn_NormalizeRowFilters](<queries/fn_NormalizeRowFilters.pq>) — Function
- [src_SchemaRegistry](<queries/src_SchemaRegistry.pq>) — Transform

### 01 Утиліти

- [fn_ToBool](<queries/fn_ToBool.pq>) — Function
- [fn_ToRate](<queries/fn_ToRate.pq>) — Function
- [fn_Opt](<queries/fn_Opt.pq>) — Function
- [fn_Canon](<queries/fn_Canon.pq>) — Function
- [fn_TypeKey](<queries/fn_TypeKey.pq>) — Function

### 02 Читання джерел

- [fn_ActiveSources](<queries/fn_ActiveSources.pq>) — Function
- [fn_SourcePath](<queries/fn_SourcePath.pq>) — Function
- [fn_Options](<queries/fn_Options.pq>) — Function
- [fn_ResolveSource](<queries/fn_ResolveSource.pq>) — Function
- [fn_ParseSource](<queries/fn_ParseSource.pq>) — Function
- [fn_ParseHeaders](<queries/fn_ParseHeaders.pq>) — Function
- [fn_ReadSource](<queries/fn_ReadSource.pq>) — Function
- [fn_TemplateColumns](<queries/fn_TemplateColumns.pq>) — Function

### 03 Перевірка налаштувань

- [fn_RegistryIssues](<queries/fn_RegistryIssues.pq>) — Function
- [pq_RegistryCheck](<queries/pq_RegistryCheck.pq>) — Transform
- [pq_SchemaCheck](<queries/pq_SchemaCheck.pq>) — Transform
- [pq_RowFiltersCheck](<queries/pq_RowFiltersCheck.pq>) — Transform

### 04 Каталог і дрейф схеми

- [pq_GenerateSchemaTemplate](<queries/pq_GenerateSchemaTemplate.pq>) — Transform
- [pq_ExtractedColumns](<queries/pq_ExtractedColumns.pq>) — Transform
- [pq_SchemaIssues](<queries/pq_SchemaIssues.pq>) — Transform
- [pq_SchemaDrift](<queries/pq_SchemaDrift.pq>) — Transform

### 05 Нормалізація

- [fn_GetSchemaRules](<queries/fn_GetSchemaRules.pq>) — Function
- [fn_ApplyRowFilters](<queries/fn_ApplyRowFilters.pq>) — Function
- [fn_ApplyRulesEx](<queries/fn_ApplyRulesEx.pq>) — Function
- [fn_ApplyRules](<queries/fn_ApplyRules.pq>) — Function
- [fn_EffectiveSettings](<queries/fn_EffectiveSettings.pq>) — Function
- [fn_LoadVerdict](<queries/fn_LoadVerdict.pq>) — Function
- [fn_EnforceSchema](<queries/fn_EnforceSchema.pq>) — Function

### 06 Якість даних

- [fn_SchemaReport](<queries/fn_SchemaReport.pq>) — Function
- [fn_RowReport](<queries/fn_RowReport.pq>) — Function
- [fn_QcReports](<queries/fn_QcReports.pq>) — Function
- [fn_SummaryTable](<queries/fn_SummaryTable.pq>) — Function
- [qc_Reports](<queries/qc_Reports.pq>) — Transform
- [pq_QualitySummary](<queries/pq_QualitySummary.pq>) — Transform
- [pq_QualityColumns](<queries/pq_QualityColumns.pq>) — Transform
- [pq_QualityErrors](<queries/pq_QualityErrors.pq>) — Transform
- [fn_ColumnNullCounts](<queries/fn_ColumnNullCounts.pq>) — Function

### 07 Призначення

- [fn_ReadPurpose](<queries/fn_ReadPurpose.pq>) — Function
- [q_Registry](<queries/q_Registry.pq>) — Transform
- [q_GO](<queries/q_GO.pq>) — Transform

### 08 Модель і звіт

- [Calendar](<queries/Calendar.pq>) — Workbook
- [fxReadNamedCell](<queries/fxReadNamedCell.pq>) — Function
- [targetReportCell](<queries/targetReportCell.pq>) — Transform

### 99 Тести

- [test_All](<queries/test_All.pq>) — Transform
- [test_AuditFixes](<queries/test_AuditFixes.pq>) — Transform
- [test_Canon](<queries/test_Canon.pq>) — Workbook
- [test_Extensions](<queries/test_Extensions.pq>) — Transform
- [test_HeadersConsistency](<queries/test_HeadersConsistency.pq>) — Workbook
- [test_Load](<queries/test_Load.pq>) — Workbook
- [test_LoadVerdict](<queries/test_LoadVerdict.pq>) — Transform
- [test_NullCounts](<queries/test_NullCounts.pq>) — Transform
- [test_RegistryIssues](<queries/test_RegistryIssues.pq>) — Transform
- [test_RowFilters](<queries/test_RowFilters.pq>) — Transform
- [test_Summary](<queries/test_Summary.pq>) — Workbook
- [test_ToRate](<queries/test_ToRate.pq>) — Transform
## Макроси VBA

Код модулів — у теці vba/. Імпорт у книгу: редактор VBA, File > Import File (спершу modEngineCommon).

| Модуль | Макроси (Alt+F8) | Опис |
| --- | --- | --- |
| [modSetupTestEngine](<vba/modSetupTestEngine.bas>) | SetupTestEngine | перетворення КОПІЇ книги Engine на тестову (етап C) |
| [modCleanExportSelectedSheets](<vba/modCleanExportSelectedSheets.bas>) | CleanExportSelectedSheets |  |
| [modSelectSourceFiles](<vba/modSelectSourceFiles.bas>) | SelectSourceFiles | вибір файлів-джерел для рядків tblSourceRegistry |
| [modImportPowerQueryFromFolder](<vba/modImportPowerQueryFromFolder.bas>) | ImportPQQueriesFromFolder |  |
| [modEngineCommon](<vba/modEngineCommon.bas>) |  | спільні процедури макросів Power Query Engine (аудит D13) |
| [modExportPowerQueryCodesToRepo](<vba/modExportPowerQueryCodesToRepo.bas>) | ExportPowerQueriesToRepo | експорт коду Power Query і таблиць конфігурації для Git |
| [modTableCatalog](<vba/modTableCatalog.bas>) | UpdateHelp | каталог таблиць книги, підписи над таблицями і довідка (v4.11e) |
| [modPOWERQUERYPROFILER](<vba/modPOWERQUERYPROFILER.bas>) | ProfilePowerQueries | час оновлення кожного підключення Power Query (етап D) |
| [modRunLog](<vba/modRunLog.bas>) | RefreshWithRunLog | оновлення книги з журналом запусків (етап D, крок «журнал запусків») |

