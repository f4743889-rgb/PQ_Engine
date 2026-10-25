# PQ_ENGINE

> Файл створює макрос ExportPowerQueriesToRepo з таблиць книги (hlpProject, tblTableCatalog, hlp-таблиці, cfgQueryGroups). Вручну не редагувати: зміни вносяться в книгу.

## Правила структури книги

### Розміщення

1. Шапка кожної таблиці - не вище рядка 3. Рядок над шапкою порожній; ще вище (рядок 1) - підпис таблиці, який ставить макрос UpdateHelp.
2. Таблиці ростуть униз і вправо, тому кілька таблиць на одному аркуші ставлять лише поруч, ніколи одна під одною.
3. Таблиця з запиту Power Query (q_, pq_, test_) - одна на аркуші. Аркуш бажано називати так само, як запит. У тестовій книзі на аркуші «Тести» - лише test_All; test_Summary і test_Load - лише підключення.
4. Кожна таблиця займає щонайменше 3 колонки (місце для підпису). Між сусідніми таблицями - щонайменше одна порожня колонка.
5. Нижче і правіше таблиці нічого не розміщують: ні формул, ні приміток.

### Назви

1. Префікс задає роль таблиці: tbl - ручне введення і налаштування; q_ - дані призначення (результат запиту); pq_ - діагностика і звіти Engine; hlp - довідка; test_ - тести (лише тестова книга).
2. Назви таблиць унікальні і завжди з префіксом. Таблиця з запиту називається як запит (без tbl_). Довідка до tblX за замовчуванням - hlpX.

### Довідка

1. Аркуш HELP: tblTableCatalog, hlpProject (ця таблиця) і всі hlp-таблиці - поруч, шапки в рядку 3.
2. Призначення таблиці описують лише в tblTableCatalog[Призначення]. Правила заповнення колонок - у hlp-таблиці, вказаній у tblTableCatalog[Довідка].
3. Ця таблиця (hlpProject) - копія правил з коду модуля modTableCatalog: UpdateHelp перезаписує її щоразу. Зміни правил - через погодження, код і документ архітектури, а не вручну.

### Каталог

1. Колонки каталогу, які заповнюєте ви: Призначення; Підпис (TRUE - показувати підпис над таблицею, FALSE - ні); Довідка (назва hlp-таблиці). Решту заповнює UpdateHelp, ваші значення зберігаються між запусками.
2. Колонка «Порушення» показує відхилення від цих правил. Макрос їх не виправляє; єдина автоматична зміна - вставка рядків над таблицею, шапка якої вище рядка 3.

### Підпис

1. Підпис - три клітинки над першими колонками таблиці: назва \| «ℹ Призначення» \| «? Як заповнювати» (для hlp-таблиці - «← до tblX»). Вручну не редагують: UpdateHelp перезаписує їх.

### Нова таблиця

1. Нова таблиця з запиту: завантажити на новий аркуш → запустити UpdateHelp (додасть у каталог, зсуне шапку в рядок 3, поставить підпис) → заповнити «Призначення», за потреби створити hlp-таблицю і вписати її в «Довідку».

### Після змін

1. Після вставлення рядків чи колонок перед таблицею, перейменування таблиці чи аркуша - запустити UpdateHelp. Інакше підпис покаже «⚠ немає в каталозі».

### Службові

1. Без підпису за замовчуванням: tblTableCatalog, hlpProject, tblRunLog, test_… і tblTest….

### Git

1. Репозиторій дзеркалить робочу книгу: ExportPowerQueriesToRepo запускають лише в ній (у тестовій книзі він зупиняється). Тестова конфігурація - tests\test_config.xlsx, тестові запити - tests\queries, фікстури - tests\fixtures.

### Технічне

1. Книга має бути збережена: формула назви в підписі використовує CELL("filename"). Потрібен Excel 2021 або новіший (LET, XLOOKUP, IFS).

## Таблиці книги

| Таблиця | Аркуш | Тип | Запит | Призначення | Довідка |
| --- | --- | --- | --- | --- | --- |
| pq_ExtractedColumns | pq_ExtractedColumns | Результат запиту | pq_ExtractedColumns | Результат запиту pq_ExtractedColumns. |  |
| pq_QualitySummary | pq_QualitySummary | Результат запиту | pq_QualitySummary | Зведення якості по джерелах (Status, QualityFlag, кількість рядків і помилок). |  |
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
| tblBusinessRules | RULES | Ручна |  | Елемент бізнес-даних, довідник правил опрацювання господарських операцій (на початковій стадії заповнення…) |  |
| hlpSourceRegistry | HELP | Ручна |  | Довідник з детальним описом та призначенням колонок tblSourceRegistry |  |
| tblErrorPolicy | HELP | Ручна |  | Довідник політик помилок (Strict, Report, Lenient) з описами. |  |
| hlpErrorThreshold | HELP | Ручна |  | Довідник допустимих значень колонки ErrorThreshold і результату, який вони дають. |  |
| hlpRowFilters | HELP | Ручна |  | Довідник з детальним описом та призначенням колонок tblRowFilters |  |
| hlpSchemaRegistry | HELP | Ручна |  | Довідник з детальним описом та призначенням колонок tblSchemaRegistry |  |
| tblCustomTransform | HELP | Ручна |  | Довідник дій CustomTransform схеми (Data Validation). |  |
| hlpProject | HELP | Ручна |  | Загальна довідка проєкту: правила структури книги, каталогу і довідки. |  |

## Довідки до таблиць

### hlpSourceRegistry

Як заповнювати таблицю **tblSourceRegistry**.

| Колонка | Тип даних / Валідація | Призначення та Логіка | Інфо |
| --- | --- | --- | --- |
| SourceID | Text / Auto | Унікальний код джерела (наприклад, SALES_MON, BANK_STMT). Зручно для прив'язки до схем. |  |
| TargetType (Призначення) | Data Validation (Список) | Категорія/призначення файлу. Прив'язує файл до відповідної бізнес-логіки та схеми. | tblTargetType |
| FullPath (ПовнийШлях) | Text (Ручно / Макрос) | Повна адреса файлу. Має пріоритет над Path + FileName (fn_SourcePath). Макрос SelectSourceFiles записує FullPath і очищає Path і FileName цього рядка. |  |
| Path (Шлях) | Text | Каталог розміщення файлу. |  |
| FileName (Ім'я файлу) | Text | Назва файлу з розширенням (.xlsx, .csv). |  |
| TableSheetName (TableName) | Text (Опційно) | Назва листа або таблиці всередині Excel-файлу/бази даних. |  |
| PeriodType (Тип періоду) | Data Validation (Список) | Визначення періодичності звіту. | tblPeriodRules |
| PeriodNumber (Номер періоду) | Data Validation (Список) | Динамічний список, залежний від PeriodType. Визначається іменованою формулою GetPeriodNumbers. |  |
| IsRequired (Required) | Logical (TRUE / FALSE) | Якщо порожнє або TRUE, і файл відсутній — Power Query видає критичну помилку/стоп. Якщо FALSE — файл пропускається. |  |
| Enabled | Logical (TRUE / FALSE) | Дозволяє тимчасово відключати обробку окремого джерела без видалення рядка з таблиці. |  |
| Delimiter | Text (Опційно) | Роздільник CSV для цього джерела (наприклад, ; або ,). Порожньо — cfgSettings[CsvDelimiter]. |  |
| Encoding | Number (Опційно) | Кодова сторінка CSV: 65001 (UTF-8) або 1251. Мітка BOM у файлі має пріоритет. Порожньо — cfgSettings[CsvEncoding]. |  |
| Culture | Text (Опційно) | Культура розпізнавання чисел і дат під час типізації (наприклад, uk-UA). Порожньо — cfgSettings[Culture]. |  |
| ErrorPolicy | Text | Рівень суворості обробки помилок (Strict, Report, Lenient), опис — у tblErrorPolicy. Порожньо — cfgSettings[ErrorPolicy]. Задається лише тут або в cfgSettings. | tblErrorPolicy |
| ErrorThreshold | Number | Поріг частки рядків з помилками (лише для Report і Lenient). Порожньо — cfgSettings[ErrorThreshold]. Допустимі значення — у hlpErrorThreshold. | hlpErrorThreshold |

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
| Коли діє поріг | — | Лише для ErrorPolicy = Report і Lenient (Strict зупиняє за будь-якої помилки). Завантаження зупиняється (ПеревищенийПоріг), коли частка рядків з помилками більша за поріг; частка, що дорівнює порогу, проходить. |

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
| TrimFlag | Logical (TRUE / FALSE) | Якщо TRUE, Power Query виконає Text.Trim + Text.Clean для цієї колонки після завантаження. |  |
| CustomTransform | Text | Назви дій з tblCustomTransform; кілька дій розділяються символом ; (наприклад, Upper;NullToZero). Невідома назва дає помилку НевідомеПеретворення (pq_SchemaCheck). | tblCustomTransform |
| Use | Logical (TRUE / FALSE) | Позначка, чи використовувати стовпець. Порожне або TRUE значення означає "так", FALSE - ні (дозволяє тимчасово вимкнути колонку) |  |

## Запити Power Query

Тека — як у редакторі Power Query (cfgQueryGroups); тип і залежності — у metadata/queries.yml.

### 00 Конфігурація

- [cfgQueryGroups](<queries/cfgQueryGroups.pq>) — Config
- [cfgReportTypes](<queries/cfgReportTypes.pq>) — Config
- [cfgSettings](<queries/cfgSettings.pq>) — Config — Централізований конфіг
- [cfgStatuses](<queries/cfgStatuses.pq>) — Config
- [cfgTransforms](<queries/cfgTransforms.pq>) — Standalone — Білий список дій. Stage визначає, до чи після типізації виконується дія. Додавати нові дії можна лише тут. Назви дій порівнюються без урахування регістру
- [cfgTypeMap](<queries/cfgTypeMap.pq>) — Config — Карта типів даних: значення DataType схеми → тип M
- [fn_LoadConfig](<queries/fn_LoadConfig.pq>) — Function
- [fn_NormalizeRowFilters](<queries/fn_NormalizeRowFilters.pq>) — Function
- [fn_NormalizeSchema](<queries/fn_NormalizeSchema.pq>) — Function
- [src_SchemaRegistry](<queries/src_SchemaRegistry.pq>) — Transform

### 01 Утиліти

- [fn_Canon](<queries/fn_Canon.pq>) — Function
- [fn_Opt](<queries/fn_Opt.pq>) — Function — значення з реєстру або дефолт; порожня клітинка = дефолт
- [fn_ToBool](<queries/fn_ToBool.pq>) — Function — Гнучке приведення до logical
- [fn_ToRate](<queries/fn_ToRate.pq>) — Function — Розбір ErrorThreshold: число, відсоток або текст → частка 0..1, off → вимкнено
- [fn_TypeKey](<queries/fn_TypeKey.pq>) — Function — нормалізація значення DataType: Int64.Type → int64

### 02 Читання джерел

- [fn_ActiveSources](<queries/fn_ActiveSources.pq>) — Function — визначення «активного джерела» (Enabled, порожнє = TRUE)
- [fn_Options](<queries/fn_Options.pq>) — Function — Формування опцій обробки для парсерів
- [fn_ParseHeaders](<queries/fn_ParseHeaders.pq>) — Function — Точний та швидкий парсер заголовків
- [fn_ParseSource](<queries/fn_ParseSource.pq>) — Function — чистий парсер повних даних
- [fn_ReadSource](<queries/fn_ReadSource.pq>) — Function — Головна функція завантаження даних
- [fn_ResolveSource](<queries/fn_ResolveSource.pq>) — Function — єдине місце, де відкривається файл; помилки як дані
- [fn_SourcePath](<queries/fn_SourcePath.pq>) — Function — Шлях до файла рядка реєстру: FullPath або Path + FileName
- [fn_TemplateColumns](<queries/fn_TemplateColumns.pq>) — Function — єдине місце читання шаблонів

### 03 Перевірка налаштувань

- [fn_RegistryIssues](<queries/fn_RegistryIssues.pq>) — Function
- [pq_RegistryCheck](<queries/pq_RegistryCheck.pq>) — Transform — перевірка реєстру джерел перед обробкою
- [pq_RowFiltersCheck](<queries/pq_RowFiltersCheck.pq>) — Transform
- [pq_SchemaCheck](<queries/pq_SchemaCheck.pq>) — Transform — Перевірка самої схеми. Порожній результат означає, що схема в порядку

### 04 Каталог і дрейф схеми

- [pq_ExtractedColumns](<queries/pq_ExtractedColumns.pq>) — Transform — унікальні колонки на призначення
- [pq_GenerateSchemaTemplate](<queries/pq_GenerateSchemaTemplate.pq>) — Transform — Генератор схеми зі статусами
- [pq_SchemaDrift](<queries/pq_SchemaDrift.pq>) — Transform — схема проти реальних файлів
- [pq_SchemaIssues](<queries/pq_SchemaIssues.pq>) — Transform — таблиця для аркуша: які джерела не дали заголовків і чому

### 05 Нормалізація

- [fn_ApplyRowFilters](<queries/fn_ApplyRowFilters.pq>) — Function — відсікання службових рядків до типізації
- [fn_ApplyRules](<queries/fn_ApplyRules.pq>) — Function — конвеєр без перевірок і політик
- [fn_ApplyRulesEx](<queries/fn_ApplyRulesEx.pq>) — Function
- [fn_EffectiveSettings](<queries/fn_EffectiveSettings.pq>) — Function
- [fn_EnforceSchema](<queries/fn_EnforceSchema.pq>) — Function — перевірка обов'язкових колонок і застосування політики помилок
- [fn_GetSchemaRules](<queries/fn_GetSchemaRules.pq>) — Function — «компілює» схему призначення в запис правил
- [fn_LoadVerdict](<queries/fn_LoadVerdict.pq>) — Function

### 06 Якість даних

- [fn_ColumnNullCounts](<queries/fn_ColumnNullCounts.pq>) — Function
- [fn_QcReports](<queries/fn_QcReports.pq>) — Function
- [fn_RowReport](<queries/fn_RowReport.pq>) — Function — звіт якості для одного рядка реєстру
- [fn_SchemaReport](<queries/fn_SchemaReport.pq>) — Function — звіт про відповідність даних схемі (не залежить від ErrorPolicy)
- [fn_SummaryTable](<queries/fn_SummaryTable.pq>) — Function
- [pq_QualityColumns](<queries/pq_QualityColumns.pq>) — Transform
- [pq_QualityErrors](<queries/pq_QualityErrors.pq>) — Transform
- [pq_QualitySummary](<queries/pq_QualitySummary.pq>) — Transform

### 07 Призначення

- [fn_ReadPurpose](<queries/fn_ReadPurpose.pq>) — Function — Читання й об'єднання файлів призначення
- [q_GO](<queries/q_GO.pq>) — Transform
- [q_Registry](<queries/q_Registry.pq>) — Transform

### 08 Модель і звіт

- [Calendar](<queries/Calendar.pq>) — Workbook
- [fxReadNamedCell](<queries/fxReadNamedCell.pq>) — Function
- [targetReportCell](<queries/targetReportCell.pq>) — Transform
## Макроси VBA

Код модулів — у теці vba/. Імпорт у книгу: редактор VBA, File > Import File (спершу modEngineCommon).

| Модуль | Макроси (Alt+F8) | Опис |
| --- | --- | --- |
| [modCleanExportSelectedSheets](<vba/modCleanExportSelectedSheets.bas>) | CleanExportSelectedSheets | експорт виділених аркушів в окремі книги без запитів і зв'язків (v4.11) |
| [modEngineCommon](<vba/modEngineCommon.bas>) |  | спільні процедури макросів Power Query Engine (аудит D13; v4.11) |
| [modExportPowerQueryCodesToRepo](<vba/modExportPowerQueryCodesToRepo.bas>) | ExportPowerQueriesToRepo | експорт коду Power Query, таблиць, модулів VBA і README для Git (v4.11) |
| [modImportPowerQueryFromFolder](<vba/modImportPowerQueryFromFolder.bas>) | ImportPQQueriesFromFolder | імпорт або оновлення запитів Power Query з файлів .pq (v4.11) |
| [modPOWERQUERYPROFILER](<vba/modPOWERQUERYPROFILER.bas>) | ProfilePowerQueries | час оновлення кожного підключення Power Query (етап D; v4.12) |
| [modRunLog](<vba/modRunLog.bas>) | RefreshWithRunLog | оновлення книги з журналом запусків (етап D; v4.11) |
| [modSelectSourceFiles](<vba/modSelectSourceFiles.bas>) | SelectSourceFiles | вибір файлів-джерел для рядків tblSourceRegistry |
| [modSetupTestEngine](<vba/modSetupTestEngine.bas>) | SetupTestEngine | перетворення КОПІЇ книги Engine на тестову (етап C; v4.11) |
| [modTableCatalog](<vba/modTableCatalog.bas>) | UpdateHelp | каталог таблиць книги, підписи над таблицями і довідка (v4.11) |

