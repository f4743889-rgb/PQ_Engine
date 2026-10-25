# Power Query Engine: етапи, склад і застереження

Стан на 24.10.2026 (v4.10: етапи A–D плану стійкості, оптимізація звіту якості, виправлення за двома аудитами коду; test\_All — 11 наборів PASS). Основні потоки перевірено на реальних даних (CSV до 199 тис. рядків, Excel), непідтверджені пункти позначено в розділах 4, 7 і 8. Код надавався без запуску з боку автора відповідей, тому статуси відображають лише ваші повідомлення в розмові.

## 1. Мета і концепція

Engine це єдине модульне ядро в Power Query (M) для читання, валідації, нормалізації та аудиту даних із файлів-джерел (CSV, XLSX). Підключення баз даних (DB) у концепції передбачено, але не реалізовано.

- Менше дублювання: базові функції підключення, нормалізації та контролю якості зібрано в одному місці.
- Керування метаданими: шляхи, періоди, схеми, типи й перетворення живуть у таблицях Excel (tblSourceRegistry, tblSchemaRegistry, tblDefaultHeaders), а не в коді.
- Підтримуваність: зміни в загальній логіці завантаження не зачіпають бізнес-правил.
- Фокус на бізнес-логіці: нове джерело вимагає рядка в реєстрі та схеми, а не нового коду очищення.

## 2. Архітектура: три шари і потік даних

Відповідальність розділено між завантаженням, нормалізацією з контролем якості та об'єднанням кількох файлів одного призначення.

| Шар | Призначення | Склад |
| --- | --- | --- |
| 1. Core ETL | Читання джерел, помилки як дані, шаблони-запобіжники | cfgSettings, fn\_LoadConfig, fn\_ToBool, fn\_Canon, fn\_Opt, fn\_Options, fn\_SourcePath, fn\_ActiveSources, fn\_ResolveSource, fn\_ParseSource, fn\_ParseHeaders, fn\_ReadSource, fn\_TemplateColumns |
| 2a. Генерація і контроль схеми | Збір заголовків, довгий список колонок для Data Validation, перевірка реєстру, схеми й фільтрів, дрейф | pq\_GenerateSchemaTemplate, pq\_ExtractedColumns (на аркуші як tblSchemaColumns), pq\_SchemaIssues, pq\_RegistryCheck, fn\_NormalizeSchema, src\_SchemaRegistry, pq\_SchemaCheck, pq\_SchemaDrift, pq\_RowFiltersCheck |
| 2b. Schema Enforcer | Відсікання службових рядків, нормалізація даних за схемою, політики й поріг помилок | cfgTypeMap, fn\_TypeKey, cfgTransforms, fn\_NormalizeRowFilters, fn\_ApplyRowFilters, fn\_GetSchemaRules, fn\_ApplyRulesEx, fn\_ApplyRules, fn\_ToRate, fn\_EffectiveSettings, fn\_LoadVerdict, fn\_EnforceSchema |
| 2c. QC Suite | Аудит якості даних | cfgReportTypes, cfgStatuses, fn\_SchemaReport, fn\_RowReport, fn\_QcReports, fn\_SummaryTable, qc\_Reports, pq\_QualitySummary, pq\_QualityColumns, pq\_QualityErrors |
| Тести | Регресійні й модульні тести (окрема тестова книга) | test\_Summary, test\_Load, test\_All, test\_ToRate, test\_Canon, test\_LoadVerdict, test\_RegistryIssues, test\_Extensions, test\_NullCounts, test\_AuditFixes, test\_HeadersConsistency, test\_RowFilters |
| 3. Диспетчер | Об'єднання файлів призначення | fn\_ReadPurpose |

Потік обробки:

1. tblSourceRegistry → fn\_ActiveSources → fn\_ResolveSource (доступ до файла, помилки як дані) → fn\_ParseHeaders / fn\_ParseSource.
2. Генерація схеми: pq\_GenerateSchemaTemplate → pq\_ExtractedColumns. Результат pq\_ExtractedColumns завантажується на аркуш як tblSchemaColumns (довгий список унікальних колонок блоками по призначенню) і є джерелом списків Data Validation. Джерела, що не дали заголовків, видно в pq\_SchemaIssues.
3. Ви заповнюєте tblSchemaRegistry: зразкову схему даних проєкту. Список SourceColumnName залежить від TargetType рядка (іменована формула GetDependentList).
4. Контроль: pq\_RegistryCheck (реєстр, зокрема значення ErrorThreshold), pq\_SchemaCheck (схема), pq\_RowFiltersCheck (правила tblRowFilters), pq\_SchemaDrift (файли проти схеми).
5. Нормалізація одного файла: fn\_ReadSource → fn\_GetSchemaRules (схема + правила tblRowFilters призначення) → fn\_EnforceSchema. Усередині fn\_ApplyRulesEx: відсікання службових рядків (fn\_ApplyRowFilters) → канонічна форма → видалення порожніх рядків → до-типові дії → перейменування → типи → після-типові дії. Потім fn\_EnforceSchema застосовує політику (Strict, Report, Lenient) і поріг помилок.
6. Об'єднання: fn\_ReadPurpose читає всі активні файли призначення, нормалізує кожен окремо і додає SourceID, PeriodType, PeriodNumber. Перед читанням зупиняється, якщо fn\_RegistryIssues знайшла для призначення блокуючу проблему реєстру. Виклик: `fn_ReadPurpose(fn_LoadConfig(), "Registry")` (реєстр, схема і фільтри службових рядків передаються одним записом).
7. Аудит: qc\_Reports → pq\_QualitySummary (зокрема RowsDropped, ErrorRate, QualityFlag), pq\_QualityColumns, pq\_QualityErrors. Звіти показують дані з помилками незалежно від політики, а QualityFlag прогнозує, чи зупиниться завантаження.

Таблиці Excel і їхні ролі:

| Таблиця | Роль |
| --- | --- |
| tblSourceRegistry | Один рядок = один файл. Колонки: SourceID, TargetType, FullPath, Path, FileName, TableSheetName, PeriodType, PeriodNumber, IsRequired, Enabled. Необов'язково: Delimiter, Encoding, ErrorPolicy, ErrorThreshold (перевизначають значення з cfgSettings для джерела) |
| tblRowFilters | Правила відсікання службових рядків до типізації: TargetType, SourceColumnName, Operator, Value, Use. Оператори з tblOperator: Equals, StartsWith, Contains, IsBlank, NotNumeric |
| tblDefaultHeaders | Шаблон-запобіжник, один рядок = одна очікувана колонка (TargetType, SourceColumnName, необов'язково ColumnOrder). Діє, коли шляху до файла немає. Колонка DataType не використовується |
| tblSchemaRegistry | Зразкова схема проєкту: TargetType, SourceColumnName, TargetColumnName, DataType, IsRequired, TrimFlag, CustomTransform, Use. Створюється на основі tblSchemaColumns зі списками Data Validation |
| tblSchemaColumns | Результат запиту pq\_ExtractedColumns на окремому аркуші: рядок = колонка призначення (TargetType, SourceColumnName, FirstIdx, FilesWithColumn, InTemplate, FilesTotal). Блоки по призначенню суцільні. Джерело списків Data Validation |

## 3. Дорожня карта

Статуси зведено за результатами тестування на реальних даних; деталі в розділі 7.

| Етап | Зміст | Статус |
| --- | --- | --- |
| 1. Метадані й конфігурація | tblSourceRegistry, tblSchemaRegistry, tblDefaultHeaders, cfgSettings | Завершено |
| 2. Core Engine | fn\_ResolveSource, fn\_ParseSource, fn\_ParseHeaders, fn\_ReadSource, fn\_SourcePath, fn\_ActiveSources | Завершено; Excel перевірено на таблиці, аркуші та збігу імен |
| 3. Генерація і контроль схеми | pq\_GenerateSchemaTemplate, pq\_ExtractedColumns (tblSchemaColumns), pq\_SchemaIssues, pq\_RegistryCheck, pq\_SchemaCheck, pq\_SchemaDrift | Завершено; списки Data Validation працюють |
| 4. Schema Enforcer | fn\_GetSchemaRules, fn\_ApplyRules, fn\_EnforceSchema, cfgTypeMap, cfgTransforms | Завершено; політика Lenient не перевірена |
| 5. QC Suite | fn\_SchemaReport, fn\_RowReport, qc\_Reports, pq\_QualitySummary, pq\_QualityColumns, pq\_QualityErrors | Завершено |
| 6. Диспетчеризація | fn\_ReadPurpose: перевірки SourceID, файлів і періодів, метадані після Enforcer | Завершено |
| 6A. Стійкість, етап A | Відсікання службових рядків (tblRowFilters), поріг помилок ErrorThreshold, індикатор QualityFlag | Завершено 15.10.2026 |
| 6B. Стійкість, етап B | Макрос вибору файлів SelectSourceFiles; крайні випадки джерел (дублікати заголовків, #N/A, відкриті файли); правило регістру; BasePath відкладено | Завершено 16.10.2026 |
| 6C. Стійкість, етап C | Регресійні тести: тестова книга, 21 тестовий файл, 24 сценарії зведення, 25 сценаріїв завантаження; аудит D4–D7 | Завершено 16.10.2026 |
| 6D. Стійкість, етап D | Спільний модуль VBA (modEngineCommon), журнал запусків (RefreshWithRunLog), експорт коду й таблиць конфігурації для Git, виправлений профайлер, розширення файлів у cfgSettings; усунуто повторне читання файлів (Table.Buffer); аудит D8 (розширення), D13 | Завершено 17.10.2026 |
| 7. Бізнес-обробка | Специфічна логіка призначень над нормалізованими даними | Не розпочато (за межами Engine) |
| 8. DB-джерела | Підключення баз даних | Не реалізовано |

Вилучено або замінено: fn\_TemplateTable, src\_SourceRegistry і pq\_GenerateSchemaTemplate\_Wide видалено (список колонок для Data Validation тепер довгий, tblSchemaColumns), fxListsToTable замінено на Table.FromColumns, формат DefaultColumnsList зі списком через крапку з комою замінено реляційною tblDefaultHeaders. fn\_ResolveFolderFile у поточній моделі не використовується (режим папки не потрібен), її рекомендовано видалити.

## 4. Реєстр запитів і функцій

Статуси зведено за вашими повідомленнями в розмові: «Підтв.» означає, що ви підтвердили результат тестом, «Очікує» означає, що перевірки не було.

| Назва | Тип | Шар | Призначення | Залежить від | Статус |
| --- | --- | --- | --- | --- | --- |
| cfgSettings | запис | 1 | Параметри: CsvDelimiter, CsvEncoding, Culture, TrueValues, ErrorPolicy, ErrorThreshold, MetaColumns; імена іменованих клітинок NameTargetType, NameMVO, NameReportDate; імена таблиць книги Tables; розширення файлів Extensions (Csv, Excel; з v4.8). Читають і запити M, і макроси (engLoadSettings) | — | Підтв. |
| fn\_ToBool | функція | 1 | Нормалізація логічних значень з дефолтом | cfgSettings | Підтв. |
| fn\_Opt | функція | 1 | Значення з рядка реєстру або дефолт | — | Підтв. |
| fn\_Options | функція | 1 | Delimiter, Encoding, TableSheetName рядка реєстру | fn\_Opt, cfgSettings | Підтв. |
| fn\_SourcePath | функція | 1 | Повний шлях за рядком реєстру (FullPath має пріоритет) | — | Підтв. |
| fn\_Canon | функція | 1 | Правило регістру: значення довідника у написанні з таблиці-довідника (tblTargetType, tblPeriodRules) | — | Підтв. (test\_Canon) |
| fn\_ActiveSources | функція | 1 | Єдине визначення активного джерела (Enabled, порожнє = TRUE); TargetType і PeriodType у написанні довідників | fn\_ToBool, fn\_Canon | Підтв. |
| fn\_ResolveSource | функція | 1 | Доступ до файла (перевірка першого байта), помилки як дані | fn\_SourcePath | Підтв. |
| fn\_ParseSource | функція | 1 | Повне читання CSV чи Excel з бінарного вмісту, Excel за Item; формат за cfgSettings[Extensions] | cfgSettings | Підтв. (CSV, Excel, test\_Extensions) |
| fn\_ParseHeaders | функція | 1 | Швидке читання заголовків; CSV за cfgSettings[Extensions][Csv] | fn\_ParseSource, cfgSettings | Підтв. |
| fn\_ReadSource | функція | 1 | Читання джерела з урахуванням IsRequired | fn\_ResolveSource, fn\_Options, fn\_ParseSource, fn\_ToBool | Підтв. |
| fn\_TemplateColumns | функція | 1 | Колонки шаблону з tblDefaultHeaders | tblDefaultHeaders | Підтв. |
| pq\_GenerateSchemaTemplate | запит | 2a | Заголовки файлів (довгий формат) і статуси | fn\_ResolveSource, fn\_ParseHeaders, fn\_Options, fn\_TemplateColumns, fn\_ActiveSources | Підтв. |
| pq\_ExtractedColumns | запит | 2a | Унікальні колонки призначень блоками по призначенню, з лічильниками файлів; на аркуші як tblSchemaColumns (джерело списків Data Validation); вхід буферизовано (v4.8) | pq\_GenerateSchemaTemplate | Підтв. |
| pq\_SchemaIssues | запит | 2a | Джерела, що не дали заголовків, і причини | pq\_GenerateSchemaTemplate | Підтв. |
| fn\_RegistryIssues | функція | 2a | Усі перевірки реєстру: дублі SourceID (у всьому реєстрі), файлу і періоду (у призначенні), порожній SourceID, недопустимий період, некоректний ErrorThreshold, TargetType поза довідником; колонка Blocking | fn\_ActiveSources, fn\_SourcePath, fn\_ToRate, fn\_ToBool | Підтв. (test\_RegistryIssues) |
| pq\_RegistryCheck | запит | 2a | Звіт fn\_RegistryIssues: Issue, Key, Count, TargetTypes, Blocking | fn\_LoadConfig, fn\_RegistryIssues | Підтв. |
| pq\_RowFiltersCheck | запит | 2a | Перевірка tblRowFilters: порожні поля, невідомий Operator, порожнє Value, TargetType поза схемою | fn\_NormalizeSchema | Підтв. (порожній результат) |
| fn\_NormalizeSchema | функція | 2a | Нормалізація tblSchemaRegistry | fn\_ToBool | Підтв. |
| src\_SchemaRegistry | запит (лише підключення) | 2a | Нормалізована схема для запитів, що збирають результати | fn\_NormalizeSchema | Підтв. |
| pq\_SchemaCheck | запит | 2a | Перевірка коректності самої схеми | src\_SchemaRegistry, cfgTypeMap, cfgTransforms, fn\_TypeKey | Підтв. |
| pq\_SchemaDrift | запит | 2a | Порівняння файлів зі схемою; підказка «регістр відрізняється» для назв колонок; входи буферизовано (v4.8) | pq\_ExtractedColumns, src\_SchemaRegistry | Підтв. |
| cfgTypeMap | запис | 2b | Словник типів | — | Підтв. |
| fn\_TypeKey | функція | 2b | Нормалізація назви типу (Int64.Type → int64) | — | Підтв. |
| cfgTransforms | запис | 2b | Білий список дій CustomTransform зі стадією Pre або Post | — | Підтв. |
| fn\_NormalizeRowFilters | функція | 2b | Нормалізація tblRowFilters (обрізання, Use: порожнє = TRUE) | fn\_ToBool | Підтв. |
| fn\_ApplyRowFilters | функція | 2b | Відсікання службових рядків за правилами; повертає Data, Dropped, RulesApplied, RulesSkipped | — | Підтв. (test\_RowFilters, реальний файл Registry) |
| fn\_GetSchemaRules | функція | 2b | Компіляція правил схеми призначення; необов'язковий rowFilters потрапляє в Rules[RowFilters] | cfgTypeMap, cfgTransforms, fn\_TypeKey | Підтв. |
| fn\_ApplyRulesEx | функція | 2b | Конвеєр нормалізації без політик, з відсіканням службових рядків першим кроком; повертає Data, Dropped, RulesSkipped | fn\_ApplyRowFilters, cfgSettings | Підтв. |
| fn\_ApplyRules | функція | 2b | Обгортка fn\_ApplyRulesEx, що повертає лише дані | fn\_ApplyRulesEx | Підтв. |
| fn\_ToRate | функція | 2b | Розбір порогу ErrorThreshold у частку 0..1 або null | cfgSettings | Підтв. (test\_ToRate) |
| fn\_EffectiveSettings | функція | 2b | Ефективні Policy, Threshold, Culture джерела: реєстр → options → cfgSettings | fn\_Opt, fn\_ToRate, cfgSettings | Підтв. (тести) |
| fn\_LoadVerdict | функція | 2b | Єдине правило «чи зупиниться завантаження» (Stop, Reason, Flag) для fn\_EnforceSchema і QualityFlag | — | Підтв. (test\_LoadVerdict) |
| fn\_EnforceSchema | функція | 2b | Перевірка обов'язкових колонок, політики Strict, Report, Lenient і порогу; третій параметр — запис налаштувань fn\_EffectiveSettings | fn\_ApplyRules, fn\_EffectiveSettings, fn\_LoadVerdict | Підтв. (Strict, Report, Lenient, поріг) |
| fn\_SchemaReport | функція | 2c | Звіт відповідності схемі, зокрема RowsDropped, FilterRulesSkipped і DuplicateHeaders | fn\_ApplyRulesEx | Підтв. |
| cfgReportTypes | запис | 2c | Типи таблиць звітів: Summary, Columns, Errors | — | Підтв. |
| cfgStatuses | запис | 2c | Статуси зведення: порядок, «блокуючий», опис | — | Підтв. |
| fn\_RowReport | функція | 2c | Звіт одного джерела, що не падає; додає ErrorRate, ErrorThreshold, QualityFlag | fn\_ResolveSource, fn\_ParseSource, fn\_SchemaReport, fn\_EffectiveSettings, fn\_LoadVerdict, cfgStatuses | Підтв. (усі статуси, тести) |
| fn\_QcReports | функція | 2c | Звіти якості активних джерел: fn\_QcReports(config, optional sourceIds); фільтр SourceID після fn\_ActiveSources (v4.9) | fn\_ActiveSources, fn\_GetSchemaRules, fn\_RowReport | Підтв. (test\_Summary) |
| fn\_SummaryTable | функція | 2c | Таблиця зведення зі списку звітів: типи cfgReportTypes[Summary], порядок cfgStatuses (v4.9) | cfgReportTypes, cfgStatuses | Підтв. |
| qc\_Reports | запит (лише підключення) | 2c | Колектор звітів: fn\_QcReports(fn\_LoadConfig()) | fn\_LoadConfig, fn\_QcReports | Підтв. |
| pq\_QualitySummary | запит | 2c | Підсумок по джерелах зі статусами, RowsDropped, ErrorRate, QualityFlag: fn\_SummaryTable(qc\_Reports) | qc\_Reports, fn\_SummaryTable | Підтв. |
| pq\_QualityColumns | запит | 2c | Звіт по колонках схеми | qc\_Reports | Підтв. |
| pq\_QualityErrors | запит | 2c | Клітинки з помилками: рядок, причина, значення | qc\_Reports | Підтв. |
| fn\_LoadConfig | функція | 1 | Конфігурація з книги одним записом: Registry, Schema, RowFilters, RowFiltersRaw; імена таблиць з cfgSettings[Tables] | cfgSettings, fn\_NormalizeSchema, fn\_NormalizeRowFilters | Підтв. |
| fn\_ReadPurpose | функція | 3 | Читання й об'єднання файлів призначення: fn\_ReadPurpose(config, targetType, options); блокуючі проблеми реєстру з fn\_RegistryIssues | fn\_ActiveSources, fn\_RegistryIssues, fn\_ReadSource, fn\_GetSchemaRules, fn\_EnforceSchema, fn\_EffectiveSettings | Підтв. |
| q\_Registry | запит | 3 | Запит призначення Registry (зразок для інших q\_…): fn\_ReadPurpose(fn\_LoadConfig(), "Registry") | fn\_LoadConfig, fn\_ReadPurpose | Підтв. |
| test\_RowFilters, test\_RowFiltersReal, test\_ToRate, test\_Canon, test\_LoadVerdict, test\_RegistryIssues, test\_Extensions | запити | тести | Модульні тести: фільтри, розбір порогу, правило регістру, правило зупинки, перевірки реєстру, розширення файлів (вміст CSV задано текстом, файлів не відкриває) | fn\_ApplyRowFilters, fn\_ToRate, fn\_Canon, fn\_LoadVerdict, fn\_RegistryIssues, fn\_ParseSource, fn\_ParseHeaders | Підтв. |
| test\_Summary, test\_Load, test\_All | запити | тести | Регресійні тести (етап C): зведення проти tblTestExpected (з v4.9 лише для SourceID з tblTestExpected, сам читає книгу; src\_TestExpected видалено), завантаження проти tblTestLoad, підсумок | fn\_LoadConfig, fn\_QcReports, fn\_SummaryTable, fn\_ReadPurpose | Підтв. |
| cfgQueryGroups | запис | 1 | Структура тек Power Query (розділ 5.1): Groups, Prefixes; читає макрос експорту | — | Очікує |

Поза Engine у книзі: Calendar (календар для Power Pivot), fxReadNamedCell і targetReportCell (читання іменованих клітинок). Зв'язки між таблицями налаштовано в моделі даних Power Pivot.

## 5. Порядок створення, завантаження та правило запитів

Рекомендований порядок створення (кожен наступний залежить від попередніх): cfgSettings → fn\_ToBool, fn\_ToRate, fn\_Opt, fn\_Options, fn\_SourcePath, fn\_ActiveSources → fn\_ResolveSource → fn\_ParseSource → fn\_ParseHeaders → fn\_ReadSource → fn\_TemplateColumns → pq\_GenerateSchemaTemplate → pq\_ExtractedColumns → pq\_SchemaIssues, pq\_RegistryCheck → cfgTypeMap, fn\_TypeKey, cfgTransforms, fn\_NormalizeSchema, fn\_NormalizeRowFilters → src\_SchemaRegistry → pq\_SchemaCheck, pq\_SchemaDrift, pq\_RowFiltersCheck → fn\_GetSchemaRules → fn\_ApplyRowFilters → fn\_ApplyRulesEx → fn\_ApplyRules → fn\_EnforceSchema → fn\_SchemaReport → fn\_RowReport → qc\_Reports → pq\_QualitySummary, pq\_QualityColumns, pq\_QualityErrors → fn\_ReadPurpose → q\_… (запити призначень). Під час імпорту макросом ImportPQQueriesFromFolder порядок неважливий: посилання розв'язуються під час оновлення.

Завантаження (рекомендація): лише підключення для cfg-запитів, fn-функцій, src\_SchemaRegistry і qc\_Reports. На аркуш варто завантажувати те, що ви переглядаєте або використовуєте: pq\_ExtractedColumns (як таблицю tblSchemaColumns, джерело списків Data Validation), pq\_SchemaIssues, pq\_RegistryCheck, pq\_SchemaCheck, pq\_RowFiltersCheck, pq\_SchemaDrift, pq\_QualitySummary та запити призначень (q\_…, що викликають fn\_ReadPurpose). pq\_QualityColumns і pq\_QualityErrors перечитують файли через qc\_Reports, тому їх краще виключити з «Оновити все» (властивості підключення) і запускати за потреби.

Кожна таблиця на аркуші, яку має оновлювати «Оновити все», повинна мати в підключенні позначку «Оновлювати це підключення під час оновлення всіх» (Дані → Запити й підключення → Підключення → Властивості). Без неї таблиця мовчки лишається застарілою, хоча запит у редакторі вже дає новий результат (17.10.2026: так сталося з pq\_QualitySummary на аркуші «Якість»; RefreshWithRunLog тепер оновлює цю таблицю явно).

Правило трьох видів запитів, яке запобігає помилці Formula.Firewall:

| Вид запиту | Приклади | Робота з книгою |
| --- | --- | --- |
| Відкривають файли | pq\_GenerateSchemaTemplate, qc\_Reports, запити виклику fn\_ReadPurpose (q\_…), test\_RowFiltersReal | Читають таблиці книги напряму, викликають лише функції (fn\_, cfg), на запити src\_ і pq\_ не посилаються. Схему беруть через fn\_NormalizeSchema, фільтри через fn\_NormalizeRowFilters |
| Збирають результати інших запитів | pq\_ExtractedColumns, pq\_SchemaIssues, pq\_SchemaDrift, pq\_Quality\* | Самі книгу й файли не читають, схему беруть із src\_SchemaRegistry |
| Лише перевіряють таблиці книги | pq\_RegistryCheck, pq\_SchemaCheck, pq\_RowFiltersCheck | Файлів не відкривають, тому допустимі обидва способи |

### 5.1. Теки Power Query (з v4.9)

Запити в редакторі Power Query згруповано за шарами в порядку потоку даних. Єдине джерело структури — запит cfgQueryGroups (Groups: тека → імена запитів; Prefixes: q\_ → «07 Призначення», test\_ → «99 Тести»). Макрос експорту записує теку кожного запиту в поле group файла queries.yml і показує запити, яких немає в cfgQueryGroups. Теки розставляються вручну: VBA не має доступу до тек; ImportPQQueriesFromFolder і SetupTestEngine під час оновлення наявного запиту теку зберігають, новий запит з'являється поза теками. Теки не впливають на обчислення і Formula.Firewall.

| Тека | Призначення | Запити |
| --- | --- | --- |
| 00 Конфігурація | Параметри Engine, довідники в M, читання таблиць-налаштувань | cfgSettings, cfgQueryGroups, cfgTypeMap, cfgTransforms, cfgReportTypes, cfgStatuses, fn\_LoadConfig, fn\_NormalizeSchema, fn\_NormalizeRowFilters, src\_SchemaRegistry |
| 01 Утиліти | Дрібні функції без доменної логіки | fn\_ToBool, fn\_ToRate, fn\_Opt, fn\_Canon, fn\_TypeKey |
| 02 Читання джерел | Шлях, доступ до файла, розбір CSV/Excel, помилки як дані | fn\_ActiveSources, fn\_SourcePath, fn\_Options, fn\_ResolveSource, fn\_ParseSource, fn\_ParseHeaders, fn\_ReadSource, fn\_TemplateColumns |
| 03 Перевірка налаштувань | Контроль таблиць книги без читання файлів | fn\_RegistryIssues, pq\_RegistryCheck, pq\_SchemaCheck, pq\_RowFiltersCheck |
| 04 Каталог і дрейф схеми | Заголовки файлів, каталог колонок для Data Validation, файли проти схеми | pq\_GenerateSchemaTemplate, pq\_ExtractedColumns, pq\_SchemaIssues, pq\_SchemaDrift |
| 05 Нормалізація | Schema Enforcer: правила, фільтри службових рядків, типи, політики, поріг | fn\_GetSchemaRules, fn\_ApplyRowFilters, fn\_ApplyRulesEx, fn\_ApplyRules, fn\_EffectiveSettings, fn\_LoadVerdict, fn\_EnforceSchema |
| 06 Якість даних | Звіти якості, що не падають, і підсумки | fn\_SchemaReport, fn\_RowReport, fn\_QcReports, fn\_SummaryTable, qc\_Reports, pq\_QualitySummary, pq\_QualityColumns, pq\_QualityErrors |
| 07 Призначення | Межа Engine: дані для бізнес-обробки | fn\_ReadPurpose, q\_Registry, усі нові q\_… |
| 08 Модель і звіт | Поза Engine | Calendar, fxReadNamedCell, targetReportCell |
| 99 Тести | Лише тестова книга | test\_\* |

## 6. Застереження щодо використання

### 6.1. Formula.Firewall

- Дотримуйтеся правила трьох видів запитів (розділ 5). Запити, що відкривають файли, не повинні посилатися на запити src\_ чи pq\_: так було з qc\_Reports до переходу на fn\_NormalizeSchema.
- src\_SchemaRegistry має лишатися запитом «Лише підключення».
- Запити, що відкривають файли, отримують конфігурацію через fn\_LoadConfig(): функція вбудовується в запит-викликач, тож читання книги лишається в ньому (як у fn\_TemplateColumns). Перевірено на q\_Registry, qc\_Reports, test\_Load, pq\_GenerateSchemaTemplate.
- Виняток: якщо таблицю, отриману через fn\_LoadConfig, передати всередину власної функції запиту, що відкриває файли, Power Query дає Formula.Firewall (test\_RowFiltersReal, 16.10.2026). Тоді таблиці читаються в самому запиті: `Excel.CurrentWorkbook(){[Name = cfgSettings[Tables][Registry]]}[Content]`; імена таблиць лишаються в cfgSettings[Tables].
- Якщо помилка з'явиться знову: Table.Buffer або List.Buffer значень з книги перед доступом до файлів, або налаштування конфіденційності для довірених локальних джерел.

### 6.2. Реєстр джерел

- SourceID має бути унікальним у всьому реєстрі (з v4.7; раніше завантаження перевіряло лише в межах призначення): fn\_ReadPurpose зупиняється з помилкою ДублюванняSourceID для кожного призначення, де трапляється дубль. Порожній SourceID завантаження не зупиняє, але видно в pq\_RegistryCheck. Усі перевірки реєстру реалізує fn\_RegistryIssues; колонка Blocking у pq\_RegistryCheck показує, які проблеми зупинять завантаження.
- Імена таблиць книги (tblSourceRegistry, tblSchemaRegistry, tblRowFilters, tblDefaultHeaders, tblTargetType, tblPeriodRules) задано лише в cfgSettings[Tables]; перейменовуючи таблицю, змініть і там.
- Пара PeriodType + PeriodNumber має бути унікальною серед активних джерел призначення (ДублюванняПеріоду, параметр AllowDuplicatePeriods).
- Той самий файл і лист не можуть бути в кількох активних джерелах одного призначення (ДублюванняФайлу, параметр AllowDuplicateFiles); різні листи одного файла дозволені. Перевірка не залежить від IsRequired. pq\_RegistryCheck показує дубль лише для активних рядків.
- Порожні Enabled та IsRequired означають TRUE. Значення з TrueValues (TRUE, Так, 1, +) розпізнаються.
- ErrorPolicy і ErrorThreshold у рядку реєстру мають пріоритет над options виклику fn\_ReadPurpose, а ті над cfgSettings. Зведення якості (qc\_Reports) бачить лише реєстр і cfgSettings, тому політику для перевірок задавайте в реєстрі. Правила заповнення ErrorThreshold див. у 6.7.
- Якщо заповнено FullPath, він має пріоритет над Path і FileName.
- Вимкнене джерело (Enabled = FALSE) однаково зникає з даних, звітів і перевірок реєстру (fn\_ActiveSources).

### 6.3. Читання файлів

- File.Contents лінивий, тому звичайний try не виявляє відсутній файл. fn\_ResolveSource читає перший байт усередині try: недоступний обов'язковий файл дає помилку \[ФайлНедоступний\] \[SourceID\], необов'язковий дає порожній результат зі статусом NODATA. Порожні (0 байт) файли не передбачаються.
- Роздільник і кодування беруться з cfgSettings (перевірте CsvEncoding: 65001 або 1251) і можуть перевизначатися колонками Delimiter та Encoding.
- Підтримувані розширення: csv, txt, xlsx, xlsm, xls.
- Excel: TableSheetName порівнюється з колонкою Item результату Excel.Workbook (а не з Name, який Excel робить унікальним і може доповнити цифрою) без урахування регістру. Спочатку шукається таблиця, потім аркуш. Якщо ім'я не задано, вибір автоматичний лише за однозначності, інакше помилка НеоднозначнийВибір. Для аркуша заголовок має бути в першому рядку.
- CSV: заголовки читаються з першого рядка (лапки, BOM обробляються).
- Помилкові клітинки в самому джерелі Excel (наприклад, #N/A) рахуються як рядки з помилками і підпадають під політику та поріг: Report лишає їх (на аркуші клітинка порожня), Strict зупиняє завантаження з повідомленням ПомилкиТипізації, хоча помилка походить із джерела. Формули, що залежать від помилкової клітинки, теж дають помилку (тест: =NA() у «Кількість» дала ErrorColumns = 2 разом із сумою). Що саме сталося, видно в Reason і Message у pq\_QualityErrors.
- Дублікати заголовків: Power Query бере першу колонку з назвою, а повтори перейменовує на «Назва\_1», «Назва\_2». Якщо повторюється колонка схеми, Engine мовчки бере першу копію; тому pq\_QualitySummary показує DuplicateHeaders і Status = WARN, pq\_SchemaDrift — колонку з суфіксом зі статусом «Немає у схемі», а SelectSourceFiles попереджає під час вибору CSV. Колонку, що справді називається «X\_1» поряд із «X», теж буде позначено.
- Файл, відкритий в Excel або Блокноті, читанню не заважає (перевірено). Справжнє блокування іншою програмою дає ФайлНедоступний.

### 6.4. Шаблони і генерація схеми

- Шаблон береться, коли шляху немає (статус OK) або шлях є, але файл недоступний (WARN). Без файла й шаблону статус ERROR, а SourceColumnName порожнє (текст помилки в назву колонки не потрапляє). Щоб лишити шаблон лише для порожнього шляху, у pq\_GenerateSchemaTemplate змініть умову на `TplCols <> null and NoPath`.
- Колонки шаблону не є фактами про файли: pq\_SchemaDrift порівнює схему лише з реальними файлами.
- Джерела, що не дали заголовків, видно в pq\_SchemaIssues (SourceID, статус, повідомлення). Без неї вони тихо випадають зі списку колонок.
- pq\_QualityColumns і pq\_QualityErrors додають до об'єднання порожню таблицю потрібної форми, тому для порожнього реєстру Excel не показує загальне попередження про запит без стовпців. pq\_ExtractedColumns зберігає колонки навіть для порожнього результату.
- Data Validation для SourceColumnName: pq\_ExtractedColumns завантажено на аркуш як tblSchemaColumns, блоки по призначенню суцільні (призначення за першою появою, колонки за FirstIdx). Джерелом списку в Перевірці даних служить іменована формула GetDependentList: `=LET(target;tblSchemaRegistry[@TargetType];src;tblSchemaColumns[SourceColumnName];types;tblSchemaColumns[TargetType];first;MATCH(target;types;0);cnt;COUNTIF(types;target);OFFSET(INDEX(src;first);0;0;cnt;1))` (роздільник «;» для української локалі).
- Excel не дозволяє в джерелі списку оператори посилань (INDEX():INDEX()), константи масивів і LAMBDA, тому використано OFFSET усередині іменованої формули з відносним посиланням на рядок таблиці. Для порожнього TargetType або призначення, якого немає в tblSchemaColumns, кнопка списку неактивна, наявні значення зберігаються.
- Якщо назви призначень можуть містити символи \* ? \~, замініть MATCH і COUNTIF на `XMATCH(target;types;0)` і `SUM(--(types=target))`.
- tblSchemaRegistry наповнюйте макросом, який дописує відсутні пари TargetType + SourceColumnName без перезапису заповнених полів, а не формулами.

### 6.5. Контроль схеми

- Збіги назв колонок точні (регістр і пробіли), як у Table.Combine: це заголовки файлів, а не довідник, тому правило регістру (6.9) на них не поширюється. Зайвий пробіл у схемі дасть пару «Немає у схемі» та «Відсутня»; якщо назви відрізняються лише регістром, pq\_SchemaDrift додає «: регістр відрізняється».
- Назви в CustomTransform мають збігатися з cfgTransforms без урахування регістру (upper = Upper); кілька дій розділяються крапкою з комою (Upper;NullToZero). Помилки введення показує pq\_SchemaCheck.
- pq\_SchemaDrift перевіряє лише призначення, де прочитано хоча б один файл; решта позначається «Не перевірено (немає файлів)». Колонки з Use = Ні мають статус «Ігнорується».

### 6.6. Schema Enforcer і політики помилок

- Порядок етапів фіксований: відсікання службових рядків за tblRowFilters (на колонках джерела, до вибору колонок схеми) → канонічна форма (усі колонки схеми в порядку схеми, відсутні необов'язкові колонки = null) → видалення порожніх рядків → до-типові дії (TrimFlag = Trim + Clean, потім текстові CustomTransform) → перейменування → типи (культура з cfgSettings) → після-типові дії. Кроки PostSteps мають тип третім елементом.
- Колонки з Use = Ні виключено з усіх правил. Обов'язковість колонок перевіряється лише для реальних даних; коли джерела немає, повертається порожня таблиця тієї ж форми й з тими ж типами.
- Політики: Strict зупиняє запит із помилкою \[ПомилкиТипізації\] та числом рядків і колонок (повідомлення збирається з буфера лише рядків з помилками), Report лишає помилки в клітинках, Lenient замінює їх на null (Table.ReplaceErrorValues; до v4.6 ця гілка не працювала, ваду знайшов тест L06). Політика задається в cfgSettings, в options виклику fn\_ReadPurpose або в колонці ErrorPolicy реєстру.
- Поріг ErrorThreshold діє лише для Report і Lenient: якщо частка рядків з хоча б однією помилкою типізації перевищує поріг, запит зупиняється помилкою \[ПеревищенийПоріг\] з кількістю рядків, часткою і порогом. У Strict поріг не перевіряється, бо зупиняє будь-яка помилка. Підрахунок частки робить ще один прохід по даних, але лише коли поріг заданий.
- У режимі Report на аркуші клітинки з помилками виглядають порожніми, тому відрізнити їх від null можна лише за звітами (pq\_QualityErrors). Звіти не залежать від політики.
- RowNo у звіті нумерує рядки даних після видалення порожніх рядків (з 1, без шапки), тож може відрізнятися від рядка у файлі.
- Службові рядки всередині таблиць (наприклад, «Всього» в числовій колонці) відсікаються до типізації правилами tblRowFilters (6.7). Рядки, яких правила не охопили, дають помилки типізації; після типізації оригінальний текст у даних не зберігається.

### 6.7. Службові рядки, поріг помилок і індикатор якості

**tblRowFilters (fn\_ApplyRowFilters).**

- Рядок відсікається, якщо збігається хоча б одне ввімкнене правило його призначення. Порівняння нечутливе до регістру й ігнорує пробіли на краях.
- Оператори: Equals (значення дорівнює одному з варіантів Value), StartsWith, Contains, IsBlank (клітинка порожня), NotNumeric (непорожня клітинка не читається як число в культурі cfgSettings). Перелік має збігатися в tblOperator, fn\_ApplyRowFilters і pq\_RowFiltersCheck.
- Кілька варіантів у Value розділяються символом `|`: `Всього|Разом`, `х|x` (кирилична й латинська літера). Для IsBlank і NotNumeric Value не потрібне.
- NotNumeric приховує й справжні помилки даних, тому використовуйте його лише для колонок, де нечислові рядки точно службові.
- Правило діє на колонках джерела до вибору колонок схеми, тому може посилатися й на колонку, якої немає в схемі. Якщо колонки немає у файлі, правило пропускається без помилки, а FilterRulesSkipped у зведенні це показує.
- Use розбирається як fn\_ToBool: порожнє або значення з cfgSettings[TrueValues] вмикає правило, будь-яке інше (FALSE, Ні, 0) вимикає. Однаково в конвеєрі й у test\_RowFiltersReal.
- Кількість відсічених рядків видно в RowsDropped у pq\_QualitySummary. EmptyRowsRemoved тепер рахує лише повністю порожні рядки.
- Перед зміною правил варто запустити test\_RowFiltersReal: він показує, які значення відсікаються (DroppedValues, DroppedSample) і які нечислові значення лишаються (KeptNonNumeric).
- Без 5-го параметра rowFilters fn\_ReadPurpose службові рядки не відсікає.

**ErrorThreshold (fn\_ToRate).** Задається в cfgSettings (мовою M: `0.05` або `"5%"`), в options виклику fn\_ReadPurpose або в колонці ErrorThreshold реєстру.

| Значення в клітинці | Результат |
| --- | --- |
| 5% у відсотковому форматі (Excel зберігає 0,05) | 5% |
| число 0,05 | 5% |
| текст `5%`, `0,05`, `0.05`, `12,5 %` | 5%, 5%, 5%, 12,5% |
| `1` або `100%` | 100% (поріг ніколи не спрацює) |
| порожньо | значення з options або cfgSettings; якщо там null, поріг вимкнено |
| `off`, `ні`, `no`, `-`, `false`, FALSE | поріг вимкнено для джерела, навіть якщо в cfgSettings він заданий |
| `5`, `"5"`, `50`, 5000% (=50), від'ємне, понад 100%, `abc`, TRUE | помилка НекоректнийПоріг, рядок у pq\_RegistryCheck, QualityFlag = BAD\_THRESHOLD |

Число або текст без знака `%` завжди читається як частка від 0 до 1, щоб помилково введене значення не перетворювалося тихо на інший поріг. Для колонки рекомендовано відсотковий формат і Data Validation «Десяткове число від 0 до 1». Таке обмеження не дозволить ввести `off`, тому, якщо вимкнення для окремих джерел потрібне, валідацію не ставте, а за помилками стежте в pq\_RegistryCheck.

**QualityFlag у pq\_QualitySummary** прогнозує результат завантаження через fn\_ReadPurpose:

| Значення | Коли |
| --- | --- |
| FAIL | Завантаження зупиниться: перевищено поріг (Report, Lenient), є помилки при Strict, бракує обов'язкової колонки, файл недоступний чи не читається, немає схеми |
| WARN | Помилки є, але завантаження пройде |
| OK | Помилок немає |
| BAD\_THRESHOLD | Некоректне значення ErrorThreshold |
| FAIL (з v4.6) | Також невідоме значення ErrorPolicy: завантаження зупиняється з НевідомаПолітика |
| порожньо | Даних немає (NODATA) |

ErrorRate = ErrorRows / Rows (рядки після відсікання службових і порожніх). QualityFlag і зупинку в fn\_EnforceSchema визначає одна функція fn\_LoadVerdict, тому прогноз не може розійтися з реальною поведінкою. Статуси, їхній порядок і «блокуючі» статуси описано в cfgStatuses; структуру трьох звітів — у cfgReportTypes (нова колонка звіту додається там і там, де її обчислюють).

### 6.8. Макроси VBA

| Модуль, процедура | Призначення |
| --- | --- |
| modEngineCommon (з v4.8) | Спільні процедури eng\*: параметри з cfgSettings (engLoadSettings, engSetting, engTableName, engIsExtension, engToBool), таблиці (engFindTable, engCell, engCellOpt, engSourcePath), запити (engQueryNameOf за Location=, engFindQueryTable, engLoadQueryToSheet, engRefreshAllSync), текст M (engStripMComments, engMIdentifiers), файли (engFileExists, engSafeFileName, engEnsureFolder, engReadUtf8, engWriteUtf8) |
| modSelectSourceFiles, SelectSourceFiles | Вибір файлів для рядків tblSourceRegistry, перевірка структури CSV, запис FullPath, оновлення і підсумок QualityFlag |
| modRunLog, RefreshWithRunLog (з v4.8) | «Оновити все» з вимірюванням часу і записом журналу запусків у tblRunLog |
| modExportPowerQueryCodesToRepo, ExportPowerQueriesToRepo | Експорт для Git: код запитів (.pq), queries.yml (тип, таблиця завантаження, залежності), таблиці конфігурації в CSV |
| modPOWERQUERYPROFILER, ProfilePowerQueries | Час оновлення кожного підключення (аркуші PQ\_Profiler і PQ\_ProfilerSummary) |
| modCleanExportSelectedSheets, CleanExportSelectedSheets | Експорт виділених аркушів у окремі книги без запитів і підключень; таблиці запитів від'єднуються від QueryTable |
| modImportPowerQueryFromFolder, ImportPQQueriesFromFolder | Імпорт або оновлення запитів з теки .pq (ім'я файла = ім'я запиту) |
| modSetupTestEngine, SetupTestEngine | Підготовка тестової книги з пакета (розділ 6.11) |
| modTableCatalog, UpdateHelp (з v4.11) | Каталог таблиць tblTableCatalog і загальна довідка hlpProject на аркуші HELP, підписи над таблицями, перевірка правил розміщення (розділ 6.14) |

modExport (ExportSheetToCleanWorkbook) і старий модуль вибору файлів (MainModuleBas, ConfigTable) замінено й можна видалити. Модулі .bas зберігаються в кодуванні Windows-1251 (CRLF) для імпорту в редакторі VBA.

**Встановлення.** modEngineCommon імпортується першим: решта модулів, крім modImportPowerQueryFromFolder і modSetupTestEngine, викликають його процедури. Ці два навмисно самодостатні: ними збирають книгу з нуля, коли спільного модуля ще немає. Перед імпортом нової версії старий модуль видаляють (інакше редактор створить копію з суфіксом 1), після імпорту — Debug → Compile VBAProject.

**SelectSourceFiles: порядок роботи.**

1. Передумова: у tblSourceRegistry уже є стільки рядків, скільки файлів потрібно для роботи, із заповненими SourceID, TargetType, PeriodType, PeriodNumber, IsRequired, Enabled. Макрос рядків не додає й не видаляє; без активних рядків він зупиняється з підказкою.
2. Обробляються лише активні рядки: TargetType заповнений, Enabled порожнє або TRUE (як fn\_ActiveSources). Рядки з Enabled = FALSE пропускаються.
3. Якщо частина файлів уже задана, макрос пропонує вибрати заново для всіх рядків або лише для рядків без дійсного файла (шлях порожній або файла немає).
4. Діалог показує SourceID, TargetType, період і позначку «обов'язковий» та відкривається в теці поточного файла рядка. Скасування для необов'язкового рядка пропускає його, для обов'язкового пропонує повторити або вийти.
5. Для CSV перевіряються обов'язкові колонки схеми (tblSchemaRegistry: IsRequired = TRUE, Use не FALSE) з точним збігом назв, як у Engine. Роздільник береться з колонки Delimiter реєстру, інакше з cfgSettings[CsvDelimiter]; кодування з BOM файла, інакше з колонки Encoding, інакше з cfgSettings[CsvEncoding]. Якщо колонок бракує, попередження показує прочитаний заголовок, використане кодування і підказку щодо Encoding чи Delimiter; файл можна замінити, залишити або вийти. Excel-файли перевіряє Engine після оновлення (pq\_SchemaIssues, pq\_QualitySummary).
6. Підсумок перед записом: файл кожного рядка й попередження (обов'язковий без файла, рядки без періоду, той самий файл і лист у кількох рядках призначення, тобто майбутня помилка ДублюванняФайлу). Скасування на будь-якому кроці лишає таблицю без змін.
7. Після підтвердження записується FullPath, а Path і FileName цього рядка очищаються. PeriodType і PeriodNumber макрос не змінює: перевірте їх навпроти кожного файлу.
8. За бажанням макрос оновлює всі запити й показує підсумок QualityFlag з pq\_QualitySummary (OK, WARN, FAIL і SourceID, для яких завантаження зупиниться).

Ручне редагування FullPath або Path + FileName у tblSourceRegistry так само допустиме.

**Єдине джерело параметрів.** CsvDelimiter, CsvEncoding, TrueValues, імена таблиць (Tables) і розширення файлів (Extensions) макроси читають з тексту запиту cfgSettings (ThisWorkbook.Queries("cfgSettings").Formula) через engLoadSettings під час кожного запуску, без коментарів M. Змінювати їх треба лише в cfgSettings. Якщо запит або ключ не знайдено, макрос попереджає й пропонує продовжити з аварійними значеннями (;, 1251, стандартні TrueValues, стандартні імена таблиць і розширення). Фільтр діалогу SelectSourceFiles і перелік дозволених форматів будуються з cfgSettings[Extensions]. Перенести параметри в таблицю Excel, яку читали б і cfgSettings, і макрос, не можна: cfgSettings використовують функції всередині запитів, що відкривають файли, і читання книги в ньому повернуло б Formula.Firewall. Логіку (правила fn\_ToBool, fn\_SourcePath, fn\_ActiveSources) макрос повторює, бо VBA не може викликати функції M; відповідники позначено коментарями в модулі.

**Кодування CSV.** Файли з облікової системи мають мітку UTF-8 (BOM), тому Engine читає їх правильно попри CsvEncoding = 1251: мітка має пріоритет (підтверджено 15.10.2026: макрос із визначенням BOM знаходить колонки, без нього ні). Файл UTF-8 без мітки за CsvEncoding = 1251 дасть спотворені назви колонок; тоді задайте Encoding = 65001 у рядку реєстру.

**RefreshWithRunLog: журнал запусків.**

1. Якщо pq\_QualitySummary не завантажено на аркуш, макрос пропонує завантажити його на аркуш «Якість» у A3 (одноразово) — до оновлення, щоб не витрачати час даремно.
2. Синхронне «Оновити все» (фонове оновлення тимчасово вимикається), потім явне оновлення таблиці pq\_QualitySummary: журнал не залежить від позначки «Оновлювати під час оновлення всіх».
3. Для кожного рядка зведення в tblRunLog (аркуш RunLog, створюється при першому запуску, шапка в рядку 3) дописується рядок: RunID, RunAt, DurationSec, RefreshError, SourceID, TargetType, PeriodType, PeriodNumber, Status, QualityFlag, Message, Rows, ErrorRows, RowsDropped, FilePath, FileModified, FileSizeKB, User. Шлях — за тим самим правилом, що fn\_SourcePath; дата зміни й розмір беруться на момент запуску (порожні, якщо файла немає).
4. Підсумкове повідомлення: тривалість, OK/WARN/FAIL і SourceID, для яких завантаження зупиниться.

Історія накопичується; старі рядки можна видаляти вручну. Колонки, яких немає в tblRunLog, макрос пропускає.

**ExportPowerQueriesToRepo: експорт для Git.** Структура поруч із книгою: powerquery\queries\\*.pq (код запитів, UTF-8), powerquery\metadata\queries.yml, powerquery\tables\\*.csv. Перед експортом старі .pq і .csv видаляються, тож Git бачить і видалені запити. У CSV потрапляють лише таблиці, які заповнюють вручну (SourceType = xlSrcRange): реєстр, схема, фільтри, довідники, описи; таблиці-результати запитів і журнал tblRunLog — ні (EXCLUDED\_TABLES). CSV: роздільник «;», числа з крапкою, дати ISO, формули як формули. queries.yml: group (тека з cfgQueryGroups), type (Function, Config, Source, Workbook, Transform, Standalone), loads\_to\_table (фактичне ім'я таблиці на аркуші), depends\_on — за точними ідентифікаторами M без коментарів і текстових літералів. З v4.11 також powerquery\README.md — довідка проєкту, згенерована з книги: правила (hlpProject, за розділами), таблиці (tblTableCatalog: Таблиця, Аркуш, Тип, Запит, Призначення, Довідка — без мінливих Діапазон/Рядків/Порушення), hlp-довідки до таблиць (кожна як таблиця Markdown), запити за теками з посиланнями на .pq. Дати в README немає, щоб Git бачив лише змістовні зміни. Вручну README не редагують. Також з v4.11: модулі VBA в powerquery\vba\ (VBComponent.Export, cp1251; модулі аркушів і книги — лише якщо в них є код; потрібен «Довіряти доступ до об'єктної моделі проєкту VBA», без нього й для захищеного паролем проєкту тека vba не чіпається), розділ «Макроси VBA» в README (модуль, макроси Alt+F8, опис із заголовка); tblTableCatalog більше не вивантажується в CSV (мінливі колонки; зміст — у README). Репозиторій: .gitattributes зберігає .bas/.cls/.frm у Git в UTF-8, на диску в cp1251 і CRLF (working-tree-encoding, Git 2.18+; модулі не завантажувати через вебінтерфейс GitHub; якщо .bas уже були закомічені до .gitattributes — `git add --renormalize .`); CSV, yml, json і фікстури — байт у байт. .gitignore: книги в корені, тимчасові файли Excel, теки з реальними даними. Фікстури тестів (tests\fixtures, test\_config.xlsx, gen\_fixtures.py) копіюються в репозиторій вручну з пакета, не макросом.

**ProfilePowerQueries: профайлер.** Ім'я запиту береться з Location= рядка підключення (правильне і в українському Excel, де підключення називається «Запит – X» з нерозривним пробілом). Підсумок запуску — рядок на аркуші PQ\_ProfilerSummary. Сортування: новіші запуски вгорі, у межах запуску найдовші вгорі. Кожне підключення оновлюється окремо, тому його час включає обчислення всіх запитів, від яких воно залежить; сума тривалостей більша за час «Оновити все».

### 6.9. Правило регістру і довідники

Значення довідників порівнюються без урахування регістру й пробілів на краях: Registry, REGISTRY і regIStrY — те саме призначення. Причина: Data Validation зі списком із довідника приймає значення в будь-якому регістрі.

| Значення | Довідник (джерело Data Validation) | Як забезпечено |
| --- | --- | --- |
| TargetType | tblTargetType | fn\_Canon на вході (fn\_ActiveSources, fn\_NormalizeSchema, fn\_NormalizeRowFilters, fn\_ReadPurpose); порівняння в fn\_GetSchemaRules, fn\_TemplateColumns, макросі |
| PeriodType | tblPeriodRules | fn\_Canon у fn\_ActiveSources |
| DataType | tblDataType | fn\_TypeKey (нижній регістр) |
| Логічні значення (IsRequired, Enabled, Use, TrimFlag) | tblLogical | fn\_ToBool і ToBool у макросі (нижній регістр, cfgSettings[TrueValues]) |
| CustomTransform | tblCustomTransform | fn\_GetSchemaRules, pq\_SchemaCheck |
| ErrorPolicy | tblErrorPolicy | Text.Proper у fn\_EnforceSchema, fn\_RowReport |
| Operator | tblOperator | нижній регістр у fn\_ApplyRowFilters, pq\_RowFiltersCheck |

- fn\_Canon приводить TargetType і PeriodType до написання з довідника, тому звіти, дані (SourceID, PeriodType у результаті fn\_ReadPurpose) і зв'язки в моделі Power Pivot отримують одне написання. Значення поза довідником лишається як є і з'являється в pq\_RegistryCheck («TargetType немає в tblTargetType»). Data Validation такого не пропускає, але вставлення з буфера обміну Data Validation обходить, тому перевірку залишено.
- Не довідники й лишаються чутливими до регістру: назви колонок файлів (SourceColumnName, TargetColumnName), SourceID, Value в tblRowFilters (там порівняння й так нечутливе до регістру), шляхи до файлів.
- У довіднику не повинно бути двох значень, що відрізняються лише регістром: fn\_Canon візьме перше.

### 6.10. Загальне

- Оператор ?? потребує сучасної версії Power Query.
- Код надавався без запуску з боку автора відповідей, тому статуси перевірки відображають лише повідомлення користувача.
- Обсяг прочитаних даних наведено за візуальними спостереженнями, а не за вимірюваннями: до виправлення Strict файл на 68 МБ читався приблизно десять разів (706 МБ); після виправлення Strict зупиняється приблизно на 40 МБ, Report читає 80–90 МБ, повне завантаження 199 тис. рядків займає менше хвилини.


### 6.11. Регресійні тести (етап C)

- Тестова книга — окрема копія Engine (test\_engine\TEST\_POWER\_QUERY\_ENGINE.xlsm пакета); її реєстр, схема, фільтри й довідник призначень тестові (TestCsv, TestXlsx, TestNoSchema). Не використовуйте її для реальних даних.
- Підготовка й оновлення коду: макрос SetupTestEngine (modSetupTestEngine) з розпакованого пакета; повторний запуск безпечний і так само оновлює запити після нових патчів.
- Запуск: «Оновити все», потім аркуш «Тести»: у test\_All усі рядки PASS. test\_Summary (зведення проти tblTestExpected) і test\_Load (fn\_ReadPurpose проти tblTestLoad) показують кожну перевірку з очікуваним і фактичним значенням.
- Правило роботи: перед кожною зміною коду Engine і після неї — прогін тестів; зміну переносимо в робочу книгу лише після зеленого test\_All. Новий сценарій — файл у fixtures, рядок у реєстрі й очікування, без зміни коду (і в test\_config.xlsx, щоб пережив повторний SetupTestEngine).
- Набір v4.6: 21 тестовий файл, 24 сценарії зведення (84 перевірки), 25 сценаріїв завантаження, модульні test\_ToRate (24), test\_Canon (7), test\_LoadVerdict (12), test\_RowFilters (1).
- З v4.7–v4.8 додано test\_RegistryIssues (9) і test\_Extensions (7): test\_All підсумовує 8 наборів.
- Для журналу запусків і профайлера в тестовій книзі з'являються аркуші «Якість» (pq\_QualitySummary), RunLog, PQ\_Profiler, PQ\_ProfilerSummary; на тести вони не впливають.
- Очікування C13 (UTF-8 без BOM при CsvEncoding = 1251 дає спотворені заголовки) фіксує поточну поведінку, а не бажану.

### 6.12. Продуктивність: повторне читання файлів

- У M таблиця, на яку запит посилається кілька разів (фільтр, з'єднання, ToRecords, SelectRows у циклі), не кешується, а обчислюється заново при кожному перегляді. Якщо вона відкриває файли, кожен перегляд знову читає всі файли-джерела, а вкладені повтори множаться.
- Правило: у запиті, що збирає результати, вхідну таблицю, яку переглядають більше одного разу, обгортати Table.Buffer одразу при отриманні (pq\_ExtractedColumns, pq\_SchemaDrift, test\_Summary з v4.8). Table.Buffer не застосовувати до таблиць із клітинками-помилками, які мають лишитися помилками в окремих клітинках; там рахувати за один перегляд (test\_Load: помилки рахуються одним Table.AddColumn, а не SelectRowsWithErrors на кожну колонку).
- Ефект у тестовій книзі (17.10.2026): pq\_SchemaDrift читав 10–12 МБ при ~100 КБ тестових файлів (~100 повних проходів), після виправлення оновлення в редакторі майже миттєве; «Оновити все» з журналом — 34–39 с замість 502 с.
- Синхронне оновлення (макроси) робить довгі оновлення помітними; фонове «Оновити все» в Excel їх приховувало.
- Тести не повинні залежати від реальних джерел у реєстрі тестової книги: test\_Summary будує звіти лише для SourceID з tblTestExpected (v4.9, патч G). До цього 200 МБ реальних CSV у тестовій книзі давали 11–12 хв «Оновити все» (3–4 повні читання).
- Обсяг прочитаного в редакторі Power Query не дорівнює часу. Окремий Table.RowCount сирого CSV подвоює обсяг читання, але коштує мало; Table.Buffer сирого файла на всю ширину (252 колонки джерел проти 131 у схемі) коштує дорого. Патч H (буфер сирих даних замість окремого підрахунку) збільшив час pq\_QualitySummary з 217 до 336 с і відкочений. Оптимізації приймаються лише за вимірюванням часу (профайлер), не за обсягом читання.
- Звіт якості (v4.9, патчі J–L; вимірювання 23–24.10.2026 діагностичними запитами з позначками часу DateTime.LocalNow усередині одного запиту):
  - Читання і нормалізація найбільшого файла (GO3, 70 МБ, 81 тис. рядків, 16 колонок схеми) — ~3–4 с; ~90% часу звіту йшло на перевірки по колонках.
  - J: список перевірок колонок (ColList) переглядався чотири рази (таблиця колонок, таблиця помилок, порожні колонки, кількість колонок з помилками) — List.Buffer. Правило 6.12 стосується і списків: лінивий List.Transform, який переглядають кілька разів, обчислюється кілька разів.
  - K: пошук рядків з помилками окремо для кожної колонки по всій таблиці (GO3: 7,1 с) замінено одним відбором рядків з помилками і перевірками колонок лише по них (0,5 с).
  - L: підрахунок null окремо для кожної колонки (GO3: 7,1 с; KGO1, 21 колонка: 10,0 с) замінено одним Table.Profile (1,1 і 1,2 с; на KGO1 перевірено 44 342 null — результат той самий). Table.Profile і Table.ToColumns падають на колонці з клітинкою-помилкою (Registry1, #N/A), тому такі колонки рахуються старим способом (fn\_ColumnNullCounts, test\_NullCounts). Діставання колонки з буферизованої таблиці (Table.Column, Table.ToColumns) коштує ~0,4 с на колонку для 80 тис. рядків — дорожче за сам підрахунок.
  - Підсумок по джерелах (zz\_diag\_SourceCost): GO3 24,5 → 4,8 с, KGO1 33,6 → 4,8 с; сума 8 реальних джерел ~120–158 с → 22 с.
  - Цілком (профайлер, 24.10.2026): pq\_QualitySummary 202,8 → 67,7 с; RefreshWithRunLog 367,8 → 116 с. Розрив між сумою звітів (~25 с, один перегляд у діагностиці) і профайлером (~68 с) не пояснюється повторними переглядами в межах запиту: патч M (List.Buffer і Table.Buffer у fn\_SummaryTable) дав 70,0 с — без змін, відкочено. Імовірно, Excel виконує запит під час завантаження на аркуш кількома незалежними обчисленнями; з M на це не вплинути. ~70 с — практична межа звіту якості для 200 МБ; далі — рішення R3 (звіт під час кожного оновлення чи на вимогу).
  - Хибні кроки: H (буфер сирого файла на всю ширину) — гірше; M — без ефекту; прогноз «у 3–4 рази» для J — фактично −24%; етап «лише таблиця колонок» у першій діагностиці не обчислював перевірок (Table.RowCount(Table.Buffer(...)) не примушує обчислення полів), тому показав 1,7 с замість ~21 с. Висновок: діагностика має примушувати обчислення саме тих полів, які використовує реальний запит.
- Вимірювання 23.10.2026 (тестова книга, 200 МБ реальних CSV, 33 джерела): RefreshWithRunLog після холодного запуску 367,8 с (до G1 — 533 с, до G — 11–12 хв); профайлер: pq\_QualitySummary 217 с з 261 с суми, решта 40 підключень разом 44 с.

### 6.13. Правила після аудиту коду (v4.10)

- **ErrorPolicy, ErrorThreshold, Culture** джерела задаються лише в реєстрі (колонки рядка) або в cfgSettings. Параметр options у fn\_ReadPurpose їх не змінює (лише AllowDuplicatePeriods, AllowDuplicateFiles): звіт якості (QualityFlag, типізація) обчислюється без options, тож інакше прогноз розходився б із завантаженням. Необов'язкова колонка Culture у tblSourceRegistry: культура розбору чисел і дат для джерела, порожньо = cfgSettings[Culture].
- **Невідомі DataType і CustomTransform** виявляються під час компіляції правил (fn\_GetSchemaRules) за назвами, тож джерело отримує NO\_SCHEMA, а не READ\_ERROR чи тисячі помилкових клітинок. Збіг цільової назви з колонкою метаданих — помилка ЗбігЗМетаколонкою.
- **Клітинка-помилка** (#N/A з xlsx, помилка типу): вважається непорожньою при відсіканні порожніх рядків; ніколи не збігається з правилом фільтра службових рядків (рядок не зникає, а потрапляє у звіт як ERROR).
- **Фільтри службових рядків** порівнюють текст у культурі cfgSettings, а не машини; справжнє число ніколи не NotNumeric.
- **Джерела без періоду** (PeriodType і PeriodNumber порожні) не беруть участі в перевірці «дубль періоду»; той самий файл двічі ловить «дубль файлу».
- **EmptyColumns** містить лише колонки, що є у файлі; відсутні необов'язкові — лише в MissingOptional.
- **Заголовки файлів не обрізаються** (рішення 24.10.2026): зміну у вивантаженні має бути видно. Натомість pq\_SchemaDrift підказує «пробіли на краях назви», а помилка «відсутні обов'язкові колонки» називає заголовки з пробілами.
- **fn\_ParseHeaders** враховує BOM так само, як повне читання (test\_HeadersConsistency: 17 файлів, заголовки збігаються).
- **cfgSettings[MetaColumns]** — запис [Id, PeriodType, PeriodNumber] (імена за ключем, не за позицією); cfgSettings[Tables][Period] = tblPeriod (Calendar).
- **fn\_EnforceSchema**: для Strict і порогу помилок нормалізовані дані буферизуються один раз (замість 2–3 розборів файла).
- Обмеження: помилка, що виникає посеред читання файла (а не під час побудови таблиці), у fn\_ReadPurpose не отримує префікса [SourceID] — перехопити її без повного буфера даних неможливо.

**Профайлер (v4.10):** ручний режим обчислення під час заміру (раніше кожне завантаження таблиці перераховувало книгу: чисті тести в пам'яті займали 7–9 с, тепер 0,4–1,5 с); «лише підключення» не міряються; фоновий режим підключень відновлюється; необов'язковий прогрівальний прохід. Вимір 24.10.2026: pq\_QualitySummary 44,9 с після прогріву, 83,4 с холодний старт. Окремі запити без читання файлів інколи отримують +7–9 с незалежно від коду (імовірно, запуск нового службового процесу Power Query) — порівнювати лише повторювані заміри.

### 6.14. Структура книги і довідка (v4.11, обов'язкові правила)

Погоджено 24.10.2026. Ті самі правила макрос записує в таблицю hlpProject на аркуші HELP (якщо її ще немає).

**Розміщення**
1. Шапка кожної таблиці — не вище рядка 3. Рядок 2 порожній, рядок 1 — підпис таблиці (ставить UpdateHelp).
2. Таблиці ростуть униз і вправо, тому кілька таблиць на аркуші ставлять лише поруч, ніколи одна під одною.
3. Таблиця з запиту Power Query (q\_, pq\_) — одна на аркуші; аркуш бажано називати як запит.
4. Кожна таблиця займає щонайменше 3 колонки (місце для підпису); між сусідніми таблицями — щонайменше одна порожня колонка.
5. Нижче й правіше таблиці нічого не розміщують.

**Назви:** префікс задає роль — tbl (ручне введення, налаштування), q\_ (дані призначення), pq\_ (діагностика і звіти), hlp (довідка), test\_ (тести, лише тестова книга). Назви унікальні й завжди з префіксом; довідка до tblX за замовчуванням — hlpX.

**Довідка.** Аркуш HELP: tblTableCatalog, hlpProject і hlp-таблиці поруч, шапки в рядку 3. Призначення таблиці описують лише в tblTableCatalog[Призначення]; правила заповнення колонок — у hlp-таблиці, вказаній у tblTableCatalog[Довідка].

**Каталог** (14 колонок): №, Таблиця, Аркуш, Діапазон, Рядків, Колонок, Тип, Запит, У моделі даних, Призначення, Підпис, Довідка, Порушення, Колонки таблиці. Призначення, Підпис (TRUE/FALSE) і Довідка заповнює користувач, повторний запуск їх зберігає. Таблиця і Довідка — гіперпосилання на відповідну таблицю (Довідка — лише якщо hlp-таблицю знайдено, інакше текст). Порушення макрос лише показує: назва без префікса, шапка вище рядка 3, перетин блоків по колонках (зокрема таблиці одна під одною), немає порожньої колонки між блоками, таблиця запиту не одна на аркуші, довідку не знайдено, не заповнено призначення, рядок підпису зайнятий чужим вмістом.

**Підпис** — три формули в клітинках рядка (шапка − 2) над першими колонками таблиці (рішення 25.10.2026: без іменованих формул — імена рівня аркуша давали б багато імен з одним призначенням, а варіант рівня книги без назви аркуша (`!A3`) повертав порожнє значення). Excel 2021+: LET, XLOOKUP, IFS; LAMBDA і TEXTAFTER не потрібні. Формули пишуться через Formula2 англійською, Excel показує їх мовою інтерфейсу.
- назва (з приміткою «Назва таблиці»): аркуш — текст після останньої «]» у CELL("filename") (останнє входження замінюється на «|» через SUBSTITUTE), перша клітинка шапки — ADDRESS(ROW;COLUMN;4) (не CELL("address"), що в стилі R1C1 повертає R3C1); пошук у каталозі XLOOKUP за двома умовами;
- «ℹ Призначення» — HYPERLINK на клітинку «Призначення» рядка каталогу (адреса через ADDRESS(ROW(hit);COLUMN(hit);4): CELL("address") для клітинки іншого аркуша повертає адресу з назвою книги й аркуша і дає «посилання неприпустиме»); порожнє призначення — текст «ℹ Призначення (не заповнено)» без посилання;
- «? Як заповнювати» — IFS: HYPERLINK("#"&hlp) з колонки «Довідка», якщо таблиця існує (ISREF(INDIRECT(...))), інакше «⚠ … не знайдено»; для hlp-таблиці — «← до tblX».
Книга має бути збережена. Клітинки підпису позначаються прихованими іменами engLbl\_<таблиця>; наступний запуск їх очищає і пише заново. Вручну підписи не редагують.

**Службові таблиці** (Підпис = FALSE за замовчуванням): tblTableCatalog (у каталог не входить), hlpProject, tblRunLog, test\_…, tblTest….

**Нова таблиця з запиту:** завантажити на новий аркуш → UpdateHelp (додасть у каталог, вставить рядки, щоб шапка була в рядку 3, поставить підпис; «Призначення» отримає стартовий опис) → заповнити «Призначення», за потреби створити hlpX і вписати в «Довідку». Після вставлення рядків чи колонок перед таблицею, перейменування таблиці чи аркуша — знову UpdateHelp, інакше підпис покаже «⚠ немає в каталозі». Запити лише в моделі даних (без таблиці на аркуші) у каталог не потрапляють.

**Що макрос змінює сам:** створює аркуш HELP і каталог, якщо їх немає; вставляє рядки зверху аркуша, де шапка якоїсь таблиці вище рядка 3; приводить каталог до 14 колонок у стандартному порядку (якщо праворуч немає місця — зупиняється з повідомленням); створює hlpProject праворуч від усього на HELP. Чужі таблиці не пересуває.

## 7. Результати перевірки

Зведено за вашими тестами на реальних даних.

| Сценарій | Результат |
| --- | --- |
| Дубль SourceID | fn\_ReadPurpose зупиняється (ДублюванняSourceID), pq\_RegistryCheck показує дубль |
| Дубль періоду | Зупинка з помилкою ДублюванняПеріоду |
| Однаковий FullPath у різних періодах | Помилка ДублюванняФайлу незалежно від IsRequired; pq\_RegistryCheck показує дубль лише для активних рядків |
| Шлях не задано | Повідомляється; для необов'язкового джерела статус NODATA |
| Обов'язковий файл відсутній | Зупинка з помилкою \[ФайлНедоступний\] \[SourceID\] |
| Необов'язковий файл відсутній | Без помилки, у даних немає рядків цього джерела, статус NODATA |
| Усі джерела необов'язкові | Правильні шляхи: дані завантажено; зламані шляхи: без помилок, NODATA для кожного |
| Вимкнене джерело | Зникає з даних, pq\_QualitySummary і перевірки реєстру |
| Порожній результат | Усі колонки схеми в порядку схеми, SourceID, PeriodType, PeriodNumber, типи, без рядків |
| Повністю порожній реєстр | Без попередження Excel про запит без стовпців; pq\_RegistryCheck показує «SourceID порожній» і «Дубль TargetType + період» для порожніх рядків |
| Excel: таблиця за іменем | Заголовки з таблиці |
| Excel: збіг імен таблиці й аркуша | Береться таблиця (відбір за Item), регістр не важливий |
| Excel: неіснуюче ім'я | Статус у pq\_SchemaIssues; без шаблону колонок немає, із шаблоном колонки з шаблону |
| Excel: повний цикл зі схемою | Дані завантажено |
| Strict: 17 826 рядків з помилками в одній колонці | Зупинка з повідомленням, дані не завантажено, звіти без змін |
| Report: ті самі дані | Усе завантажено, помилки лишаються в клітинках (на аркуші порожні), звіти показують 17 826 помилок |
| Lenient | До v4.6 не працював: помилки лишалися в клітинках (тест L06, ErrorCells = 2). Після виправлення ErrorCells = 0 |
| tblRowFilters: test\_RowFilters (вбудовані дані) | Dropped = 3, RulesApplied = 2, RulesSkipped = 1, OK = TRUE |
| tblRowFilters: реальний файл Registry (test\_RowFiltersReal) | 440 рядків, відсічено 22 («Всього», «Разом» у № з/п, «х» у Ціна, грн.), RulesSkipped = 0, KeptNonNumeric порожній |
| tblRowFilters у конвеєрі | pq\_RowFiltersCheck порожній; Registry1: RowsDropped = 22, FilterRulesSkipped = 0, ErrorRows = 0; q\_Registry без службових рядків і помилок; інші призначення без змін |
| Обидва правила вимкнено, Strict | Зупинка ПомилкиТипізації (22 рядки, колонки № з/п, Ціна); зведення: ERROR, ErrorRate 5%, QualityFlag = FAIL |
| Обидва правила вимкнено, Report, поріг 1% | q\_Registry зупиняється ПеревищенийПоріг; QualityFlag = FAIL |
| Обидва правила вимкнено, Report, поріг 50% | q\_Registry завантажено, клітинки з помилками на аркуші порожні; QualityFlag = WARN |
| ErrorThreshold = abc | QualityFlag = BAD\_THRESHOLD, ErrorThreshold = null |
| ErrorThreshold = 5000% | До виправлення читалося як 50%; після виправлення fn\_ToRate це помилка (test\_ToRate, 24 випадки, усі Pass = TRUE) |
| SelectSourceFiles: скасування на підсумку | Таблицю не змінено |
| SelectSourceFiles: Enabled = FALSE | Рядок пропущено |
| SelectSourceFiles: CSV іншого призначення | Перелік відсутніх колонок |
| SelectSourceFiles: той самий CSV (UTF-8 з BOM) | До виправлення хибне попередження (читання як 1251); після визначення BOM попередження немає |
| SelectSourceFiles: запис і оновлення | Підсумок макроса збігся з pq\_QualitySummary |
| CleanExportSelectedSheets: аркуш із таблицею запиту | Новий файл відкривається без попереджень |
| CSV: перейменування колонки в дубль обов'язкової (до B1) | Обов'язкова колонка зникла: макрос і Engine показують відсутню колонку |
| CSV: тихий дубль колонки схеми (до B1) | Power Query взяв першу (хибну) колонку, дані мовчки завантажено в q\_Registry; повтор отримав суфікс \_1 |
| CSV: тихий дубль колонки схеми (після B1) | Макрос попереджає; pq\_QualitySummary: Status = WARN, DuplicateHeaders з назвою; pq\_SchemaDrift: колонка з \_1 «Немає у схемі» |
| Excel: =NA() у колонці «Кількість», Report | ErrorRows = 1, ErrorColumns = 2 (разом із сумою-формулою), QualityFlag = WARN; q\_Registry завантажено, клітинки порожні |
| Файл відкритий в Excel або Блокноті під час оновлення | Оновлення без помилок |
| Правило регістру (test\_Canon, 7 випадків) | Усі Pass = TRUE |
| TargetType regIStrY у реєстрі, REGISTRY у tblRowFilters | Дані й звіти як для Registry; RowsDropped без змін |
| PeriodType «квартал» і «Квартал» з однаковим номером | pq\_RegistryCheck показує дубль, завантаження зупиняється ДублюванняПеріоду |
| SourceColumnName у схемі в іншому регістрі | pq\_SchemaDrift: «Немає у схемі: регістр відрізняється» і «…: регістр відрізняється» |
| TargetType поза tblTargetType | Не перевірено: Data Validation не дозволяє ввести |
| Етап C, перший прогін (16.10.2026) | 140 з 141 перевірки; FAIL лише L06 (Lenient), виправлено патчем C1 |
| Етап C після C1 і патча D (аудит D4–D7) | test\_All: усі набори PASS (6 наборів) |
| Пошкоджений Excel-файл (X05) | READ\_ERROR; завантаження зупиняється з DataFormat.Error |
| Патч G: тести лише для тестових SourceID (200 МБ реальних CSV) | test\_All PASS; RefreshWithRunLog 533 с (з подвійним оновленням зведення F3) |
| Патч G1: без другого оновлення зведення | RefreshWithRunLog 367,8 с після холодного запуску; FAIL лише в очікуваних тестових сценаріях |
| Патч H: буфер сирих даних | pq\_QualitySummary 336 с замість 217 с — відкочено |
| Патч H2: вужчий буфер перед фільтрами | pq\_QualitySummary 202,8 с (−7%) — залишено |
| Профайлер робочої книги без тестів (24.10.2026, Run 1) | pq\_QualitySummary 28,4 с (усі 8 реальних джерел, ~200 МБ), q\_GO 21,6 с, pq\_ExtractedColumns 2,5 с, Calendar 1,9 с; разом ~55 с |
| q\_GO = fn\_ReadPurpose(fn\_LoadConfig(), "GO") на аркуш (24.10.2026, v4.10) | 3 файли (GO1–GO3, ~170 МБ, ~200 тис. рядків), 3 796 542 клітинки — ~20 с разом з автопідбором ширини колонок |
| Патч N (аудит коду): 1.1–1.8, єдине джерело | test\_All 11 наборів; test\_AuditFixes виявив, що виправлення 1.2 у N не працювало (помилка клітинки виникала до try) |
| Патч N2: 1.2 (try), підказки про пробіли, test\_RowFilters 13 випадків, test\_Load одне завантаження | test\_All 11 наборів PASS |
| Патч N3 (повторний аудит): 2.1–2.5, профайлер, експорт | test\_All 11 наборів PASS (test\_AuditFixes 15/15); профайлер — див. 6.13 |
| Патчі J, K, L (24.10.2026) | test\_All: 9 наборів PASS (test\_NullCounts 7/7); звіт якості 8 реальних джерел: 22 с замість ~120 с (GO3 4,8 с, KGO1 4,8 с) |
| Патч E (аудит D2, D10) | test\_All: 7 наборів PASS, зокрема test\_RegistryIssues (8 правил і загальна кількість); pq\_RegistryCheck показує дублі C01/C16 і C01/C17 з Blocking = TRUE |
| test\_RowFiltersReal через fn\_LoadConfig | Formula.Firewall; виправлено прямим читанням таблиць з іменами з cfgSettings[Tables] |
| Правила «х» і «Всього» в одному рядку | Вимкнення одного правила не змінює результат: рядок відсікає інше. Ізольована перевірка правила «х» на цьому файлі неможлива |
| Data Validation: вибір TargetType | Список SourceColumnName містить колонки лише цього призначення; зміна TargetType змінює список |
| Data Validation: порожній TargetType або призначення без tblSchemaColumns | Кнопка списку неактивна, наявні записи лишаються |
| Етап D, патч F (17.10.2026) | Компіляція VBA без помилок; test\_All: 8 наборів PASS, зокрема test\_Extensions |
| SelectSourceFiles після переведення на modEngineCommon | Фільтр діалогу з cfgSettings[Extensions]; непідтримуваний файл (.pdf) відхиляється з переліком дозволених форматів |
| RefreshWithRunLog, перший запуск (V4) | 502 с, журнал не записано: pq\_QualitySummary не був на аркуші, перевірка йшла після оновлення. Виправлено патчем F1 (перевірка до оновлення, завантаження на аркуш «Якість»); причину тривалості — патчем F2 |
| RefreshWithRunLog після F1, F2 (V4, V5) | 34 с і 39 с; RunID 1 і 2, попередні рядки збережено |
| Неіснуючий файл у C01 (V6) | Журнал: FilePath заповнено, FileModified і FileSizeKB порожні. Таблиця pq\_QualitySummary на аркуші лишалася застарілою (OK), бо в підключенні не було позначки «Оновлювати під час оновлення всіх»; Engine правильно давав FILE\_ERROR. Виправлено патчем F3 і позначкою |
| Профайлер (V7, V8) | 41 підключення, сума 31–37 с, найдовший test\_All; імена запитів без «Запит – »; сортування правильне; старі рядки SUMMARY видаляються з повідомленням |
| Експорт для Git (V9–V11) | 59 .pq, 20 CSV таблиць конфігурації (без таблиць-результатів і tblRunLog); queries.yml: точні залежності (fn\_Options → fn\_Opt, cfgSettings), тип Function, loads\_to\_table; видалений запит зникає з теки після повторного експорту |
| CleanExportSelectedSheets (V12) | «a:b» в імені файла → «a\_b» |

## 8. Відкриті питання і подальший розвиток

1. ~~Перевірити політику Lenient~~: перевірено тестом L06, ваду виправлено (v4.6).
2. ~~Перевірити решту статусів fn\_RowReport~~: усі статуси покрито регресійними тестами (v4.6).
3. Ігнорування повністю порожніх рядків реєстру у fn\_ActiveSources: запропоновано, застосування не підтверджено. До цього порожні рядки таблиці дають «SourceID порожній» і «Дубль TargetType + період» у pq\_RegistryCheck.
4. ~~Колонка ErrorPolicy у tblSourceRegistry~~: реалізовано й перевірено (v4.3).
5. Нумерація RowNo до видалення порожніх рядків, якщо потрібна відповідність рядкам файла (змінить fn\_ApplyRules і fn\_SchemaReport).
6. ~~Відсікання службових рядків до типізації~~: реалізовано (tblRowFilters, v4.3). Нові службові рядки в CSV додаються правилами за потреби.
7. Колонка ColumnOrder у tblSchemaRegistry (порядок колонок, стійкий до сортування таблиці).
8. Перевірка «шаблон проти схеми» як окремий звіт.
9. Вибір Excel-таблиці за структурою (за обов'язковими колонками схеми), коли TableSheetName порожнє.
10. Видалити fn\_ResolveFolderFile, якщо вона ще є в книзі.
11. DB-джерела: у концепції, не реалізовано.
12. ~~Етап B: крайні випадки джерел~~: завершено (v4.5). BasePath (відносні шляхи) відкладено: файли локальні й можуть бути в різних теках, а SelectSourceFiles записує повні шляхи й виявляє недійсні; повернутися, якщо тека проєкту з даними переноситиметься цілком.
13. ~~Етап C: регресійні тести~~: завершено (v4.6), див. розділ 6.11.
14. ~~Етап D: журнал запусків і контроль версій коду~~: завершено (v4.8), див. розділ 6.8.
15. Перевірка структури Excel-файлів у SelectSourceFiles (TableSheetName, заголовки) до оновлення; зараз її робить лише Engine.
16. Аудит дублювань (duplication\_audit\_v4.4.md): D1, D3, D9 закрито у v4.5; D4–D7 закрито у v4.6; D2 і D10 закрито у v4.7; D13, D8 (розширення файлів) і помилки профайлера закрито у v4.8. Відкриті: D8 для операторів фільтрів, політик, типів і дій CustomTransform (перевірка таблиць Data Validation проти M), D11, D14.
17. Перевірити на реальних даних робочої книги зміни v4.6–v4.8 (патчі D, E, F–F3), зокрема тривалість «Оновити все» і pq\_SchemaDrift після буферизації.
18. test\_All заново обчислює test\_Summary і test\_Load, які також завантажені на аркуш, тож у тестовій книзі тести виконуються двічі за «Оновити все». Прийнятно (≈30 с); за потреби test\_All можна завантажувати на аркуш без двох інших.
19. Решта запитів, що збирають результати (qc\_Reports-залежні звіти, pq\_SchemaIssues), переглядають вхід один раз; правило 6.12 застосовувати до кожного нового запиту.
20. ~~Патч H2~~: вужчий буфер перед фільтрами службових рядків — залишено: pq\_QualitySummary 202,8 с проти 217 с (−7%, близько до розкиду між запусками; виграш лише для призначень із фільтрами, зараз Registry). Решта підключень: test\_All 15,8 с, test\_Summary 11,6 с, test\_Load 6,2 с, pq\_ExtractedColumns 3,5 с — R1 для CSV не потрібен.
22. ~~Вартість перевірок по колонках у звіті якості~~: виміряно діагностичними запитами (zz\_diag\_ReportCost, zz\_diag\_SourceCost, zz\_diag\_ColumnCost) і усунуто патчами J, K, L (розділ 6.12). Профайлер після J–L: pq\_QualitySummary 67,7 с (було 202,8 с).
21. R3 зовнішнього аудиту: звіт якості під час кожного оновлення чи на вимогу. Цифри робочої книги (24.10.2026): звіт якості 28 с на всі джерела, q\_GO 22 с. Рекомендація: лишити звіт якості в «Оновити все» (він коштує приблизно одне читання всіх джерел і не множиться з кількістю q\_…); переглянути, коли повне оновлення перевищить ~2 хв — тоді виключити pq\_QualitySummary з «Оновити все» і оновлювати його через RefreshWithRunLog.
23. Одне розбирання файла у звіті якості (аудит 3.1): не впроваджено. Окремий Table.RowCount сирого CSV коштує ~3–4 с з ~45 с pq\_QualitySummary (<10%), а буферизація замість нього вже давала гірший результат (патч H). Повернутися лише з вимірюванням.
24. Тест на файл 0 байт: потрібні новий тестовий файл, рядок у test\_config.xlsx і повторний SetupTestEngine.
25. cfgRowOperators (аудит D8, решта переліків): оператори фільтрів досі описані в fn\_ApplyRowFilters, pq\_RowFiltersCheck і tblOperator.
26. fn\_Canon читає tblTargetType кілька разів за одне завантаження (мілісекунди) — за потреби передавати готові функції Canon через fn\_LoadConfig.
27. BasePath (Path як іменована клітинка з базовим каталогом) — відкладено рішенням користувача.
28. engLoadQueryToSheet і завантаження нових q\_: ставити таблицю на окремий аркуш у A3 відразу (зараз Excel ставить у A1, і рядки вставляє UpdateHelp).

## 9. Журнал змін

| Версія | Дата | Опис змін |
| --- | --- | --- |
| v1.0 | - | Початкова версія: концепція, три шари, дорожня карта |
| v1.1 | - | Структура tblSourceRegistry; розділення на tblSourceRegistry і tblSchemaRegistry |
| v2.0 | - | Core ETL Engine та інспектор схем: pq\_SchemaDrift, pq\_ExtractedColumns, fn\_TemplateColumns; fn\_ReadPurpose винесено в архів |
| v3.0 | - | Schema Enforcer: двостадійний конвеєр (Pre і Post), pq\_SchemaCheck, оновлений pq\_SchemaDrift, політики Strict, Report, Lenient, захист MetaColumns |
| v3.1 | - | fn\_SchemaReport: звіт із порожніми рядками, відсутніми колонками та зразками помилок (RowNo, Detail) |
| v4.0 | - | QC Suite (qc\_Reports, pq\_QualitySummary, pq\_QualityColumns, pq\_QualityErrors, fn\_RowReport), fn\_NormalizeSchema, оновлена fn\_ReadPurpose, видалено fn\_TemplateTable |
| v4.1 | 12.10.2026 | За результатами тестування: fn\_ActiveSources, fn\_SourcePath; перевірки дублів SourceID, файлів і періодів; перевірка першого байта в fn\_ResolveSource; відбір Excel за Item; захист порожніх результатів (pq\_GenerateSchemaTemplate\_Wide, pq\_QualityColumns, pq\_QualityErrors); pq\_SchemaIssues; Strict із буфером; правило трьох видів запитів. Документ перероблено: додано концепцію, дорожню карту, результати перевірки й відкриті питання |
| v4.2 | 13.10.2026 | Списки Data Validation: pq\_GenerateSchemaTemplate\_Wide вилучено; pq\_ExtractedColumns відсортовано блоками по призначенню й завантажено як tblSchemaColumns; джерелом списку SourceColumnName є іменована формула GetDependentList (LET і OFFSET за tblSchemaRegistry\[@TargetType\]); документ оновлено |
| v4.3 | 15.10.2026 | Етап A плану стійкості. tblRowFilters із fn\_NormalizeRowFilters, fn\_ApplyRowFilters і pq\_RowFiltersCheck; конвеєр fn\_ApplyRulesEx (фільтри першим кроком), fn\_ApplyRules як обгортка; rowFilters у fn\_GetSchemaRules і 5-м параметром fn\_ReadPurpose; qc\_Reports читає tblRowFilters. Поріг ErrorThreshold (cfgSettings, options, реєстр; fn\_ToRate без правила «число > 1 = відсотки»); помилка ПеревищенийПоріг у fn\_EnforceSchema; перевірка порогу в pq\_RegistryCheck. У pq\_QualitySummary нові колонки RowsDropped, FilterRulesSkipped, ErrorRate, ErrorThreshold, QualityFlag. Тестові запити test\_RowFilters, test\_RowFiltersReal, test\_ToRate, q\_Registry. Документ доповнено розділом 6.7 |
| v4.4 | 15.10.2026 | Макрос SelectSourceFiles (modSelectSourceFiles) для tblSourceRegistry: активні рядки, перевірка CSV за схемою з визначенням BOM, підсумок і запис лише після підтвердження, оновлення з підсумком QualityFlag. CleanExportSelectedSheets доповнено від'єднанням таблиць запитів (з modExport); modExport і старий модуль вибору файлів замінено. BasePath відкладено. Макрос читає CsvDelimiter, CsvEncoding, TrueValues з cfgSettings (єдине джерело, без дубльованих констант). Розділ 6.8 про макроси, «Загальне» тепер 6.9 |
| v4.5 | 16.10.2026 | Етап B: виявлення дублікатів заголовків (DuplicateHeaders у звітах, попередження в SelectSourceFiles), поведінка #N/A і відкритих файлів задокументована за тестами. Правило регістру для довідників: fn\_Canon (TargetType, PeriodType у написанні довідників), CustomTransform без урахування регістру, TargetType поза tblTargetType у pq\_RegistryCheck, підказка «регістр відрізняється» в pq\_SchemaDrift. З аудиту закрито D1 (правило регістру), D3 (Use через fn\_ToBool у fn\_ApplyRowFilters), D9 (культура з cfgSettings у fn\_ApplyRowFilters, fn\_EnforceSchema, Calendar). Розділ 6.9 «Правило регістру і довідники», «Загальне» тепер 6.10 |
| v4.6 | 16.10.2026 | Етап C: регресійні тести (тестова книга, SetupTestEngine, test\_Summary, test\_Load, test\_All; розділ 6.11). Виправлено політику Lenient (Table.ReplaceErrorValues; патч C1). Аудит D4–D7: cfgReportTypes (структура звітів), cfgStatuses (статуси), fn\_LoadVerdict (одне правило зупинки для fn\_EnforceSchema і QualityFlag), fn\_EffectiveSettings (Policy, Threshold, Culture); fn\_EnforceSchema приймає запис налаштувань; внутрішнє Label замінено на SourceID; невідома ErrorPolicy тепер дає QualityFlag = FAIL; test\_LoadVerdict |
| v4.7 | 16.10.2026 | Аудит D2 і D10. fn\_RegistryIssues: одна реалізація перевірок реєстру для pq\_RegistryCheck і fn\_ReadPurpose (SourceID унікальний у всьому реєстрі; колонки TargetTypes, Blocking). fn\_LoadConfig і cfgSettings[Tables]: конфігурація одним записом, імена таблиць в одному місці. Нова сигнатура fn\_ReadPurpose(config, targetType, options). Оновлено qc\_Reports, q\_Registry, src\_SchemaRegistry, pq\_GenerateSchemaTemplate, pq\_RowFiltersCheck, test\_Load; test\_RowFiltersReal читає таблиці прямо (Formula.Firewall). test\_RegistryIssues |
| v4.8 | 17.10.2026 | Етап D (патчі F, F1–F3). VBA: спільний модуль modEngineCommon (аудит D13), усі макроси, крім імпорту й SetupTestEngine, переведено на нього; новий modRunLog (RefreshWithRunLog, tblRunLog); експорт для Git v2 (точні залежності, очищення застарілих файлів, таблиці конфігурації в CSV, тип Function); профайлер (ім'я запиту за Location=, підсумок на PQ\_ProfilerSummary, сортування, видалення старих SUMMARY). M: cfgSettings[Extensions] у fn\_ParseSource і fn\_ParseHeaders (аудит D8 у частині розширень), test\_Extensions (test\_All — 8 наборів); Table.Buffer у pq\_ExtractedColumns, pq\_SchemaDrift, test\_Summary і лічба помилок за один перегляд у test\_Load (оновлення 502 → 34 с). Розділи 6.8 (макроси) і 6.12 (продуктивність); позначка «Оновлювати під час оновлення всіх» у розділі 5 |
| v4.9 | 23.10.2026 | Патчі G, G1, H (відкочено), H2, I, J, K, L. fn\_SchemaReport: List.Buffer для перевірок колонок, помилки по колонках лише серед рядків з помилками, null через fn\_ColumnNullCounts (Table.Profile); test\_NullCounts; test\_All — 9 наборів; звіт якості реальних джерел у ~5 разів швидший. fn\_QcReports і fn\_SummaryTable (тіла qc\_Reports і pq\_QualitySummary як функції); test\_Summary рахує лише тестові SourceID і сам читає tblTestExpected, src\_TestExpected видалено. RefreshWithRunLog: замість другого оновлення зведення перевіряє й вмикає «Оновлювати під час оновлення всіх». Структура тек Power Query (розділ 5.1), cfgQueryGroups і поле group у queries.yml. Розділ 6.12 доповнено вимірюваннями на 200 МБ і висновком «обсяг читання не дорівнює часу». Зовнішній аудит концепції: external\_review\_2026-10-20.md |
| v4.10 | 24.10.2026 | Патчі N, N2, N3 за двома аудитами коду (розділ 6.13): правила компілюються з перевіркою назв DataType/CustomTransform; клітинки-помилки не ламають відсікання порожніх рядків і не збігаються з фільтрами; фільтри в культурі cfgSettings; ErrorPolicy/ErrorThreshold/Culture лише з реєстру чи cfgSettings; джерела без періоду не дають дубль періоду; EmptyColumns без відсутніх колонок; BOM у fn\_ParseHeaders; MetaColumns — запис; tblPeriod у cfgSettings[Tables]; один буфер у fn\_EnforceSchema для Strict і порогу; підказки про пробіли в заголовках; ЗбігЗМетаколонкою. Тести: test\_AuditFixes, test\_HeadersConsistency, новий test\_RowFilters (13 випадків), test\_Load з одним завантаженням; test\_RowFiltersReal видалено; test\_All — 11 наборів. Профайлер: ручний режим обчислення, прогрів, без «лише підключень», відновлення фонового режиму. Експорт: типи Function/Config/Source/Workbook/Transform/Standalone |
| v4.11 | 24.10.2026 | modTableCatalog: процедура UpdateHelp (ListAllTables видалено). Аркуш HELP замість «Список\_Таблиць» (перейменовано один раз вручну/макросом, код перейменування прибрано), каталог з колонками Підпис, Довідка, Порушення, таблиця hlpProject з правилами структури книги, підписи над таблицями (назва, посилання на призначення в каталозі, посилання на hlp-довідку), шапки не вище рядка 3. Правила — розділ 6.14; ExportPowerQueriesToRepo створює powerquery\README.md з hlpProject, tblTableCatalog, hlp-таблиць і cfgQueryGroups; прибрано одноразовий код (перейменування аркуша, очищення імен підписів v4.11b–c, видалення старих блоків SUMMARY у профайлері, стартовий опис таблиць «Опис\_»); tblRunLog і pq\_QualitySummary з RefreshWithRunLog ставляться в A3; експорт модулів VBA у powerquery\vba і розділ «Макроси VBA» в README, tblTableCatalog виключено з CSV; шаблон репозиторію (.gitattributes, .gitignore, README, tests\ з фікстурами, docs\) |
