// cfgSettings
[
    CsvDelimiter = ";",
    CsvEncoding  = 1251,
    Culture      = "uk-UA",
    TrueValues   = {"true", "1", "так", "yes", "y", "+"},
    ErrorPolicy = "Report",   // Strict | Report | Lenient
    ErrorThreshold = null,    // поріг частки рядків з помилками для Report/Lenient: 0.05 або "5%"; null = вимкнено
    MetaColumns = {"SourceID", "PeriodType", "PeriodNumber"},

    // Іменовані клітинки Excel
    NameTargetType = "targetReportCell",
    NameMVO = "МВО",
    NameReportDate = "ДатаЗвіту",

    // Таблиці книги (fn_LoadConfig, fn_Canon, fn_RegistryIssues, fn_TemplateColumns)
    Tables = [Registry = "tblSourceRegistry", Schema = "tblSchemaRegistry", RowFilters = "tblRowFilters",
              DefaultHeaders = "tblDefaultHeaders", TargetTypes = "tblTargetType", PeriodRules = "tblPeriodRules",
              Period = "tblPeriod"],

    // Розширення файлів-джерел (малими літерами, без крапки): fn_ParseSource, fn_ParseHeaders,
    // макрос вибору файлів (фільтр діалогу і перевірка CSV)
    Extensions = [Csv = {"csv", "txt"}, Excel = {"xlsx", "xlsm", "xls"}]
]