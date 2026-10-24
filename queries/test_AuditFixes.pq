// test_AuditFixes: регресійні тести вад, знайдених аудитом коду (v4.10). Очікування: у всіх рядках Pass = TRUE.
// Дані задано в коді (файлів не відкриває); довідник tblTargetType має містити TestCsv.
let
    Cfg = cfgSettings,
    // схема у форматі fn_NormalizeSchema
    Schema = (rows as list) as table =>
        #table(type table [TargetType = text, SourceColumnName = text, TargetColumnName = nullable text,
                           DataType = nullable text, IsRequired = logical, TrimFlag = logical,
                           CustomTransform = nullable text, Use = logical], rows),
    Reason = (f as function) as text => let r = try f() in if r[HasError] then r[Error][Reason] else "без помилки",

    Ok      = Schema({{"TestCsv", "a", null, "text", true, false, null, true},
                      {"TestCsv", "b", null, "number", false, false, null, true}}),
    BadType = Schema({{"TestCsv", "a", null, "НемаТакогоТипу", true, false, null, true}}),
    BadTr   = Schema({{"TestCsv", "a", null, "text", true, false, "НемаТакоїДії", true}}),
    Meta    = Schema({{"TestCsv", "a", Cfg[MetaColumns]{0}, "text", true, false, null, true}}),
    RulesOk = fn_GetSchemaRules(Ok, "TestCsv"),

    // 1.2: рядок, у якому перед клітинкою-помилкою лише порожні клітинки
    RawErr  = Table.TransformColumnTypes(#table({"a", "b"}, {{null, "zz"}, {"x", "1"}}), {{"b", type number}}, "en-US"),
    // 1.3: справжні числа в колонці з правилом NotNumeric
    RawNum  = #table({"a", "b"}, {{"x", 10.5}, {"y", "абв"}, {"z", null}}),
    Flt     = #table(type table [TargetType = text, SourceColumnName = text, Operator = text, Value = nullable text, Use = logical],
                     {{"TestCsv", "b", "NotNumeric", null, true}}),
    // 1.5: два джерела одного призначення без періоду
    Reg     = #table({"SourceID", "TargetType", "Path", "FileName", "PeriodType", "PeriodNumber", "Enabled"},
                     {{"A1", "TestCsv", "D:\x", "a.csv", null, null, true},
                      {"A2", "TestCsv", "D:\x", "b.csv", null, null, true}}),
    Codes   = Table.Column(fn_RegistryIssues(Reg), "Code"),
    // 1.6: необов'язкова колонка b відсутня у файлі
    Rep     = fn_SchemaReport(#table({"a"}, {{"1"}, {"2"}}), RulesOk, "T1", "uk-UA")[Summary],

    Cases = {
        {"1.1 невідомий DataType -> помилка правил",      Reason(() => fn_GetSchemaRules(BadType, "TestCsv")), "НевідомийТипДаних"},
        {"1.1 невідоме CustomTransform -> помилка правил", Reason(() => fn_GetSchemaRules(BadTr, "TestCsv")),   "НевідомеПеретворення"},
        {"3 збіг з колонкою метаданих",                    Reason(() => fn_GetSchemaRules(Meta, "TestCsv")),    "ЗбігЗМетаколонкою"},
        {"1.1 коректна схема компілюється",                 Reason(() => fn_GetSchemaRules(Ok, "TestCsv")),      "без помилки"},
        {"1.2 клітинка-помилка після порожніх: рядків",
            Text.From(Table.RowCount(fn_ApplyRulesEx(RawErr, RulesOk, "uk-UA")[Data])), "2"},
        {"1.3 NotNumeric не відкидає справжнє число",
            Text.From(fn_ApplyRowFilters(RawNum, Flt, "TestCsv", "uk-UA")[Dropped]), "1"},
        {"1.5 джерела без періоду - не дубль періоду",
            Text.From(List.Contains(Codes, "ДублюванняПеріоду")), "false"},
        {"1.6 відсутня необов'язкова - не в EmptyColumns",  Rep[EmptyColumns] ?? "", ""},
        {"1.6 відсутня необов'язкова - у MissingOptional",  Rep[MissingOptional] ?? "", "b"},
        {"1.4 options не змінюють ErrorPolicy",
            fn_EffectiveSettings(null, [ErrorPolicy = "Lenient"])[Policy],
            Text.Proper(Text.From(Cfg[ErrorPolicy] ?? "Report"))},
        {"3 Culture з колонки реєстру",                     fn_EffectiveSettings([Culture = "en-US"])[Culture], "en-US"},
        {"заголовок з пробілом: підказка в повідомленні",
            Text.From(Text.Contains(
                let r = try Table.RowCount(fn_EnforceSchema(#table({"a "}, {{"1"}}), RulesOk))
                in  if r[HasError] then r[Error][Message] else "",
                "пробілами на краях")), "true"}
    }
in
    Table.FromRows(
        List.Transform(Cases, (c) => {c{0}, c{1}, c{2}, c{1} = c{2}}),
        type table [Case = text, Actual = nullable text, Expected = nullable text, Pass = logical])