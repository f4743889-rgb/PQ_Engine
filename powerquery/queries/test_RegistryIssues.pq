// test_RegistryIssues: модульний тест fn_RegistryIssues на вбудованому реєстрі. Очікування: Pass = TRUE.
// Потребує в книзі довідників tblTargetType (TestCsv, TestXlsx) і tblPeriodRules (Місяць до 12).
let
    Reg = #table(
        {"SourceID", "TargetType", "FullPath", "Path", "FileName", "TableSheetName",
         "PeriodType", "PeriodNumber", "IsRequired", "Enabled", "ErrorPolicy", "ErrorThreshold"},
        {
            {"A", "TestCsv",  null, "D:\x", "f1.csv",  null, "Місяць", 1,  null, null,  null, null},
            {"A", "TestXlsx", null, "D:\x", "f2.xlsx", null, "Місяць", 1,  null, null,  null, null},
            {"B", "TestCsv",  null, "D:\x", "f1.csv",  null, "Місяць", 2,  null, null,  null, null},
            {"C", "testcsv",  null, "D:\x", "f3.csv",  null, "місяць", 1,  null, null,  null, null},
            {"",  "TestXlsx", null, "D:\x", "f5.xlsx", null, "Місяць", 13, null, null,  null, null},
            {"E", "TestXlsx", null, "D:\x", "f6.xlsx", null, "Місяць", 3,  null, null,  null, "abc"},
            {"A", "TestCsv",  null, "D:\x", "f7.csv",  null, "Місяць", 4,  null, false, null, null},
            {"G", "НемаТакого", null, "D:\x", "f8.csv", null, "Місяць", 5,  null, null,  null, null}
        }),
    I  = fn_RegistryIssues(Reg),
    IA = fn_RegistryIssues(Reg, [AllowDuplicatePeriods = true]),

    Find = (t as table, code as text) as nullable record =>
        let r = Table.SelectRows(t, each [Code] = code) in if Table.IsEmpty(r) then null else r{0},
    Show = (r as nullable record) as text =>
        if r = null then "(немає)"
        else r[Key] & " | " & Text.From(r[Count]) & " | " & (if r[Blocking] then "блок" else "звіт")
             & " | " & Text.Combine(List.Sort(r[Targets]), ","),

    Cases = {
        {"SourceID унікальний у всьому реєстрі", I,  "ДублюванняSourceID",  "A | 2 | блок | TestCsv,TestXlsx"},
        {"Той самий файл у призначенні",         I,  "ДублюванняФайлу",     "d:\x\f1.csv| | 2 | блок | TestCsv"},
        {"Той самий період (регістр не важить)", I,  "ДублюванняПеріоду",   "Місяць|1 | 2 | блок | TestCsv"},
        {"Порожній SourceID",                    I,  "ПорожнійSourceID",    " | 1 | звіт | TestXlsx"},
        {"Недопустимий період",                  I,  "НедопустимийПеріод",  "Місяць|13 | 1 | звіт | TestXlsx"},
        {"Некоректний поріг",                    I,  "НекоректнийПоріг",    "E | abc | 1 | звіт | TestXlsx"},
        {"Призначення поза довідником",          I,  "НевідомеПризначення", "НемаТакого | 1 | звіт | НемаТакого"},
        {"AllowDuplicatePeriods знімає блок",    IA, "ДублюванняПеріоду",   "Місяць|1 | 2 | звіт | TestCsv"}
    },
    Rows = List.Transform(Cases, (c) =>
        let a = try Show(Find(c{1}, c{2}))
        in  [Case = c{0}, Expected = c{3},
             Actual = if a[HasError] then "ПОМИЛКА: " & a[Error][Message] else a[Value],
             Pass = not a[HasError] and a[Value] = c{3}]),
    Total = [Case = "Усього проблем (вимкнений рядок ігнорується)", Expected = "7",
             Actual = Text.From(Table.RowCount(I)), Pass = Table.RowCount(I) = 7]
in
    Table.FromRecords(Rows & {Total}, type table [Case = text, Expected = text, Actual = text, Pass = logical])