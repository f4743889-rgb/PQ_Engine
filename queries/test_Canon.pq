// test_Canon: перевірка правила регістру. Очікування: у всіх рядках Pass = TRUE.
// Використовує реальні tblTargetType і tblPeriodRules: перший рядок кожного довідника.
let
    TT  = Excel.CurrentWorkbook(){[Name = "tblTargetType"]}[Content],
    PR  = Excel.CurrentWorkbook(){[Name = "tblPeriodRules"]}[Content],
    T1  = Text.Trim(Text.From(TT{0}[TargetType])),
    P1  = Text.Trim(Text.From(PR{0}[PeriodType])),
    CT  = fn_Canon("tblTargetType", "TargetType"),
    CP  = fn_Canon("tblPeriodRules", "PeriodType"),

    Cases = {
        {"TargetType у верхньому регістрі",  CT(Text.Upper(T1)),          T1},
        {"TargetType у нижньому з пробілами", CT("  " & Text.Lower(T1) & " "), T1},
        {"PeriodType у верхньому регістрі",  CP(Text.Upper(P1)),          P1},
        {"Значення поза довідником",          CT(" НемаТакого "),          "НемаТакого"},
        {"Порожнє значення",                  CT("  "),                    null},
        {"null",                              CT(null),                    null},
        {"Невідома таблиця-довідник",         fn_Canon("tblНемає", "X")("abc"), "abc"}
    }
in
    Table.FromRows(
        List.Transform(Cases, (c) => {c{0}, c{1}, c{2}, c{1} = c{2}}),
        type table [Case = text, Actual = nullable text, Expected = nullable text, Pass = logical])