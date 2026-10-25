// test_FlagVsLoad: перехресна перевірка звіту якості і завантаження на тих самих фікстурах
// (v4.15, рецензія коду 25.10.2026): для кожного сценарію tblTestLoad з одним SourceID
// QualityFlag = FAIL тоді й лише тоді, коли завантаження зупиняється (Outcome = ERROR).
// Обидва рішення приймає fn_LoadVerdict; тест ловить розбіжність у тому, що йому передають.
// Бере фактичні результати test_Summary і test_Load (файлів не відкриває, книгу не читає).
// Джерело, якого немає у зведенні (Enabled = FALSE), має порожній QualityFlag.
// Очікування: у всіх рядках Pass = TRUE.
let
    S = test_Summary,
    L = test_Load,
    FlagOf = (id as text) as text =>
        let r = Table.SelectRows(S, each Text.Trim([SourceID] ?? "") = id and [Field] = "QualityFlag")
        in  if Table.IsEmpty(r) then "" else Text.Upper(Text.Trim(r{0}[Actual] ?? "")),
    Single = Table.SelectRows(L, each not Text.Contains([SourceIDs] ?? "", ",")),
    Rows = List.Transform(Table.ToRecords(Single), (c) =>
        let
            id   = Text.Trim(c[SourceIDs] ?? ""),
            flag = FlagOf(id),
            out  = Text.Upper(Text.Trim(c[Outcome] ?? ""))
        in
            [Case = c[Case], SourceID = id, QualityFlag = flag, Outcome = out,
             Pass = (flag = "FAIL") = (out = "ERROR")])
in
    Table.Sort(
        Table.FromRecords(Rows,
            type table [Case = text, SourceID = text, QualityFlag = text, Outcome = text, Pass = logical]),
        {{"Pass", Order.Ascending}, {"Case", Order.Ascending}})
