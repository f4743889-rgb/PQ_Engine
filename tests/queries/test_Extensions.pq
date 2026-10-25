// test_Extensions: розширення файлів з cfgSettings[Extensions]. Очікування: у всіх рядках Pass = TRUE.
// Файли не відкриває: вміст CSV задано текстом, тому запит не залежить від тестових файлів.
let
    Ex   = cfgSettings[Extensions],
    Csv  = Text.ToBinary("a;b#(lf)1;2", TextEncoding.Utf8),
    Opts = [Delimiter = ";", Encoding = 65001, TableSheetName = ""],
    Src  = (ext as text) => [Path = "test." & ext, Ext = ext, Content = Csv],
    Code = (f as function) => try f() otherwise null,
    ErrCode = (f as function) =>
        let r = try f() in if r[HasError] then r[Error][Reason] else "без помилки",

    Cases = {
        {"Extensions: списки малими літерами",
            List.AllTrue(List.Transform(Ex[Csv] & Ex[Excel], each _ = Text.Lower(Text.Trim(_)))), true},
        {"Extensions: Csv і Excel не перетинаються",
            List.IsEmpty(List.Intersect({Ex[Csv], Ex[Excel]})), true},
        {"ParseSource: csv читається",
            Code(() => Table.ColumnNames(fn_ParseSource(Src("csv"), Opts))), {"a", "b"}},
        {"ParseSource: TXT у верхньому регістрі",
            Code(() => Table.ColumnNames(fn_ParseSource(Src("TXT"), Opts))), {"a", "b"}},
        {"ParseSource: невідоме розширення",
            ErrCode(() => fn_ParseSource(Src("pdf"), Opts)), "НепідтримуванийФормат"},
        {"ParseHeaders: csv",
            Code(() => fn_ParseHeaders(Src("csv"), Opts)), {"a", "b"}},
        {"ParseHeaders: невідоме розширення",
            ErrCode(() => fn_ParseHeaders(Src("pdf"), Opts)), "НепідтримуванийФормат"}
    },
    Show = (v) => if v is list then Text.Combine(List.Transform(v, Text.From), ",") else Text.From(v)
in
    Table.FromRows(
        List.Transform(Cases, (c) => {c{0}, Show(c{1}), Show(c{2}), c{1} = c{2}}),
        type table [Case = text, Actual = nullable text, Expected = nullable text, Pass = logical])
