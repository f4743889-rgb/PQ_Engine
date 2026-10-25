// test_HeadersConsistency: швидке читання заголовків (fn_ParseHeaders) дає ті самі колонки, що й повне
// читання (fn_ParseSource), для кожного тестового джерела, файл якого відкривається (аудит 1.8, BOM).
// Лише SourceID з tblTestExpected (реальні джерела тестової книги не беруться). Запит відкриває файли,
// тому сам читає книгу і викликає лише функції. Очікування: у всіх рядках Pass = TRUE.
let
    Exp  = Excel.CurrentWorkbook(){[Name = "tblTestExpected"]}[Content],
    Ids  = List.Distinct(List.Transform(Table.Column(Exp, "SourceID"), each Text.Trim(Text.From(_ ?? "")))),
    Rows = List.Select(Table.ToRecords(fn_ActiveSources(fn_LoadConfig()[Registry])),
               (r) => List.Contains(Ids, Text.Trim(Text.From(r[SourceID] ?? "")))),
    Check = (r as record) as nullable list =>
        let
            Src  = fn_ResolveSource(r),
            Opt  = fn_Options(r),
            H    = try fn_ParseHeaders(Src, Opt),
            Full = try Table.ColumnNames(fn_ParseSource(Src, Opt))
        in
            // джерела, які не читаються (немає файла, пошкоджений, непідтримуваний), тут не перевіряються
            if Src[ErrorCode] <> null or Full[HasError] then null
            else {Text.From(r[SourceID]),
                  if H[HasError] then "помилка: " & H[Error][Message] else Text.Combine(H[Value], "|"),
                  Text.Combine(Full[Value], "|")},
    Res = List.RemoveNulls(List.Transform(Rows, Check))
in
    Table.FromRows(
        List.Transform(Res, (x) => {x{0}, x{1}, x{2}, x{1} = x{2}}),
        type table [SourceID = text, Headers = nullable text, FullRead = nullable text, Pass = logical])
