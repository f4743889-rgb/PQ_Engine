// fn_ParseSource - чистий парсер повних даних
// Заголовки з першого рядка підвищує сам (CSV і аркуші Excel однаково) і зберігає сирі значення
// першого рядка в метаданих таблиці: Value.Metadata(t)[RawHeaders] (v4.15). За ними fn_SchemaReport
// точно визначає повтори назв, а не вгадує їх за суфіксом "_1" (рецензія коду 25.10.2026, CR-01).
// Таблиці Excel (Kind = "Table") мають власні унікальні заголовки, метаданих RawHeaders у них немає.
// RawHeaders обчислюється лише при зверненні (поле запису метаданих ліниве).
(src as record, opts as record) as table =>
// Формат визначає розширення з cfgSettings[Extensions]
let
    Name = opts[TableSheetName],
    Ext  = Text.Lower(src[Ext] ?? ""),
    Ex   = cfgSettings[Extensions],

    // перший рядок -> заголовки; сирі значення першого рядка -> метадані
    Promote = (t0 as table) as table =>
        Value.ReplaceMetadata(
            Table.PromoteHeaders(t0, [PromoteAllScalars = true]),
            [RawHeaders = try List.Transform(Record.FieldValues(Table.First(t0) ?? []),
                                  each if _ = null then "" else Text.From(_))
                          otherwise {}])
in
    if List.Contains(Ex[Csv], Ext) then
        Promote(
            Csv.Document(src[Content],
                [Delimiter = opts[Delimiter], Encoding = opts[Encoding], QuoteStyle = QuoteStyle.Csv]))
    else if List.Contains(Ex[Excel], Ext) then
        let
            // useHeaders = false: шапку аркуша підвищує Promote (як для CSV); таблиці мають власну
            WB     = Excel.Workbook(src[Content], false, true),
            Tables = Table.SelectRows(WB, each [Kind] = "Table"),
            Sheets = Table.SelectRows(WB, each [Kind] = "Sheet"),
            ByName = (t as table) as table =>
                        Table.SelectRows(t, each Text.Lower([Item]) = Text.Lower(Name)),
            T      = ByName(Tables),
            S      = ByName(Sheets)
        in
            if Name <> "" then
                (if not Table.IsEmpty(T) then T{0}[Data]
                else if not Table.IsEmpty(S) then Promote(S{0}[Data])
                else error Error.Record("ЛистАбоТаблицюНеЗнайдено",
                        "Не знайдено '" & Name & "' у файлі.", src[Path]))
            else if Table.RowCount(Tables) = 1 then Tables{0}[Data]
            else if Table.IsEmpty(Tables) and Table.RowCount(Sheets) = 1 then Promote(Sheets{0}[Data])
            else error Error.Record("НеоднозначнийВибір",
                    "Вкажіть TableSheetName: у файлі кілька таблиць або листів.", src[Path])
    else
        error Error.Record("НепідтримуванийФормат",
            "Формат '" & Ext & "' не підтримується. Дозволені (cfgSettings[Extensions]): "
                & Text.Combine(Ex[Csv] & Ex[Excel], ", ") & ".", src[Path])