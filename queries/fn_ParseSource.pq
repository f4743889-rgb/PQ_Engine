// fn_ParseSource - чистий парсер повних даних
(src as record, opts as record) as table =>
// Формат визначає розширення з cfgSettings[Extensions]
let
    Name = opts[TableSheetName],
    Ext  = Text.Lower(src[Ext] ?? ""),
    Ex   = cfgSettings[Extensions]
in
    if List.Contains(Ex[Csv], Ext) then
        Table.PromoteHeaders(
            Csv.Document(src[Content],
                [Delimiter = opts[Delimiter], Encoding = opts[Encoding], QuoteStyle = QuoteStyle.Csv]),
            [PromoteAllScalars = true])
    else if List.Contains(Ex[Excel], Ext) then
        let
            // useHeaders = true: листи отримують шапку з першого рядка, таблиці мають власну
            WB     = Excel.Workbook(src[Content], true, true),
            Tables = Table.SelectRows(WB, each [Kind] = "Table"),
            Sheets = Table.SelectRows(WB, each [Kind] = "Sheet"),
            ByName = (t as table) as table =>
                        Table.SelectRows(t, each Text.Lower([Item]) = Text.Lower(Name)),
            T      = ByName(Tables),
            S      = ByName(Sheets)
        in
            if Name <> "" then
                (if not Table.IsEmpty(T) then T{0}[Data]
                else if not Table.IsEmpty(S) then S{0}[Data]
                else error Error.Record("ЛистАбоТаблицюНеЗнайдено",
                        "Не знайдено '" & Name & "' у файлі.", src[Path]))
            else if Table.RowCount(Tables) = 1 then Tables{0}[Data]
            else if Table.IsEmpty(Tables) and Table.RowCount(Sheets) = 1 then Sheets{0}[Data]
            else error Error.Record("НеоднозначнийВибір",
                    "Вкажіть TableSheetName: у файлі кілька таблиць або листів.", src[Path])
    else
        error Error.Record("НепідтримуванийФормат",
            "Формат '" & Ext & "' не підтримується. Дозволені (cfgSettings[Extensions]): "
                & Text.Combine(Ex[Csv] & Ex[Excel], ", ") & ".", src[Path])