// fn_ParseHeaders - Точний та швидкий парсер заголовків
(src as record, opts as record) as list =>
let
    Names =
        if List.Contains(cfgSettings[Extensions][Csv], Text.Lower(src[Ext] ?? "")) then
            let
                // BOM має пріоритет над Encoding, як у Csv.Document (fn_ParseSource); інакше при 1251
                // BOM UTF-8 читався б як «п»ї» у першому заголовку (аудит 1.8, v4.10)
                Head   = Binary.ToList(Binary.Range(src[Content], 0, 3)),
                Enc    = if List.FirstN(Head, 3) = {239, 187, 191} then 65001
                         else if List.FirstN(Head, 2) = {255, 254} then 1200
                         else opts[Encoding],
                Line   = List.First(
                            Lines.FromBinary(src[Content], QuoteStyle.Csv, false, Enc), null),
                Clean  = Text.Replace(Line ?? "", "#(feff)", ""),
                OneRow = Csv.Document(Text.ToBinary(Clean, TextEncoding.Utf8),
                            [Delimiter = opts[Delimiter], Encoding = TextEncoding.Utf8,
                             QuoteStyle = QuoteStyle.Csv])
            in
                // PromoteHeaders дає ті самі імена (дублікати, "Column N"), що й повне читання
                Table.ColumnNames(Table.PromoteHeaders(OneRow, [PromoteAllScalars = true]))
        else
            Table.ColumnNames(fn_ParseSource(src, opts))
in
    if List.IsEmpty(Names)
    then error Error.Record("ПорожніЗаголовки", "Не вдалося визначити заголовки.", src[Path])
    else Names