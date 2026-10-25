// fn_Options - Формування опцій обробки для парсерів
(r as record) as record =>
[
    Delimiter      = Text.From(fn_Opt(r, "Delimiter", cfgSettings[CsvDelimiter])),
    Encoding       = Number.From(fn_Opt(r, "Encoding", cfgSettings[CsvEncoding])),
    TableSheetName = Text.From(fn_Opt(r, "TableSheetName", ""))
]