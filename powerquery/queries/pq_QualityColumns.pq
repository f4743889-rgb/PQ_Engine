// pq_QualityColumns: звіт по колонках схеми (структура: cfgReportTypes[Columns]).
// Звіти будуються тут же через fn_QcReports(fn_LoadConfig()), як у pq_QualitySummary (v4.12).
let
    Typ = cfgReportTypes[Columns],
    T   = Table.Combine({#table(Typ, {})} & List.Transform(fn_QcReports(fn_LoadConfig()), each _[Columns]))
in
    Value.ReplaceType(Table.SelectColumns(T, Table.ColumnNames(#table(Typ, {})), MissingField.UseNull), Typ)