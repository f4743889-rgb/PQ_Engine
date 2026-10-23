// pq_QualityColumns: звіт по колонках схеми (структура: cfgReportTypes[Columns])
let
    Typ = cfgReportTypes[Columns],
    T   = Table.Combine({#table(Typ, {})} & List.Transform(qc_Reports, each _[Columns]))
in
    Value.ReplaceType(Table.SelectColumns(T, Table.ColumnNames(#table(Typ, {})), MissingField.UseNull), Typ)