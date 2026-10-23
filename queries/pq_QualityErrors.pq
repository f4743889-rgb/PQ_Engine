// pq_QualityErrors: клітинки з помилками (структура: cfgReportTypes[Errors]; статуси описано в cfgStatuses)
let
    Typ = cfgReportTypes[Errors],
    T   = Table.Combine({#table(Typ, {})} & List.Transform(qc_Reports, each _[Errors]))
in
    Value.ReplaceType(Table.SelectColumns(T, Table.ColumnNames(#table(Typ, {})), MissingField.UseNull), Typ)