// pq_QualityErrors: клітинки з помилками (структура: cfgReportTypes[Errors]; статуси описано в cfgStatuses).
// Звіти будуються тут же через fn_QcReports(fn_LoadConfig()), як у pq_QualitySummary (v4.12).
let
    Typ = cfgReportTypes[Errors],
    T   = Table.Combine({#table(Typ, {})} & List.Transform(fn_QcReports(fn_LoadConfig()), each _[Errors]))
in
    Value.ReplaceType(Table.SelectColumns(T, Table.ColumnNames(#table(Typ, {})), MissingField.UseNull), Typ)