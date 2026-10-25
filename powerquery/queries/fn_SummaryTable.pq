// fn_SummaryTable: таблиця зведення якості зі списку звітів fn_QcReports
// (структура: cfgReportTypes[Summary]; порядок рядків: cfgStatuses[...][Order], проблеми вгорі).
// Використовують pq_QualitySummary і test_Summary.
(reports as list) as table =>
let
    Typ  = cfgReportTypes[Summary],
    Recs = List.Transform(reports, each _[Summary]),
    T    = if List.IsEmpty(Recs) then #table(Typ, {})
           else Table.FromRecords(Recs, Typ, MissingField.UseNull),
    P    = Table.AddColumn(T, "__Order",
               each Record.FieldOrDefault(cfgStatuses, [Status] ?? "", [Order = 99])[Order], Int64.Type),
    S    = Table.Sort(P, {{"__Order", Order.Ascending}, {"SourceID", Order.Ascending}})
in
    Table.RemoveColumns(S, {"__Order"})