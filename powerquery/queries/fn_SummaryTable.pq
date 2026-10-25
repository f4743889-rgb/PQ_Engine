// fn_SummaryTable: таблиця зведення якості зі списку звітів fn_QcReports
// (структура: cfgReportTypes[Summary]; порядок рядків: cfgStatuses[...][Order], проблеми вгорі).
// Використовують pq_QualitySummary і test_Summary.
// Буфер (List.Buffer / Table.Buffer) тут не допомагає: патч M у v4.9 дав 70 с без змін і відкочений
// (architecture, розділ 6.12) - не повертати без нового заміру.
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