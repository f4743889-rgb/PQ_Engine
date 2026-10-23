// fn_SummaryTable: таблиця зведення якості зі списку звітів fn_QcReports
// (структура: cfgReportTypes[Summary]; порядок рядків: cfgStatuses[...][Order], проблеми вгорі).
// Використовують pq_QualitySummary і test_Summary.
(reports as list) as table =>
let
    Typ  = cfgReportTypes[Summary],
    // List.Buffer і Table.Buffer: звіти (лінивий список) обчислюються один раз, навіть якщо
    // результат переглядають кілька разів (v4.9, патч M: діагностика ~25 с проти 67,7 с у профайлері)
    Recs = List.Buffer(List.Transform(reports, each _[Summary])),
    T    = if List.IsEmpty(Recs) then #table(Typ, {})
           else Table.FromRecords(Recs, Typ, MissingField.UseNull),
    P    = Table.AddColumn(T, "__Order",
               each Record.FieldOrDefault(cfgStatuses, [Status] ?? "", [Order = 99])[Order], Int64.Type),
    S    = Table.Sort(P, {{"__Order", Order.Ascending}, {"SourceID", Order.Ascending}})
in
    Table.Buffer(Table.RemoveColumns(S, {"__Order"}))