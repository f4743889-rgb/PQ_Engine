// fn_QcReports: звіти якості (fn_RowReport) для активних джерел реєстру
// config    - запис fn_LoadConfig() (Registry, Schema, RowFilters)
// sourceIds - необов'язковий список SourceID: звіти лише для них (тести); null = усі активні джерела.
//             Фільтр діє після fn_ActiveSources, тому вимкнені джерела лишаються відсутніми.
// Викликається із запиту, що відкриває файли (pq_QualitySummary, pq_QualityColumns, pq_QualityErrors, test_Summary): той сам читає книгу
// через fn_LoadConfig (правило трьох видів запитів).
(config as record, optional sourceIds as nullable list) as list =>
let
    Schema = config[Schema],
    Flt    = config[RowFilters],
    All    = fn_ActiveSources(config[Registry]),
    Ids    = if sourceIds = null then null
             else List.Buffer(List.Transform(sourceIds, each Text.Trim(Text.From(_ ?? "")))),
    Active = if Ids = null then All
             else Table.SelectRows(All, each List.Contains(Ids,
                      Text.Trim(Text.From(Record.FieldOrDefault(_, "SourceID", null) ?? "")))),

    TT    = (r as record) as text => Text.Trim(Text.From(Record.FieldOrDefault(r, "TargetType", null) ?? "")),
    Rows  = Table.ToRecords(Active),
    Types = List.Distinct(List.Select(List.Transform(Rows, TT), each _ <> "")),

    // правила компілюються один раз на призначення; помилка зберігається як значення
    RulesMap = Record.FromList(List.Transform(Types, (t) => try fn_GetSchemaRules(Schema, t, Flt)), Types),

    NoType = [HasError = true, Error = [Reason = "НемаПризначення", Message = "TargetType не заданий.", Detail = null]]
in
    List.Transform(Rows, (r) => fn_RowReport(r, Record.FieldOrDefault(RulesMap, TT(r), NoType)))