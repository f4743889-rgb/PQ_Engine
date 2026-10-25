// pq_SchemaIssues - таблиця для аркуша: які джерела не дали заголовків і чому
let
    T = Table.SelectRows(pq_GenerateSchemaTemplate, each [SchemaStatus] <> "OK"),
    S = Table.SelectColumns(T,
            {"SourceID", "TargetType", "PeriodType", "PeriodNumber",
             "SourcePath", "SchemaOrigin", "SchemaStatus", "StatusMessage"})
in
    Table.Distinct(S)