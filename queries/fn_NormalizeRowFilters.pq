// fn_NormalizeRowFilters: нормалізація tblRowFilters (як fn_NormalizeSchema для схеми)
// Таблицю читає запит-викликач, що сам відкриває книгу (qc_Reports, запити виклику fn_ReadPurpose):
//   fn_LoadConfig()[RowFilters] (або fn_NormalizeRowFilters(таблиця tblRowFilters))
// Рядки без TargetType, SourceColumnName або Operator відкидаються (їх показує pq_RowFiltersCheck).
// Use: порожнє = TRUE, решта через fn_ToBool.
(T0 as table) as table =>
let
    T1 = Table.SelectColumns(T0, {"TargetType", "SourceColumnName", "Operator", "Value", "Use"},
             MissingField.UseNull),

    Norm = (v as any) as nullable text =>
               let t = Text.Trim(Text.From(v ?? "")) in if t = "" then null else t,

    T2 = Table.TransformColumns(T1, {
             {"TargetType",       fn_Canon(cfgSettings[Tables][TargetTypes], "TargetType"), type nullable text},
             {"SourceColumnName", Norm, type nullable text},
             {"Operator",         Norm, type nullable text},
             {"Value",            each if _ = null then null else Text.From(_), type nullable text}}),

    T3 = Table.SelectRows(T2, each [TargetType] <> null and [SourceColumnName] <> null
                                   and [Operator] <> null),

    T4 = Table.TransformColumns(T3, {{"Use", each fn_ToBool(_, true), type logical}})
in
    T4