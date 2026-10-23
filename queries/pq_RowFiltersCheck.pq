// pq_RowFiltersCheck: перевірка таблиці tblRowFilters (файлів не відкриває)
// Порожній результат = правила коректні. Row = номер рядка даних таблиці (без шапки).
// Перелік операторів має збігатися з tblOperator і з fn_ApplyRowFilters.
let
    Cfg    = fn_LoadConfig(),
    Raw    = Cfg[RowFiltersRaw],
    Schema = Cfg[Schema],

    Ops        = {"equals", "startswith", "contains", "isblank", "notnumeric"},
    NeedsValue = {"equals", "startswith", "contains"},
    NoValue    = {"isblank", "notnumeric"},

    SchemaTypes = List.Distinct(List.Transform(
                      Table.Column(Table.SelectRows(Schema, each [Use]), "TargetType"), Text.Lower)),

    S  = (v as any) as text => try Text.Trim(Text.From(v ?? "")) otherwise "",

    T0 = Table.SelectColumns(Raw, {"TargetType", "SourceColumnName", "Operator", "Value", "Use"},
             MissingField.UseNull),
    T1 = Table.AddIndexColumn(T0, "Row", 1, 1, Int64.Type),

    // повністю порожні рядки таблиці не перевіряємо
    Rows = List.Select(Table.ToRecords(T1),
               (r) => S(r[TargetType]) & S(r[SourceColumnName]) & S(r[Operator]) & S(r[Value]) <> ""),

    Issue = (r as record, sev as text, issue as text, detail as nullable text) as record =>
        [Severity = sev, Issue = issue, Row = r[Row],
         TargetType = S(r[TargetType]), Column = S(r[SourceColumnName]), Detail = detail],

    Check = (r as record) as list =>
        let
            tt   = S(r[TargetType]),
            col  = S(r[SourceColumnName]),
            op   = Text.Lower(S(r[Operator])),
            vals = List.Select(List.Transform(Text.Split(S(r[Value]), "|"), Text.Trim), each _ <> "")
        in
            List.Select({
                if tt = "" or col = "" or op = ""
                    then Issue(r, "ERROR", "Порожнє TargetType, SourceColumnName або Operator",
                               "Правило не застосовується") else null,
                if op <> "" and not List.Contains(Ops, op)
                    then Issue(r, "ERROR", "Невідомий Operator", S(r[Operator])) else null,
                if List.Contains(NeedsValue, op) and List.IsEmpty(vals)
                    then Issue(r, "ERROR", "Value порожнє для " & S(r[Operator]),
                               "Правило не відсіче жодного рядка") else null,
                if List.Contains(NoValue, op) and S(r[Value]) <> ""
                    then Issue(r, "WARN", "Value ігнорується для " & S(r[Operator]), S(r[Value])) else null,
                if tt <> "" and not List.Contains(SchemaTypes, Text.Lower(tt))
                    then Issue(r, "WARN", "TargetType немає в tblSchemaRegistry", tt) else null
            }, each _ <> null),

    All = List.Combine(List.Transform(Rows, Check)),

    Typ = type table [Severity = nullable text, Issue = nullable text, Row = nullable Int64.Type,
                      TargetType = nullable text, Column = nullable text, Detail = nullable text]
in
    if List.IsEmpty(All) then #table(Typ, {})
    else Table.FromRecords(All, Typ, MissingField.UseNull)