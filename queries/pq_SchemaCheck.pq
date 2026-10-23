let
    S    = src_SchemaRegistry,
    W    = Table.AddColumn(S, "EffTarget", each [TargetColumnName] ?? [SourceColumnName], type text),
    Used = Table.SelectRows(W, each [Use]),

    Mk = (rows as table, sev as text, issue as text, col as function, det as function) as table =>
        #table(
            type table [Severity = text, Issue = text, TargetType = nullable text,
                        Column = nullable text, Detail = nullable text],
            List.Transform(Table.ToRecords(rows),
                (r) => {sev, issue, r[TargetType], col(r), det(r)})),

    Cnt = (tbl as table, keys as list) as table =>
        Table.SelectRows(
            Table.Group(tbl, keys, {{"N", each Table.RowCount(_), Int64.Type}}),
            each [N] > 1),

    Names = (t as nullable text) as list =>
        if t = null then {}
        else List.Select(List.Transform(Text.Split(t, ";"), Text.Trim), each _ <> ""),
    UnknownTr = (t as nullable text) as list =>
        List.Select(Names(t), (n) => not List.Contains(
            List.Transform(Record.FieldNames(cfgTransforms), Text.Lower), Text.Lower(n))),

    c1 = Mk(Cnt(W, {"TargetType", "SourceColumnName"}), "ERROR", "Дубль SourceColumnName",
            (r) => r[SourceColumnName], (r) => "Рядків: " & Text.From(r[N])),
    c2 = Mk(Cnt(Used, {"TargetType", "EffTarget"}), "ERROR", "Дубль цільової назви",
            (r) => r[EffTarget], (r) => "Рядків: " & Text.From(r[N])),
    c3 = Mk(Table.SelectRows(W, each [DataType] <> null
                                 and not Record.HasFields(cfgTypeMap, fn_TypeKey([DataType]))),
            "ERROR", "Невідомий DataType", (r) => r[SourceColumnName], (r) => r[DataType]),
    c4 = Mk(Table.SelectRows(W, each [DataType] = null),
            "WARN", "DataType порожній (буде Text)", (r) => r[SourceColumnName], (r) => null),
    c5 = Mk(Table.SelectRows(W, each not List.IsEmpty(UnknownTr([CustomTransform]))),
            "ERROR", "Невідоме CustomTransform", (r) => r[SourceColumnName],
            (r) => Text.Combine(UnknownTr(r[CustomTransform]), ", ")),
    c6 = Mk(Table.SelectRows(Used, each List.Contains(cfgSettings[MetaColumns], [EffTarget])),
            "ERROR", "Збіг із колонкою метаданих", (r) => r[EffTarget], (r) => null),
    c7 = Mk(Table.SelectRows(W, each [IsRequired] and not [Use]),
            "WARN", "IsRequired при Use = Ні", (r) => r[SourceColumnName], (r) => null),
    c8 = Mk(Table.SelectRows(
                Table.Group(W, {"TargetType"},
                    {{"U", each List.Count(List.Select(Table.Column(_, "Use"), (x) => x = true)),
                      Int64.Type}}),
                each [U] = 0),
            "ERROR", "Немає жодної використовуваної колонки", (r) => null, (r) => null)
in
    Table.Combine({c1, c2, c3, c4, c5, c6, c7, c8})