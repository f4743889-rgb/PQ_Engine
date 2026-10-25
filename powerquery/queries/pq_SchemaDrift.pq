// pq_SchemaDrift - схема проти реальних файлів
// Table.Buffer: U і Exp0 використовуються по кілька разів (фільтри, з'єднання, ToRecords);
// без буфера кожне звернення до U заново відкривало всі файли-джерела (патч F2)
let
    U    = Table.Buffer(pq_ExtractedColumns),
    Exp0 = Table.Buffer(src_SchemaRegistry),

    Checked = List.Distinct(Table.SelectRows(U, each [FilesTotal] > 0)[TargetType]),
    Found   = Table.SelectRows(U, each [FilesWithColumn] > 0),

    // кількість файлів на призначення, щоб показати її і для рядків «є лише в схемі»
    Totals    = Table.SelectRows(
                    Table.Distinct(Table.SelectColumns(U, {"TargetType", "FilesTotal"})),
                    each [TargetType] <> null),
    TotalsMap = Record.FromList(Table.Column(Totals, "FilesTotal"),
                                Table.Column(Totals, "TargetType")),

    Exp = Table.RenameColumns(
              Table.SelectRows(Exp0, each List.Contains(Checked, [TargetType])),
              {{"TargetType","ETarget"}, {"SourceColumnName","EName"}, {"IsRequired","ERequired"}, {"Use","EUse"}}),

    J = Table.NestedJoin(Found, {"TargetType", "SourceColumnName"},
                         Exp,   {"ETarget", "EName"}, "E", JoinKind.FullOuter),
    X = Table.ExpandTableColumn(J, "E", {"ETarget", "EName", "ERequired", "EUse"}),

    // для підказки «регістр відрізняється»: назви колонок порівнюються точно (як у файлах),
    // тож збіг лише без урахування регістру - майже напевно помилка в схемі
    ExpKeys   = List.Buffer(List.Transform(Table.ToRecords(Exp),
                    each Text.Lower(([ETarget] ?? "") & "|" & ([EName] ?? "")))),
    FoundKeys = List.Buffer(List.Transform(Table.ToRecords(Found),
                    each Text.Lower(([TargetType] ?? "") & "|" & ([SourceColumnName] ?? "")))),
    // для підказки «пробіли на краях»: назви збігаються лише після Text.Trim (v4.10; заголовки
    // файлів навмисно не обрізаються, щоб зміну у вивантаженні було видно)
    KeyT      = (t as nullable text, n as nullable text) as text => Text.Lower((t ?? "") & "|" & Text.Trim(n ?? "")),
    ExpKeysT  = List.Buffer(List.Transform(Table.ToRecords(Exp), each KeyT([ETarget], [EName]))),
    FoundKeysT = List.Buffer(List.Transform(Table.ToRecords(Found), each KeyT([TargetType], [SourceColumnName]))),

    Status = Table.AddColumn(X, "DriftStatus", each
        let
            Used = [EUse] ?? true,
            InFile   = [SourceColumnName] <> null,
            InSchema = [EName] <> null,
            Req      = fn_ToBool([ERequired], false)
        in
            if InSchema and not Used then "Ігнорується (Use = Ні)"
            else if InFile and not InSchema then
                 (if List.Contains(ExpKeys, Text.Lower(([TargetType] ?? "") & "|" & [SourceColumnName]))
                  then "Немає у схемі: регістр відрізняється"
                  else if List.Contains(ExpKeysT, KeyT([TargetType], [SourceColumnName]))
                  then "Немає у схемі: пробіли на краях назви" else "Немає у схемі")
            else if InSchema and not InFile then
                 (if List.Contains(FoundKeys, Text.Lower(([ETarget] ?? "") & "|" & [EName]))
                  then (if Req then "ВІДСУТНЯ обов'язкова: регістр відрізняється"
                        else "Відсутня необов'язкова: регістр відрізняється")
                  else if List.Contains(FoundKeysT, KeyT([ETarget], [EName]))
                  then (if Req then "ВІДСУТНЯ обов'язкова: у файлі назва з пробілами на краях"
                        else "Відсутня необов'язкова: у файлі назва з пробілами на краях")
                  else if Req then "ВІДСУТНЯ обов'язкова" else "Відсутня необов'язкова")
            else if [FilesWithColumn] < [FilesTotal]
                 then (if Req then "ЧАСТКОВО (обов'язкова)" else "Частково")
            else "OK", type text),

    Named = Table.AddColumn(
                Table.AddColumn(
                    Table.AddColumn(Status, "Target", each [TargetType] ?? [ETarget], type text),
                    "Column", each [SourceColumnName] ?? [EName], type text),
                "Total", each Record.FieldOrDefault(TotalsMap, [Target], 0), Int64.Type),

    Slim   = Table.SelectColumns(Named,
                {"Target", "Column", "FilesWithColumn", "Total", "DriftStatus"}),
    Result = Table.ReplaceValue(
                Table.RenameColumns(Slim, {{"Total", "FilesTotal"}}),
                null, 0, Replacer.ReplaceValue, {"FilesWithColumn"}),

    // призначення зі схеми, для яких немає жодного реального файла
    Unchecked = List.Select(List.Distinct(Table.Column(Exp0, "TargetType")),
                    each _ <> null and not List.Contains(Checked, _)),
    Skipped = #table(
        type table [Target = text, Column = text, FilesWithColumn = Int64.Type,
                    FilesTotal = Int64.Type, DriftStatus = text],
        List.Transform(Unchecked, (t) => {t, null, null, 0, "Не перевірено (немає файлів)"}))
in
    Table.TransformColumnTypes(
        Table.Combine({Result, Skipped}),
        {{"FilesWithColumn", Int64.Type}, {"FilesTotal", Int64.Type}})