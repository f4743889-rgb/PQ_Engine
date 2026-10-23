// pq_ExtractedColumns - унікальні колонки на призначення
// Table.Buffer: Src використовується кілька разів (L, Total); без буфера кожне звернення
// заново відкривало б усі файли-джерела (патч F2)
let
    Src = Table.Buffer(pq_GenerateSchemaTemplate),
    L   = Table.AddIndexColumn(
              Table.SelectRows(Src, each [SourceColumnName] <> null),
              "Idx", 0, 1, Int64.Type),

    // скільки реальних файлів на призначення
    Total = Table.Group(
                Table.Distinct(Table.SelectColumns(
                    Table.SelectRows(Src, each [SchemaOrigin] = "File"),
                    {"TargetType", "SourceID"})),
                {"TargetType"},
                {{"FilesTotal", each Table.RowCount(_), Int64.Type}}),

    Cols = Table.Group(L, {"TargetType", "SourceColumnName"}, {
                {"FirstIdx", each List.Min([Idx]), Int64.Type},
                {"FilesWithColumn",
                    (g) => List.Count(List.Distinct(
                        Table.Column(Table.SelectRows(g, (r) => r[SchemaOrigin] = "File"), "SourceID"))),
                    Int64.Type},
                {"InTemplate", (g) => List.Contains(g[SchemaOrigin], "Template"), type logical}}),

    Joined = Table.NestedJoin(Cols, {"TargetType"}, Total, {"TargetType"}, "T", JoinKind.LeftOuter),
    Exp    = Table.ExpandTableColumn(Joined, "T", {"FilesTotal"}),
    Fixed  = Table.ReplaceValue(Exp, null, 0, Replacer.ReplaceValue, {"FilesTotal"}),
    Typed  = Table.TransformColumnTypes(Fixed, {{"FilesTotal", Int64.Type}}),

    // блоки по призначенню: призначення за першою появою, колонки за FirstIdx
    Sorted = Table.Sort(Typed, {{"FirstIdx", Order.Ascending}}),
    Types  = List.Distinct(Sorted[TargetType]),
    Ranked = Table.AddColumn(Sorted, "TypeNo",
                 each List.PositionOf(Types, [TargetType]), Int64.Type),
    Blocks = Table.Sort(Ranked, {{"TypeNo", Order.Ascending}, {"FirstIdx", Order.Ascending}})
in
    Table.RemoveColumns(Blocks, {"TypeNo"})