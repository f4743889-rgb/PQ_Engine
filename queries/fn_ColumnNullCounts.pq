// fn_ColumnNullCounts: кількість null у кожній колонці таблиці (запис колонка -> число), v4.9
// cols        - колонки, для яких рахувати
// errorCols   - колонки з клітинками-помилками: для них старий спосіб (без рядків з помилкою в
//               цій колонці), бо Table.Profile на помилках не розрахований
// Інші колонки - одним вбудованим Table.Profile (GO3, 81 тис. рядків, 16 колонок: 7,1 с -> 1,1 с).
// Порожній текст "" не є null (як і раніше у fn_SchemaReport).
(tbl as table, cols as list, optional errorCols as nullable list) as record =>
let
    ErrC  = List.Intersect({errorCols ?? {}, cols}),
    Clean = List.Difference(cols, ErrC),
    Prof  = if List.IsEmpty(Clean) or Table.IsEmpty(tbl)
            then #table({"Column", "NullCount"}, {})
            else Table.Profile(Table.SelectColumns(tbl, Clean)),
    // кількість null колонки з таблиці Profile (пошук рядка; Record.FromList тут давав
    // Expression.Error «Указано ключі 1, але надано значення 2», test_NullCounts 24.10.2026)
    ProfRows = Table.ToRecords(Table.SelectColumns(Prof, {"Column", "NullCount"})),
    FromProf = (c as text) as number =>
                   List.Sum({0} & List.Transform(List.Select(ProfRows, each _[Column] = c), each _[NullCount] ?? 0)),
    Slow  = (c as text) as number =>
                List.Count(List.Select(Table.Column(Table.RemoveRowsWithErrors(tbl, {c}), c), each _ = null))
in
    Record.FromList(
        List.Transform(cols, (c) => if List.Contains(ErrC, c) then Slow(c)
                                    else FromProf(c)),
        cols)