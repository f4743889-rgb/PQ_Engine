// test_NullCounts: fn_ColumnNullCounts проти старого способу підрахунку null. Очікування: Pass = TRUE.
// Таблиця з null, порожнім текстом і клітинкою-помилкою (b у третьому рядку).
let
    T0 = #table({"a", "b", "c"}, {{"1", null, "x"}, {null, null, ""}, {"3", "zz", null}, {"4", "5", "y"}}),
    T  = Table.TransformColumnTypes(T0, {{"a", type number}, {"b", type number}}, "en-US"),
    Old = (tbl as table, c as text) as number =>
              List.Count(List.Select(Table.Column(Table.RemoveRowsWithErrors(tbl, {c}), c), each _ = null)),
    R1 = fn_ColumnNullCounts(T, {"a", "b", "c"}, {"b"}),
    R2 = fn_ColumnNullCounts(Table.SelectColumns(T, {"a", "c"}), {"a", "c"}, null),
    R3 = fn_ColumnNullCounts(Table.FirstN(T, 0), {"a", "b", "c"}, {}),
    Cases = {
        {"a: число з null",                  R1[a], Old(T, "a")},
        {"b: колонка з помилкою (старий спосіб)", R1[b], Old(T, "b")},
        {"c: порожній текст не null",        R1[c], Old(T, "c")},
        {"без колонок з помилками: a",       R2[a], 1},
        {"без колонок з помилками: c",       R2[c], 1},
        {"порожня таблиця",                  R3[a] + R3[b] + R3[c], 0},
        {"очікувані значення a, b, c",       Text.Combine({Text.From(R1[a]), Text.From(R1[b]), Text.From(R1[c])}, ","), "1,2,1"}
    }
in
    Table.FromRows(
        List.Transform(Cases, (c) => {c{0}, Text.From(c{1}), Text.From(c{2}), c{1} = c{2}}),
        type table [Case = text, Actual = nullable text, Expected = nullable text, Pass = logical])
