// test_RowFilters: fn_ApplyRowFilters - усі оператори, Use, регістр і пробіли, числові клітинки, помилки.
// Очікування: у всіх рядках Pass = TRUE. Дані задано в коді.
let
    Rules = (rows as list) as table =>
        #table(type table [TargetType = text, SourceColumnName = text, Operator = text,
                           Value = nullable text, Use = any], rows),
    // a: текст, b: число або текст, c: текст з пробілами / порожній
    Data = Table.TransformColumnTypes(
               #table({"a", "b", "c"},
                   {{"1", 10.5, "звичайний"}, {"Всього:", 99, "  РАЗОМ по групі "}, {"2", "х", null},
                    {"3", "abc", "  "}, {"4", "zz", "ok"}}),
               {{"a", type text}}),
    // клітинка-помилка в колонці c (рядок 5): правило по c не повинно ламати фільтр
    WithErr = Table.ReplaceValue(Data, "ok", null, (v, o, n) => if v = o then error "збій" else v, {"c"}),

    Drop = (rules as table, optional tbl as nullable table) as any =>
        // try навколо Dropped: правило компілюється ліниво, помилка виникає під час відбору рядків
        let r = try Text.From(fn_ApplyRowFilters(tbl ?? Data, rules, "T", "uk-UA")[Dropped])
        in  if r[HasError] then r[Error][Reason] else r[Value],
    Kept = (rules as table) as text =>
        Text.Combine(Table.Column(fn_ApplyRowFilters(Data, rules, "T", "uk-UA")[Data], "a"), ","),

    Cases = {
        {"StartsWith (регістр і пробіли у Value)", Drop(Rules({{"T", "a", "StartsWith", " всього ", null}})), "1"},
        {"Equals: кілька варіантів через |",       Drop(Rules({{"T", "b", "Equals", "х|x", null}})), "1"},
        {"Equals: число порівнюється як текст",    Drop(Rules({{"T", "b", "Equals", "99", null}})), "1"},
        {"Contains (регістр, пробіли в клітинці)", Drop(Rules({{"T", "c", "Contains", "разом", null}})), "1"},
        {"IsBlank: null і рядок з пробілів",       Drop(Rules({{"T", "c", "IsBlank", null, null}})), "2"},
        {"NotNumeric: число не відкидається",      Drop(Rules({{"T", "b", "NotNumeric", null, null}})), "3"},
        {"Use = FALSE вимикає правило",            Drop(Rules({{"T", "a", "StartsWith", "Всього", false}})), "0"},
        {"Use = «Ні» вимикає правило",            Drop(Rules({{"T", "a", "StartsWith", "Всього", "Ні"}})), "0"},
        {"правило іншого призначення не діє",      Drop(Rules({{"Інше", "a", "StartsWith", "Всього", null}})), "0"},
        {"колонки немає - правило пропущено",
            Text.From(fn_ApplyRowFilters(Data, Rules({{"T", "немає", "Equals", "1", null}}), "T", "uk-UA")[RulesSkipped]), "1"},
        {"невідомий оператор",                     Drop(Rules({{"T", "a", "Between", "1", null}})), "НевідомийОператорФільтра"},
        {"клітинка-помилка не збігається з IsBlank", Drop(Rules({{"T", "c", "IsBlank", null, null}}), WithErr), "2"},
        {"два правила: залишилися рядки",
            Kept(Rules({{"T", "a", "StartsWith", "Всього", null}, {"T", "b", "Equals", "х", null}})), "1,3,4"}
    }
in
    Table.FromRows(
        List.Transform(Cases, (c) => {c{0}, c{1}, c{2}, c{1} = c{2}}),
        type table [Case = text, Actual = nullable text, Expected = nullable text, Pass = logical])