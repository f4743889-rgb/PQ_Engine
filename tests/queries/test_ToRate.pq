// test_ToRate: перевірка розбору порогу. Очікування: у всіх рядках Pass = TRUE.
let
    Near = (a as nullable number, b as nullable number) as logical =>
        (a = null and b = null) or (a <> null and b <> null and Number.Abs(a - b) < 0.000000001),

    Good = {
        {null, null}, {"", null}, {"off", null}, {"Ні", null}, {false, null},
        {0, 0}, {0.05, 0.05}, {0.5, 0.5}, {"5%", 0.05}, {"100%", 1},
        {"0,05", 0.05}, {"0.05", 0.05}, {"1%", 0.01}, {1, 1}, {" 12,5 % ", 0.125}
    },
    // 50 = клітинка 5000% у відсотковому форматі; 5 і "5" без % теж помилка
    Bad = {"abc", 150, -1, true, "150%", 5, "5", 50, 1.5},

    Show = (x as any) as text => if x = null then "null" else Text.From(x),

    R1 = List.Transform(Good, (c) =>
            let a = try fn_ToRate(c{0})
            in  [Input    = Show(c{0}),
                 Expected = Show(c{1}),
                 Actual   = if a[HasError] then a[Error][Message] else Show(a[Value]),
                 Pass     = not a[HasError] and Near(a[Value], c{1})]),

    R2 = List.Transform(Bad, (v) =>
            let a = try fn_ToRate(v)
            in  [Input    = Show(v),
                 Expected = "НекоректнийПоріг",
                 Actual   = if a[HasError] then a[Error][Reason] else Show(a[Value]),
                 Pass     = a[HasError] and a[Error][Reason] = "НекоректнийПоріг"])
in
    Table.FromRecords(R1 & R2,
        type table [Input = text, Expected = text, Actual = nullable text, Pass = logical])
