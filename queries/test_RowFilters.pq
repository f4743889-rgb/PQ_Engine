// test_RowFilters: швидка перевірка fn_ApplyRowFilters на ваших двох прикладах.
// Очікування: Dropped = 3 (рядки "2 / х", "3 / x", "Всього:"), Kept = 2 рядки (1 і 4),
// RulesApplied = 2, RulesSkipped = 1 (правило для неіснуючої колонки), OK = true.
// Завантажувати лише як підключення; після перевірки запит можна видалити.
let
    Data = #table({"№ з/п", "Ціна, грн."},
        {{"1", "10,5"}, {"2", "х"}, {"3", "x"}, {"Всього:", "123"}, {"4", "20"}}),
    Rules = #table({"TargetType", "SourceColumnName", "Operator", "Value", "Use"},
        {
            {"Registry", "№ з/п",      "StartsWith", "Всього", null},
            {"Registry", "Ціна, грн.", "Equals",     "х|x",    null},
            {"Registry", "Немає такої", "Equals",    "z",      null},
            {"Other",    "№ з/п",      "Equals",     "1",      null}
        }),
    R = fn_ApplyRowFilters(Data, Rules, "Registry", "uk-UA"),
    Kept = Table.ToRows(R[Data])
in
    [
        Dropped      = R[Dropped],
        RulesApplied = R[RulesApplied],
        RulesSkipped = R[RulesSkipped],
        Kept         = Kept,
        OK           = R[Dropped] = 3 and R[RulesApplied] = 2 and R[RulesSkipped] = 1
                       and List.Count(Kept) = 2 and Kept{0}{0} = "1" and Kept{1}{0} = "4"
    ]