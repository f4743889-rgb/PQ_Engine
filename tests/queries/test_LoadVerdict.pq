// test_LoadVerdict: модульний тест правила «чи зупиниться завантаження». Очікування: Pass = TRUE у всіх рядках.
// Останній випадок перевіряє лінивість: при Strict некоректний поріг не обчислюється.
let
    Bad   = () => error Error.Record("НекоректнийПоріг", "тест лінивості"),
    Cases = {
        {"Report без помилок",            "Report",  {},    false, 0,    () => null, false, null,                       "OK"},
        {"Report з помилками, без порогу", "Report",  {},    true,  0.2,  () => null, false, null,                       "WARN"},
        {"Report, поріг перевищено",       "Report",  {},    true,  0.2,  () => 0.1,  true,  "ПеревищенийПоріг",         "FAIL"},
        {"Report, поріг не перевищено",    "Report",  {},    true,  0.05, () => 0.1,  false, null,                       "WARN"},
        {"Report, частка = поріг",         "Report",  {},    true,  0.1,  () => 0.1,  false, null,                       "WARN"},
        {"Strict з помилками",             "Strict",  {},    true,  0.2,  () => null, true,  "ПомилкиТипізації",         "FAIL"},
        {"Strict без помилок",             "Strict",  {},    false, 0,    () => null, false, null,                       "OK"},
        {"Strict ігнорує поріг",           "Strict",  {},    true,  0.2,  () => 0.5,  true,  "ПомилкиТипізації",         "FAIL"},
        {"Lenient, поріг перевищено",      "Lenient", {},    true,  0.2,  () => 0.1,  true,  "ПеревищенийПоріг",         "FAIL"},
        {"Відсутня обов'язкова колонка",   "Report",  {"X"}, false, 0,    () => null, true,  "ВідсутніОбов'язковіПоля",  "FAIL"},
        {"Невідома політика",              "Abc",     {},    false, 0,    () => null, true,  "НевідомаПолітика",         "FAIL"},
        {"Strict не обчислює поріг",       "Strict",  {},    false, 0,    Bad,        false, null,                       "OK"}
    },
    Run = (c as list) as record =>
        let
            hasErr = c{3}, rate = c{4},
            // поля рахуються всередині try, щоб помилка лінивого поля не зупинила весь тест
            V  = try let v = fn_LoadVerdict(c{1}, c{2}, () => hasErr, () => rate, c{5})
                     in  [Stop = v[Stop], Reason = v[Reason], Flag = v[Flag]],
            ok = not V[HasError] and V[Value][Stop] = c{6} and V[Value][Reason] = c{7}
                 and V[Value][Flag] = c{8}
        in
            [Case = c{0},
             Expected = (if c{6} then "Stop " else "") & (c{7} ?? "") & " / " & c{8},
             Actual = if V[HasError] then "ПОМИЛКА: " & V[Error][Message]
                      else (if V[Value][Stop] then "Stop " else "") & (V[Value][Reason] ?? "") & " / " & V[Value][Flag],
             Pass = ok]
in
    Table.FromRecords(List.Transform(Cases, Run),
        type table [Case = text, Expected = text, Actual = text, Pass = logical])
