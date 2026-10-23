// ============================================================================
// fn_ApplyRowFilters  (відсікання службових рядків до типізації)
//
// Прибирає з таблиці службові рядки (підсумки "Всього", повторні шапки, "х"
// замість значення) за правилами з таблиці tblRowFilters. Рядки не змінюються,
// а лише відсікаються, кількість відсічених повертається для зведення якості.
//
// Параметри:
//   tbl     - таблиця з іменами колонок, як у джерелі (SourceColumnName).
//             Викликається з fn_ApplyRulesEx до вибору колонок схеми, перейменування і типів;
//             пробіли на краях значень ігноруються самою функцією
//   rules   - вміст tblRowFilters (Excel.CurrentWorkbook читає запит-викликач,
//             а не ця функція): TargetType, SourceColumnName, Operator, Value, [Use]
//   target  - TargetType поточного джерела
//   culture - культура для NotNumeric; за замовчуванням cfgSettings[Culture]
//
// Оператори (регістр і крайні пробіли не важливі; порівняння нечутливе до регістру):
//   Equals      значення клітинки дорівнює одному з варіантів Value
//   StartsWith  значення починається з одного з варіантів Value
//   Contains    значення містить один з варіантів Value
//   IsBlank     клітинка порожня (Value не потрібен)
//   NotNumeric  клітинка непорожня, але не читається як число (Value не потрібен)
//               УВАГА: приховує і справжні помилки даних, тому використовувати
//               лише для колонок, де нечислові рядки точно службові
//
// Value: кілька варіантів розділяються символом "|" (достатньо збігу з будь-яким).
//   Приклад: "х|x" = кирилична х АБО латинська x
// Use:  як fn_ToBool: порожнє = увімкнено, значення з cfgSettings[TrueValues] = увімкнено,
//       будь-яке інше (FALSE, Ні, 0 тощо) = вимкнено.
//
// Правило, колонки якого немає в таблиці, пропускається (лічильник RulesSkipped),
// бо різні файли одного призначення можуть мати різний склад колонок.
// Невідомий Operator дає помилку НевідомийОператорФільтра.
//
// Повертає запис: [Data = таблиця, Dropped = число відсічених рядків,
//                  RulesApplied = застосовано правил, RulesSkipped = пропущено правил]
// ============================================================================
(tbl as table, rules as table, target as text, optional culture as nullable text) as record =>
let
    Cult = culture ?? cfgSettings[Culture],

    // нормалізація тексту для порівняння: обрізання пробілів, нижній регістр
    Key  = (t as nullable text) as text => Text.Lower(Text.Trim(t ?? "")),
    // значення клітинки як нормалізований текст; помилкова клітинка та null дають ""
    Norm = (x as any) as text => try Text.Lower(Text.Trim(Text.From(x ?? ""))) otherwise "",

    Cols = Table.ColumnNames(tbl),

    // правила цього призначення, які не вимкнено
    Mine = Table.SelectRows(rules, each
               Key(Text.From([TargetType] ?? "")) = Key(target)
               and fn_ToBool([Use]?, true)),

    // компіляція одного правила у функцію "рядок -> logical" (true = службовий рядок)
    Compile = (r as record) as function =>
        let
            col  = Text.Trim(Text.From(r[SourceColumnName])),
            op   = Key(Text.From(r[Operator] ?? "")),
            vals = List.Select(
                       List.Transform(Text.Split(Text.From(r[Value] ?? ""), "|"), Key),
                       each _ <> ""),
            test =
                if op = "equals" then
                    ((s as text) => List.Contains(vals, s))
                else if op = "startswith" then
                    ((s as text) => List.AnyTrue(List.Transform(vals, (v) => Text.StartsWith(s, v))))
                else if op = "contains" then
                    ((s as text) => List.AnyTrue(List.Transform(vals, (v) => Text.Contains(s, v))))
                else if op = "isblank" then
                    ((s as text) => s = "")
                else if op = "notnumeric" then
                    ((s as text) => s <> "" and (try Number.FromText(s, Cult) otherwise null) = null)
                else
                    error Error.Record("НевідомийОператорФільтра",
                                       "Оператор '" & op & "' не підтримується.", col)
        in
            (row as record) as logical => test(Norm(Record.Field(row, col))),

    Valid    = Table.SelectRows(Mine, each List.Contains(Cols, Text.Trim(Text.From([SourceColumnName] ?? "")))),
    Fns      = List.Transform(Table.ToRecords(Valid), Compile),
    HasRules = List.Count(Fns) > 0,

    // буфер потрібен, щоб підрахувати відсічені рядки без повторного читання файла
    Buf  = if HasRules then (try Table.Buffer(tbl) otherwise tbl) else tbl,
    Kept = if HasRules
           then Table.SelectRows(Buf, (row) => not List.AnyTrue(List.Transform(Fns, (f) => f(row))))
           else tbl,
    Dropped = if HasRules then Table.RowCount(Buf) - Table.RowCount(Kept) else 0
in
    [
        Data         = Kept,
        Dropped      = Dropped,
        RulesApplied = List.Count(Fns),
        RulesSkipped = Table.RowCount(Mine) - List.Count(Fns)
    ]