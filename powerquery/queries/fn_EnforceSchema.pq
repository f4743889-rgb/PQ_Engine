// fn_EnforceSchema - перевірка обов'язкових колонок і застосування політики помилок
// settings: результат fn_EffectiveSettings (Policy, Threshold, Culture); null = значення cfgSettings.
// Рішення «зупинити чи ні» приймає fn_LoadVerdict, те саме правило дає QualityFlag у звітах.
// Поріг діє лише для Report і Lenient: перевищення зупиняє запит помилкою ПеревищенийПоріг.
(raw as nullable table, rules as record, optional settings as nullable record) as table =>
let
    Eff     = settings ?? fn_EffectiveSettings(null),
    Policy  = Eff[Policy],

    // обов'язкові колонки перевіряємо лише для реальних даних
    Missing = if raw = null then {}
              else List.Difference(rules[RequiredSource], Table.ColumnNames(raw)),
    // підказка: у файлі є така сама назва з пробілами на краях (заголовки не обрізаються навмисно)
    Spaced  = if raw = null then {}
              else List.Select(Table.ColumnNames(raw), (c) => c <> Text.Trim(c) and List.Contains(Missing, Text.Trim(c))),
    SpacedHint = if List.IsEmpty(Spaced) then ""
                 else " У файлі є назви з пробілами на краях: '" & Text.Combine(Spaced, "', '") & "'.",
    // підказка: у файлі є така сама назва в іншому регістрі (назви колонок не нормалізуються
    // навмисно - діагностика замість мовчазного прирівнювання; рецензія коду 25.10.2026)
    Cased   = if raw = null then {}
              else List.Select(Table.ColumnNames(raw), (c) =>
                       List.AnyTrue(List.Transform(Missing, (m) =>
                           Text.Lower(Text.Trim(c)) = Text.Lower(m) and Text.Trim(c) <> m))),
    CasedHint  = if List.IsEmpty(Cased) then ""
                 else " У файлі є такі самі назви в іншому регістрі: '" & Text.Combine(Cased, "', '") & "'.",

    Data0   = fn_ApplyRules(raw, rules, Eff[Culture]),
    // Strict і поріг потребують перегляду даних до їх видачі: тоді один буфер замість 2-3 розборів
    // файла (рядки з помилками, кількість рядків, самі дані; аудит коду, розділ 2, v4.10).
    // Буфер зберігає клітинки-помилки (діагностика BufferOk); Report без порогу буфера не потребує.
    NeedScan = Policy = "Strict" or (try Eff[Threshold] <> null otherwise true),
    Data    = if raw <> null and NeedScan then (try Table.Buffer(Data0) otherwise Data0) else Data0,

    // обчислюються лише тоді, коли їх запитує fn_LoadVerdict або повідомлення про помилку
    ErrRows = Table.SelectRowsWithErrors(Data),
    ErrBuf  = try Table.Buffer(ErrRows) otherwise ErrRows,
    ErrCols = List.Select(Table.ColumnNames(Data),
                (c) => not Table.IsEmpty(Table.SelectRowsWithErrors(ErrBuf, {c}))),
    nAll    = Table.RowCount(Data),
    nErr    = Table.RowCount(ErrBuf),
    Rate    = if nAll = 0 then 0 else nErr / nAll,

    V   = fn_LoadVerdict(Policy, Missing,
              () => not Table.IsEmpty(ErrRows), () => Rate, () => Eff[Threshold]),
    // культура джерела (реєстр -> cfgSettings), як і для типізації (рецензія коду 25.10.2026)
    Pct = (x as number) as text => Number.ToText(x, "P1", Eff[Culture])
in
    if not V[Stop] then
        (if Policy = "Lenient"
         // помилки в клітинках стають null, типи зберігаються (Table.TransformColumns тут не годиться:
         // клітинку з помилкою він не передає у функцію; виявлено тестом L06, 16.10.2026)
         then Table.ReplaceErrorValues(Data, List.Transform(Table.ColumnNames(Data), each {_, null}))
         else Data)
    else if V[Reason] = "ВідсутніОбов'язковіПоля" then
        error Error.Record(V[Reason],
            "У джерелі для '" & rules[TargetType] & "' відсутні обов'язкові колонки: "
                & Text.Combine(Missing, ", ") & "." & SpacedHint & CasedHint, Missing)
    else if V[Reason] = "НевідомаПолітика" then
        error Error.Record(V[Reason],
            "ErrorPolicy має бути Strict, Report або Lenient, а не '" & Policy & "'.", Policy)
    else if V[Reason] = "ПомилкиТипізації" then
        error Error.Record(V[Reason],
            "'" & rules[TargetType] & "': " & Text.From(Table.RowCount(ErrBuf))
                & " рядків із помилками; колонки: " & Text.Combine(ErrCols, ", "), ErrCols)
    else
        error Error.Record(V[Reason],
            "'" & rules[TargetType] & "': рядків із помилками " & Text.From(nErr) & " з " & Text.From(nAll)
                & " (" & Pct(Rate) & "), поріг " & Pct(V[Threshold]) & ".",
            [ErrorRows = nErr, Rows = nAll, Rate = Rate, Threshold = V[Threshold]])