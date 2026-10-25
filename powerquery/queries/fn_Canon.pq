// fn_Canon: канонічне написання значень довідника (правило регістру Engine)
// Значення довідників (TargetType, PeriodType тощо) порівнюються без урахування регістру
// і пробілів на краях: "regIStrY " і "Registry" - те саме призначення.
// Повертає функцію v -> написання з довідника; значення, якого в довіднику немає,
// повертається обрізаним як є (його показує pq_RegistryCheck); порожнє -> null.
// Довідник читається один раз на виклик fn_Canon, тому функцію створюють раз на таблицю:
//   CanonT = fn_Canon("tblTargetType", "TargetType"),  ... CanonT([TargetType])
// Таблицю читає сам запит-викликач (як fn_TemplateColumns), тому Formula.Firewall не виникає.
(dictTable as text, dictColumn as text) as function =>
let
    T    = try Excel.CurrentWorkbook(){[Name = dictTable]}[Content] otherwise null,
    Vals = if T = null or not Table.HasColumns(T, dictColumn) then {}
           else List.Select(
                    List.Transform(Table.Column(T, dictColumn), each Text.Trim(Text.From(_ ?? ""))),
                    each _ <> ""),
    // перше написання для кожного ключа в нижньому регістрі
    Map  = List.Accumulate(Vals, [],
               (s, v) => let k = Text.Lower(v)
                         in  if Record.HasFields(s, k) then s else Record.AddField(s, k, v))
in
    (v as any) as any =>
        if v = null then null
        else
            let t = Text.Trim(Text.From(v))
            in  if t = "" then null else Record.FieldOrDefault(Map, Text.Lower(t), t)