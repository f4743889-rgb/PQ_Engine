// fn_RegistryIssues: усі перевірки реєстру джерел в одному місці (аудит D2)
// Використовують pq_RegistryCheck (звіт) і fn_ReadPurpose (зупиняє завантаження призначення,
// якщо для нього є блокуюча проблема), тож звіт і завантаження не можуть розійтися.
// Правила: SourceID унікальний у всьому реєстрі; той самий файл (шлях + лист) і той самий період
// не повторюються в межах призначення. Перевіряються лише активні рядки (fn_ActiveSources).
// options: [AllowDuplicatePeriods = true, AllowDuplicateFiles = true] знімають блокування цих дублів.
// Результат: Issue, Key, Count, Targets (список TargetType), Code, Blocking, Priority (менше = важливіше).
(registry as table, optional options as nullable record) as table =>
let
    O            = options ?? [],
    AllowPeriods = fn_ToBool(Record.FieldOrDefault(O, "AllowDuplicatePeriods", null), false),
    AllowFiles   = fn_ToBool(Record.FieldOrDefault(O, "AllowDuplicateFiles", null), false),
    Tbl          = cfgSettings[Tables],
    Recs         = Table.ToRecords(fn_ActiveSources(registry)),

    TT   = (r as record) as text => Text.From(Record.FieldOrDefault(r, "TargetType", null) ?? ""),
    Id   = (r as record) as text => Text.From(fn_Opt(r, "SourceID", "")),
    PKey = (r as record) as text =>
               Text.From(fn_Opt(r, "PeriodType", "")) & "|" & Text.From(fn_Opt(r, "PeriodNumber", "")),
    FKey = (r as record) as nullable text =>
               let p = fn_SourcePath(r)
               in  if p = "" then null
                   else Text.Lower(p) & "|" & Text.Lower(Text.From(fn_Opt(r, "TableSheetName", ""))),

    Issue = (issue as text, code as text, key as text, rows as list, blocking as logical,
             priority as number) as record =>
        [Issue = issue, Key = key, Count = List.Count(rows),
         Targets = List.Distinct(List.Select(List.Transform(rows, TT), each _ <> "")),
         Code = code, Blocking = blocking, Priority = priority],

    // дублікати: групи за ключем (у межах призначення, якщо perTarget) з 2+ рядків
    Dups = (keyFn as function, perTarget as logical, issue as text, code as text,
            blocking as logical, priority as number) as list =>
        let
            K      = List.Select(List.Transform(Recs, (r) => [R = r, K = keyFn(r)]),
                         each [K] <> null and [K] <> ""),
            G      = (x as record) as text => (if perTarget then TT(x[R]) & "||" else "") & x[K],
            Groups = List.Distinct(List.Transform(K, G)),
            Out    = List.Transform(Groups,
                         (g) => let m = List.Select(K, (x) => G(x) = g)
                                in  Issue(issue, code, m{0}[K], List.Transform(m, each [R]), blocking, priority))
        in
            List.Select(Out, each [Count] > 1),

    DupId     = Dups(Id,   false, "Дубль SourceID", "ДублюванняSourceID", true, 1),
    DupFile   = Dups(FKey, true,  "Дубль файлу (шлях + лист) у призначенні", "ДублюванняФайлу", not AllowFiles, 2),
    // джерела без періоду (PeriodType і PeriodNumber порожні) у перевірці періоду не беруть участі:
    // той самий файл двічі ловить DupFile (аудит 1.5, v4.10)
    PKeySet   = (r as record) as nullable text => let k = PKey(r) in if k = "|" then null else k,
    DupPeriod = Dups(PKeySet, true, "Дубль TargetType + період", "ДублюванняПеріоду", not AllowPeriods, 3),

    Blank   = List.Select(Recs, (r) => Text.Trim(Id(r)) = ""),
    BlankId = if List.IsEmpty(Blank) then {}
              else {Issue("SourceID порожній", "ПорожнійSourceID", "", Blank, false, 10)},

    // період, якого немає в tblPeriodRules (тип | номер до MaxNumber)
    PeriodRules = try Excel.CurrentWorkbook(){[Name = Tbl[PeriodRules]]}[Content]
                  otherwise #table({"PeriodType", "MaxNumber"}, {}),
    ValidKeys = List.Combine(List.Transform(
                    Table.ToRecords(Table.SelectRows(PeriodRules, each [PeriodType] <> null and [MaxNumber] <> null)),
                    (p) => List.Transform(List.Numbers(1, Number.From(p[MaxNumber])),
                               (n) => Text.Lower(Text.Trim(Text.From(p[PeriodType]))) & "|" & Text.From(n)))),
    BadP      = List.Select(Recs, (r) => PKey(r) <> "|" and not List.Contains(ValidKeys, Text.Lower(PKey(r)))),
    BadPeriod = List.Transform(List.Distinct(List.Transform(BadP, PKey)),
                    (k) => Issue("Недопустимий період (тип | номер)", "НедопустимийПеріод", k,
                                 List.Select(BadP, (r) => PKey(r) = k), false, 11)),

    // некоректний ErrorThreshold (fn_ToRate)
    BadThr = List.Transform(
                 List.Select(Recs, (r) => (try fn_ToRate(fn_Opt(r, "ErrorThreshold", null)))[HasError]),
                 (r) => Issue("Некоректний ErrorThreshold", "НекоректнийПоріг",
                              Id(r) & " | " & Text.From(fn_Opt(r, "ErrorThreshold", "")), {r}, false, 12)),

    // TargetType, якого немає в довіднику (без урахування регістру)
    TTDict = List.Transform(
                 Table.Column(try Excel.CurrentWorkbook(){[Name = Tbl[TargetTypes]]}[Content]
                              otherwise #table({"TargetType"}, {}), "TargetType"),
                 each Text.Lower(Text.Trim(Text.From(_ ?? "")))),
    BadTTr = List.Select(Recs, (r) => TT(r) <> "" and not List.Contains(TTDict, Text.Lower(TT(r)))),
    BadTT  = List.Transform(List.Distinct(List.Transform(BadTTr, TT)),
                 (t) => Issue("TargetType немає в довіднику", "НевідомеПризначення", t,
                              List.Select(BadTTr, (r) => TT(r) = t), false, 13)),

    All = DupId & DupFile & DupPeriod & BlankId & BadPeriod & BadThr & BadTT,
    Typ = type table [Issue = text, Key = text, Count = Int64.Type, Targets = list,
                      Code = text, Blocking = logical, Priority = Int64.Type]
in
    if List.IsEmpty(All) then #table(Typ, {})
    else Table.Sort(Table.FromRecords(All, Typ), {{"Priority", Order.Ascending}, {"Key", Order.Ascending}})