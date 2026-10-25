// test_HeadersConsistency: швидке читання заголовків (fn_ParseHeaders) дає ті самі колонки, що й повне
// читання (fn_ParseSource), для кожного тестового джерела (аудит 1.8, BOM).
// Лише SourceID з tblTestExpected (реальні джерела тестової книги не беруться). Запит відкриває файли,
// тому сам читає книгу і викликає лише функції. Очікування: у всіх рядках Pass = TRUE.
// Для КОЖНОГО SourceID з tblTestExpected є рядок (v4.15, рецензія коду 25.10.2026, CR-02):
//   PASS / FAIL - заголовки порівняно;
//   SKIP        - файл не читається (немає, пошкоджений, непідтримуваний, немає аркуша); Pass = TRUE
//                 лише якщо так і очікується (Status у tblTestExpected: FILE_ERROR, READ_ERROR, NODATA);
//   NOT_ACTIVE  - джерела немає серед активних (Enabled = FALSE); Pass = TRUE лише для Status = ABSENT.
// Так пропущена перевірка видна в результаті і не може тихо зникнути.
let
    Exp   = Excel.CurrentWorkbook(){[Name = "tblTestExpected"]}[Content],
    Norm  = (v as any) as text => Text.Trim(Text.From(v ?? "")),
    ExpR  = List.Select(Table.ToRecords(Exp), each Norm([SourceID]) <> ""),
    Ids   = List.Distinct(List.Transform(ExpR, each Norm([SourceID]))),
    // очікуваний Status джерела з tblTestExpected (верхній регістр; "" якщо не задано)
    StatusOf = (id as text) as text =>
        let m = List.Select(ExpR, each Norm([SourceID]) = id)
        in  if List.IsEmpty(m) then "" else Text.Upper(Norm(Record.FieldOrDefault(m{0}, "Status", null))),
    Active = List.Select(Table.ToRecords(fn_ActiveSources(fn_LoadConfig()[Registry])),
               (r) => List.Contains(Ids, Norm(r[SourceID]))),
    ActiveIds = List.Transform(Active, each Norm([SourceID])),
    SkipOk = {"FILE_ERROR", "READ_ERROR", "NODATA"},

    Check = (r as record) as record =>
        let
            id   = Norm(r[SourceID]),
            st   = StatusOf(id),
            Src  = fn_ResolveSource(r),
            Opt  = fn_Options(r),
            H    = try fn_ParseHeaders(Src, Opt),
            Full = try Table.ColumnNames(fn_ParseSource(Src, Opt)),
            Hs   = if H[HasError] then "помилка: " & H[Error][Message] else Text.Combine(H[Value], "|")
        in
            if Src[ErrorCode] <> null then
                [SourceID = id, Result = "SKIP", Reason = Src[ErrorCode], Headers = null, FullRead = null,
                 Pass = List.Contains(SkipOk, st)]
            else if Full[HasError] then
                [SourceID = id, Result = "SKIP", Reason = Full[Error][Reason] ?? Full[Error][Message],
                 Headers = Hs, FullRead = null, Pass = List.Contains(SkipOk, st)]
            else
                let full = Text.Combine(Full[Value], "|"), ok = Hs = full
                in  [SourceID = id, Result = if ok then "PASS" else "FAIL", Reason = null,
                     Headers = Hs, FullRead = full, Pass = ok],

    NotActive = List.Transform(List.Select(Ids, each not List.Contains(ActiveIds, _)),
                    (id) => [SourceID = id, Result = "NOT_ACTIVE", Reason = "немає серед активних джерел",
                             Headers = null, FullRead = null,
                             Pass = StatusOf(id) = "ABSENT"])
in
    Table.Sort(
        Table.FromRecords(List.Transform(Active, Check) & NotActive,
            type table [SourceID = text, Result = text, Reason = nullable text,
                        Headers = nullable text, FullRead = nullable text, Pass = logical]),
        {{"Pass", Order.Ascending}, {"SourceID", Order.Ascending}})
