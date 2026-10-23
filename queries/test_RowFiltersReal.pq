// test_RowFiltersReal: діагностика tblRowFilters на реальних файлах призначення.
// Нічого не змінює в конвеєрі: читає активні джерела Target через fn_ReadSource,
// застосовує fn_ApplyRowFilters і показує, що саме відсіклося.
//
// Колонки результату (один рядок = одне джерело):
//   Status         OK / NODATA (необов'язковий файл відсутній) / ERROR (див. Error)
//   RowsRaw        рядків прочитано
//   Dropped        відсічено рядків; RowsKept = RowsRaw - Dropped
//   RulesApplied   правил застосовано; RulesSkipped = правил пропущено (колонки немає у файлі)
//   DroppedValues  які значення колонок з правил були у відсічених рядках (до 8 на колонку)
//   DroppedSample  перші 10 відсічених рядків: номер рядка даних (без шапки) і значення
//   KeptNonNumeric нечислові значення колонок з правил, що ЛИШИЛИСЯ в даних (до 8 на
//                  колонку): допомагає знайти службові рядки, яких правила не охопили
//                  (наприклад, "Разом", "-"); для текстових колонок типу "№ з/п"
//                  частина значень може бути цілком легітимною
//
// Завантажувати лише як підключення або на окремий аркуш; запит читає книгу і файли
// сам (правило трьох видів запитів), а fn_* лише викликає.
let
    Target  = "Registry",
    Culture = "uk-UA",

    // Книгу читаємо прямо тут, а не через fn_LoadConfig: таблиця правил передається у функцію Check,
    // що відкриває файли, і така комбінація дає Formula.Firewall (виявлено 16.10.2026).
    // Імена таблиць - з cfgSettings[Tables], як і скрізь.
    Reg   = Excel.CurrentWorkbook(){[Name = cfgSettings[Tables][Registry]]}[Content],
    Rules = Excel.CurrentWorkbook(){[Name = cfgSettings[Tables][RowFilters]]}[Content],

    TT   = Text.Lower(Text.Trim(Target)),
    Rows = List.Select(Table.ToRecords(fn_ActiveSources(Reg)),
               (r) => Text.Lower(Text.Trim(Text.From(Record.FieldOrDefault(r, "TargetType", null) ?? ""))) = TT),

    RuleCols = List.Distinct(List.Transform(
                   Table.ToRecords(Table.SelectRows(Rules,
                       each Text.Lower(Text.Trim(Text.From([TargetType] ?? ""))) = TT)),
                   (x) => Text.Trim(Text.From(x[SourceColumnName])))),

    // безпечний показ значення клітинки (помилкові клітинки та null не ламають звіт)
    Show = (v as any) as text => try Text.From(v ?? "(порожньо)") otherwise "#ПОМИЛКА",
    IsNum = (s as text) as logical => (try Number.FromText(s, Culture) otherwise null) <> null,

    Defaults = [
        SourceID = null, PeriodType = null, PeriodNumber = null, Status = null,
        RowsRaw = null, Dropped = null, RowsKept = null,
        RulesApplied = null, RulesSkipped = null,
        DroppedValues = null, DroppedSample = null, KeptNonNumeric = null, Error = null
    ],

    Check = (r as record) as record =>
        let
            Head = [
                SourceID     = Text.From(fn_Opt(r, "SourceID", "")),
                PeriodType   = Text.From(fn_Opt(r, "PeriodType", "")),
                PeriodNumber = Text.From(fn_Opt(r, "PeriodNumber", ""))
            ],
            Att = try
                let
                    Raw0 = fn_ReadSource(r),
                    Body =
                        if Raw0 = null then
                            [Status = "NODATA"]
                        else
                            let
                                Raw     = Table.Buffer(Raw0),      // один прохід по файлу
                                Idx     = Table.AddIndexColumn(Raw, "_RowNo", 1, 1, Int64.Type),
                                F       = fn_ApplyRowFilters(Idx, Rules, Target, Culture),
                                Kept    = F[Data],
                                Drp     = Table.RemoveColumns(
                                              Table.NestedJoin(Idx, "_RowNo", Kept, "_RowNo", "k", JoinKind.LeftAnti),
                                              "k"),
                                Present = List.Select(RuleCols, each List.Contains(Table.ColumnNames(Raw), _)),
                                Sample  = Text.Combine(
                                    List.Transform(Table.ToRecords(Table.FirstN(Drp, 10)),
                                        (d) => "#" & Text.From(d[_RowNo]) & ": "
                                               & Text.Combine(List.Transform(Present,
                                                     (c) => c & "=" & Show(Record.Field(d, c))), "; ")),
                                    "#(lf)"),
                                DropVals = Text.Combine(
                                    List.Transform(Present,
                                        (c) => c & ": " & Text.Combine(
                                            List.FirstN(List.Distinct(List.Transform(Table.Column(Drp, c), Show)), 8),
                                            " | ")),
                                    "#(lf)"),
                                KeptNN = Text.Combine(
                                    List.Transform(Present,
                                        (c) => c & ": " & Text.Combine(
                                            List.FirstN(
                                                List.Distinct(
                                                    List.Select(
                                                        List.Transform(Table.Column(Kept, c), Show),
                                                        (s) => s <> "(порожньо)" and s <> "#ПОМИЛКА" and not IsNum(Text.Trim(s)))),
                                                8),
                                            " | ")),
                                    "#(lf)")
                            in
                                [
                                    Status         = "OK",
                                    RowsRaw        = Table.RowCount(Raw),
                                    Dropped        = F[Dropped],
                                    RowsKept       = Table.RowCount(Raw) - F[Dropped],
                                    RulesApplied   = F[RulesApplied],
                                    RulesSkipped   = F[RulesSkipped],
                                    DroppedValues  = DropVals,
                                    DroppedSample  = Sample,
                                    KeptNonNumeric = KeptNN
                                ]
                in
                    Body
        in
            if Att[HasError]
            then Record.Combine({Defaults, Head, [Status = "ERROR", Error = Att[Error][Message]]})
            else Record.Combine({Defaults, Head, Att[Value]}),

    Result = Table.FromRecords(
                 List.Transform(Rows, Check),
                 type table [
                     SourceID = nullable text, PeriodType = nullable text,
                     PeriodNumber = nullable text, Status = nullable text,
                     RowsRaw = nullable Int64.Type, Dropped = nullable Int64.Type, RowsKept = nullable Int64.Type,
                     RulesApplied = nullable Int64.Type, RulesSkipped = nullable Int64.Type,
                     DroppedValues = nullable text, DroppedSample = nullable text,
                     KeptNonNumeric = nullable text, Error = nullable text
                 ],
                 MissingField.UseNull)
in
    Result