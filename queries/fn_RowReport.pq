// fn_RowReport: звіт якості для одного рядка реєстру (не падає)
(row as record, rulesTry as record) as record =>
let
    Fld   = (n as text) as any => Record.FieldOrDefault(row, n, null),
    SID   = Text.From(Fld("SourceID") ?? ""),
    TT    = Text.Trim(Text.From(Fld("TargetType") ?? "")),
    Req   = fn_ToBool(Fld("IsRequired"), true),

    ColsEmpty = #table(cfgReportTypes[Columns], {}),
    ErrsEmpty = #table(cfgReportTypes[Errors], {}),

    Base = [SourceID = SID, TargetType = TT,
            PeriodType = Fld("PeriodType"), PeriodNumber = Fld("PeriodNumber"),
            SourceRequired = Req, Status = null, Message = null,
            RawRows = null, Rows = null, EmptyRowsRemoved = null,
            RowsDropped = null, FilterRulesSkipped = null, ErrorRows = null,
            ErrorColumns = null, MissingRequired = null, MissingOptional = null, EmptyColumns = null,
            DuplicateHeaders = null],

    Fail = (status as text, msg as nullable text) as record =>
        [Summary = Base & [Status = status, Message = msg], Columns = ColsEmpty, Errors = ErrsEmpty],

    // індикатор якості: FAIL = завантаження через fn_ReadPurpose зупиниться, WARN = є помилки, OK.
    // Правило те саме, що в fn_EnforceSchema (fn_LoadVerdict), налаштування - fn_EffectiveSettings.
    Eff    = fn_EffectiveSettings(row),
    ThrTry = try Eff[Threshold],
    Quality = (s as record) as record =>
        let
            rows = s[Rows], errs = s[ErrorRows],
            rate = if rows = null or errs = null then null else if rows = 0 then 0 else errs / rows,
            thr  = if ThrTry[HasError] then null else ThrTry[Value],
            st   = Record.FieldOrDefault(cfgStatuses, s[Status] ?? "", [Blocking = false]),
            miss = if (s[MissingRequired] ?? "") = "" then {} else {s[MissingRequired]},
            flag =
                if ThrTry[HasError] then "BAD_THRESHOLD"
                else if st[Blocking] then "FAIL"
                else if rate = null then null
                else fn_LoadVerdict(Eff[Policy], miss, () => errs > 0, () => rate, () => thr)[Flag]
        in
            [ErrorRate = rate, ErrorThreshold = thr, QualityFlag = flag],

    Src    = fn_ResolveSource(row),
    Parsed = if Src[ErrorCode] <> null then null
             else try fn_ParseSource(Src, fn_Options(row)),

    // примусове обчислення всередині try: підсумок як значення, таблиці буферизовані
    Forced =
        if Parsed = null or Parsed[HasError] or rulesTry[HasError] then null
        else try
            let
                Rep = fn_SchemaReport(Parsed[Value], rulesTry[Value], SID, Eff[Culture]),
                S   = Rep[Summary]
            in
                [S = Record.FromList(Record.FieldValues(S), Record.FieldNames(S)),
                 C = Table.Buffer(Rep[Columns]),
                 E = Table.Buffer(Rep[Errors])],

    Res =
    if rulesTry[HasError] then Fail("NO_SCHEMA", rulesTry[Error][Message])
    else if Src[ErrorCode] <> null then
        (if Req then Fail("FILE_ERROR", Src[ErrorMessage]) else Fail("NODATA", Src[ErrorMessage]))
    else if Parsed[HasError] then Fail("READ_ERROR", Parsed[Error][Message])
    else if Forced[HasError] then Fail("READ_ERROR", Forced[Error][Message])
    else [Summary = Base & Forced[Value][S], Columns = Forced[Value][C], Errors = Forced[Value][E]],
    Final = Res & [Summary = Res[Summary] & Quality(Res[Summary])]
in
    Final