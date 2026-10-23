// fn_SchemaReport: звіт про відповідність даних схемі (не залежить від ErrorPolicy)
(raw as nullable table, rules as record,
 optional sourceId as nullable text, optional culture as nullable text,
 optional maxPerColumn as nullable number) as record =>
let
    MaxN    = maxPerColumn ?? 5,
    RawCols = if raw = null then {} else Table.ColumnNames(raw),
    nRaw    = if raw = null then 0 else Table.RowCount(raw),

    // дублікати заголовків: Power Query лишає першу колонку з назвою, а повтори перейменовує
    // на "Назва_1", "Назва_2"; дубль колонки схеми означає, що Engine мовчки бере першу копію
    DupOf = (c as text) as nullable text =>
        let
            p    = Text.PositionOf(c, "_", Occurrence.Last),
            base = if p > 0 then Text.Start(c, p) else null,
            num  = if p > 0 then Text.Range(c, p + 1) else ""
        in
            if base <> null and num <> "" and Text.Select(num, {"0".."9"}) = num
               and List.Contains(RawCols, base) then base else null,
    DupBases    = List.Distinct(List.RemoveNulls(List.Transform(RawCols, DupOf))),
    DupInSchema = List.Select(DupBases, each List.Contains(rules[SourceCols], _)),

    // той самий конвеєр, що в Enforcer, але без політики: помилки лишаються в клітинках
    Ex       = fn_ApplyRulesEx(raw, rules, culture),
    Data     = Ex[Data],
    nDropped = Ex[Dropped],
    Indexed0 = Table.AddIndexColumn(Data, "__RowNo", 1, 1, Int64.Type),
    Indexed  = try Table.Buffer(Indexed0) otherwise Indexed0,
    nRows    = Table.RowCount(Indexed),
    // рядки з помилками відбираються один раз; перевірки колонок шукають помилки лише серед них
    // (v4.9, патч K: GO3, 81 тис. рядків, 16 колонок - 7,1 с -> 0,5 с, результат той самий).
    // Буфер зберігає клітинки-помилки (перевірено діагностикою BufferOk).
    ErrRows0 = Table.SelectRowsWithErrors(Indexed),
    ErrRows  = try Table.Buffer(ErrRows0) otherwise ErrRows0,
    nErrRows = Table.RowCount(ErrRows),

    // null по всіх колонках схеми - одним Table.Profile (v4.9, патч L); колонки з помилками - старим способом
    TargetCols = List.Transform(rules[Columns], each _[Target]),
    ErrCols    = if nErrRows = 0 then {}
                 else List.Select(TargetCols, (c) => Table.RowCount(Table.SelectRowsWithErrors(ErrRows, {c})) > 0),
    NullMap    = fn_ColumnNullCounts(Indexed, TargetCols, ErrCols),

    // --- аналіз однієї колонки ---
    Inspect = (rule as record) as record =>
        let
            c      = rule[Target],
            inSrc  = List.Contains(RawCols, rule[Source]),
            ColErr = Table.SelectRowsWithErrors(ErrRows, {c}),
            nErr   = Table.RowCount(ColErr),
            Cells  = List.Transform(
                        Table.ToRecords(Table.FirstN(ColErr, MaxN)),
                        (r) =>
                            let e = (try Record.Field(r, c))[Error]
                            in  [SourceID = sourceId, TargetType = rules[TargetType], Column = c,
                                 RowNo = r[__RowNo], Reason = e[Reason], Message = e[Message],
                                 Detail = try Text.From(e[Detail]) otherwise null]),
            nNull  = Record.Field(NullMap, c)
        in
            [SourceID     = sourceId,
             TargetType   = rules[TargetType],
             Source       = rule[Source],
             Target       = c,
             Required     = rule[Required],
             ExpectedType = Record.FieldOrDefault(rule, "TypeKey", null),
             InSource     = inSrc,
             Rows         = nRows,
             NullCount    = nNull,
             ErrorCount   = nErr,
             ErrorSamples = Text.Combine(
                                List.Transform(List.FirstN(Cells, 3),
                                    (x) => "№" & Text.From(x[RowNo]) & ": " & (x[Detail] ?? x[Message])),
                                " | "),
             ErrCells     = Cells],

    // List.Buffer: ColList переглядають чотири рази (ColTable, AllCells, EmptyCols, ErrorColumns);
    // без буфера кожен перегляд заново виконував усі перевірки колонок (v4.9, патч J).
    // Буферизуються записи; їхні поля обчислюються один раз і далі беруться з того самого запису.
    ColList = List.Buffer(List.Transform(rules[Columns], Inspect)),

    ColTable = Table.FromRecords(List.Transform(ColList, each Record.RemoveFields(_, {"ErrCells"})),
                   cfgReportTypes[Columns], MissingField.UseNull),

    AllCells = List.Combine(List.Transform(ColList, each _[ErrCells])),
    ErrTable = if List.IsEmpty(AllCells) then #table(cfgReportTypes[Errors], {})
               else Table.FromRecords(AllCells, cfgReportTypes[Errors], MissingField.UseNull),

    // --- підсумок ---
    MissReq = if raw = null then {}
              else List.Transform(
                       List.Select(rules[Columns], each _[Required] and not List.Contains(RawCols, _[Source])),
                       each _[Source]),
    MissOpt = if raw = null then {}
              else List.Transform(
                       List.Select(rules[Columns], each not _[Required] and not List.Contains(RawCols, _[Source])),
                       each _[Source]),
    EmptyCols = if nRows = 0 then {}
                else List.Transform(List.Select(ColList, each _[NullCount] = nRows), each _[Target]),

    StatusCode =
        if raw = null then "NODATA"
        else if not List.IsEmpty(MissReq) or nErrRows > 0 then "ERROR"
        else if nRows = 0 then "EMPTY"
        else if not List.IsEmpty(MissOpt) or not List.IsEmpty(EmptyCols)
                or not List.IsEmpty(DupInSchema) then "WARN"
        else "OK",

    SummaryRec = [
        SourceID         = sourceId,
        TargetType       = rules[TargetType],
        Status           = StatusCode,
        RawRows          = nRaw,
        Rows             = nRows,
        EmptyRowsRemoved = nRaw - nRows - nDropped,
        RowsDropped      = nDropped,
        FilterRulesSkipped = Ex[RulesSkipped],
        ErrorRows        = nErrRows,
        ErrorColumns     = List.Count(List.Select(ColList, each _[ErrorCount] > 0)),
        MissingRequired  = Text.Combine(MissReq, ", "),
        MissingOptional  = Text.Combine(MissOpt, ", "),
        EmptyColumns     = Text.Combine(EmptyCols, ", "),
        DuplicateHeaders = Text.Combine(DupBases, ", ")
    ]
in
    [Summary = SummaryRec, Columns = ColTable, Errors = ErrTable]