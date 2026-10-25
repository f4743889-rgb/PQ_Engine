// test_Load: регресійна перевірка завантаження fn_ReadPurpose за таблицею tblTestLoad
// Для кожного сценарію викликає fn_ReadPurpose лише для вказаних SourceID (через кому)
// і порівнює: Outcome (OK / ERROR), Reason (код помилки), кількість рядків і клітинок з помилками.
// Порожнє очікування не перевіряється. Запит відкриває файли, тому читає книгу сам
// і викликає лише функції (правило трьох видів запитів).
let
    Cfg    = fn_LoadConfig(),
    Reg    = Cfg[Registry],
    Cases  = Excel.CurrentWorkbook(){[Name = "tblTestLoad"]}[Content],

    Norm = (v as any) as text => if v = null then "" else Text.Trim(Text.From(v)),
    Same = (expected as any, actual as any) as logical =>
               Norm(expected) = "" or Text.Lower(Norm(expected)) = Text.Lower(Norm(actual)),

    Run = (c as record) as record =>
        let
            ids  = List.Select(List.Transform(Text.Split(Norm(c[SourceIDs]), ","), Text.Trim), each _ <> ""),
            sub  = Table.SelectRows(Reg, each List.Contains(ids, Norm([SourceID]))),
            tt   = Norm(c[TargetType]),
            Load = () => fn_ReadPurpose(Cfg & [Registry = sub], tt),

            // одне завантаження на сценарій (v4.10; раніше Load() виконувався двічі): буфер зберігає
            // клітинки-помилки (Report), а помилки рівня таблиці (Strict, поріг, реєстр) ловить try
            Cnt     = try Table.Buffer(Load()),
            Outcome = if Cnt[HasError] then "ERROR" else "OK",
            Reason  = if Cnt[HasError] then Norm(Cnt[Error][Reason]) else "",
            Rows    = if Cnt[HasError] then null else Table.RowCount(Cnt[Value]),
            // клітинки з помилками рахуються за один перегляд таблиці (раніше - окремий
            // перегляд, тобто повне завантаження, на кожну колонку) (патч F2)
            ErrCells =
                if Cnt[HasError] then null
                else let T    = Cnt[Value],
                         Cols = Table.ColumnNames(T),
                         PerRow = Table.AddColumn(T, "__ErrCells",
                             (r) => List.Count(List.Select(Cols, (c) => (try Record.Field(r, c))[HasError])),
                             Int64.Type)
                     in  List.Sum({0} & Table.Column(PerRow, "__ErrCells"))
        in
            [Case          = Norm(c[Case]),
             TargetType    = tt,
             SourceIDs     = Norm(c[SourceIDs]),
             ExpOutcome    = Norm(c[ExpOutcome]),
             Outcome       = Outcome,
             ExpReason     = Norm(c[ExpReason]),
             Reason        = Reason,
             ExpRows       = c[ExpRows],
             Rows          = Rows,
             ExpErrorCells = c[ExpErrorCells],
             ErrorCells    = ErrCells,
             Message       = if Cnt[HasError] then Cnt[Error][Message] else null,
             Pass          = Same(c[ExpOutcome], Outcome) and Same(c[ExpReason], Reason)
                             and Same(c[ExpRows], Rows) and Same(c[ExpErrorCells], ErrCells)],

    Valid = List.Select(Table.ToRecords(Cases), each Norm([Case]) <> ""),
    Typ   = type table [Case = text, TargetType = text, SourceIDs = text,
                        ExpOutcome = text, Outcome = text, ExpReason = text, Reason = text,
                        ExpRows = nullable number, Rows = nullable number,
                        ExpErrorCells = nullable number, ErrorCells = nullable number,
                        Message = nullable text, Pass = logical]
in
    if List.IsEmpty(Valid) then #table(Typ, {})
    else Table.FromRecords(List.Transform(Valid, Run), Typ)
