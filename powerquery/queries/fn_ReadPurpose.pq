// fn_ReadPurpose - читання й об'єднання файлів призначення
// config: запис fn_LoadConfig() (Registry, Schema, RowFilters). Запит призначення:
//   fn_ReadPurpose(fn_LoadConfig(), "Registry")
// Частина реєстру (тести, підмножина джерел): fn_ReadPurpose(Cfg & [Registry = Підмножина], "Registry")
// options (необов'язково): [AllowDuplicatePeriods = true, AllowDuplicateFiles = true]
// ErrorPolicy, ErrorThreshold і Culture задаються лише в реєстрі або cfgSettings (v4.10): options їх не
// змінюють, щоб звіт якості (QualityFlag, типізація) збігався з поведінкою завантаження.
// Перевірки реєстру (дублі SourceID, файлу, періоду) - fn_RegistryIssues, ті самі, що в pq_RegistryCheck.
(config as record, targetType as text, optional options as nullable record) as table =>
let
    Opt   = options ?? [],
    TT    = fn_Canon(cfgSettings[Tables][TargetTypes], "TargetType")(targetType) ?? "",   // правило регістру
    Rules = fn_GetSchemaRules(config[Schema], TT, Record.FieldOrDefault(config, "RowFilters", null)),

    // активні джерела цього призначення (порядок реєстру)
    Rows = List.Select(Table.ToRecords(fn_ActiveSources(config[Registry])),
               (r) => Text.From(Record.FieldOrDefault(r, "TargetType", null) ?? "") = TT),

    IdOf = (r as record) as text => Text.From(fn_Opt(r, "SourceID", "")),

    // блокуючі проблеми реєстру для цього призначення; повідомлення про найважливішу групу
    Issues = Table.SelectRows(fn_RegistryIssues(config[Registry], Opt),
                 each [Blocking] and List.Contains([Targets], TT)),
    Code   = if Table.IsEmpty(Issues) then null else Issues{0}[Code],
    Keys   = if Code = null then {} else Table.SelectRows(Issues, each [Code] = Code)[Key],
    Msg    = if Code = "ДублюванняSourceID" then
                 "Для '" & TT & "' є дубльовані SourceID: " & Text.Combine(Keys, "; ")
                     & ". Зробіть SourceID унікальними (у всьому реєстрі)."
             else if Code = "ДублюванняФайлу" then
                 "Для '" & TT & "' той самий файл (і лист) указано в кількох активних джерелах: "
                     & Text.Combine(Keys, "; ")
             else
                 "Для '" & TT & "' кілька активних джерел мають однаковий період: " & Text.Combine(Keys, "; "),

    // метадані джерела (після Enforcer, щоб не впливати на схему); імена - cfgSettings[MetaColumns]
    Meta    = cfgSettings[MetaColumns],
    AddMeta = (t as table, id as nullable text, pt as nullable text, pn as nullable number) as table =>
        Table.AddColumn(
            Table.AddColumn(
                Table.AddColumn(t, Meta[Id], each id, type text),
                Meta[PeriodType], each pt, type text),
            Meta[PeriodNumber], each pn, Int64.Type),

    // одне джерело: читання → Enforcer → метадані
    Part = (r as record) as table =>
        let
            Pt  = fn_Opt(r, "PeriodType", null),
            Pn  = fn_Opt(r, "PeriodNumber", null),
            Enf = fn_EnforceSchema(fn_ReadSource(r), Rules, fn_EffectiveSettings(r, Opt))
        in
            AddMeta(Enf, IdOf(r),
                    if Pt = null then null else Text.From(Pt),
                    if Pn = null then null else Number.From(Pn)),

    // помилка отримує префікс [SourceID]
    Safe = (r as record) as table =>
        let a = try Part(r)
        in  if a[HasError] then
                error Error.Record(a[Error][Reason],
                    "[" & IdOf(r) & "] " & a[Error][Message], a[Error][Detail])
            else a[Value]
in
    if Code <> null then
        error Error.Record(Code, Msg, Keys)
    else if List.IsEmpty(Rows) then
        // немає активних джерел: порожня таблиця канонічної форми
        AddMeta(fn_EnforceSchema(null, Rules, fn_EffectiveSettings(null, Opt)), null, null, null)
    else
        Table.Combine(List.Transform(Rows, Safe))