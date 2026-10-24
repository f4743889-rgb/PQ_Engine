// fn_GetSchemaRules
// Схема має надходити з src_SchemaRegistry (нормалізовані текст і логічні поля, колонка Use).
// rowFilters (необов'язково): результат fn_NormalizeRowFilters; правила призначення потрапляють у Rules[RowFilters].
(schema as table, targetType as text, optional rowFilters as nullable table) as record =>
let
    Rows = Table.SelectRows(schema, each Text.Lower([TargetType]) = Text.Lower(Text.Trim(targetType))),
    Used = Table.SelectRows(Rows, each [Use]),

    // --- довідкові функції ---
    TrimClean = (v as any) as any => if v is text then Text.Trim(Text.Clean(v)) else v,

    Compose = (fs as list) as function =>
        (v as any) => List.Accumulate(fs, v, (s, f) => f(s)),

    // назви дій CustomTransform без урахування регістру (правило регістру)
    TrName = (n as text) as nullable text =>
        List.First(List.Select(Record.FieldNames(cfgTransforms),
            (f) => Text.Lower(f) = Text.Lower(n)), null),
    Tr = (n as text) as record =>
        if TrName(n) <> null then Record.Field(cfgTransforms, TrName(n))
        else error Error.Record("НевідомеПеретворення",
                 "CustomTransform '" & n & "' відсутнє в cfgTransforms.", n),

    Actions = (t as nullable text, stage as text) as list =>
        let
            Ns = List.Select(List.Transform(Text.Split(t ?? "", ";"), Text.Trim), each _ <> "")
        in
            List.Transform(List.Select(Ns, (n) => Tr(n)[Stage] = stage), (n) => Tr(n)[Fn]),

    ResolveType = (raw as nullable text) as any =>
        let key = if raw = null then "text" else fn_TypeKey(raw)
        in  if Record.HasFields(cfgTypeMap, key) then Record.Field(cfgTypeMap, key)
            else error Error.Record("НевідомийТипДаних",
                     "DataType '" & raw & "' не підтримується.", raw),

    // --- правило однієї колонки ---
    Col = (r as record) as record =>
        let
            PreF  = (if r[TrimFlag] then {TrimClean} else {}) & Actions(r[CustomTransform], "Pre"),
            PostF = Actions(r[CustomTransform], "Post")
        in
            [Source   = r[SourceColumnName],
             Target   = r[TargetColumnName] ?? r[SourceColumnName],
             Required = r[IsRequired],
             Type     = ResolveType(r[DataType]),
             Pre      = if List.IsEmpty(PreF)  then null else Compose(PreF),
             Post     = if List.IsEmpty(PostF) then null else Compose(PostF),
             TypeKey = if r[DataType] = null then "text" else fn_TypeKey(r[DataType])],

    Cols = List.Transform(Table.ToRecords(Used), Col),

    TargetCols = List.Transform(Cols, each _[Target]),
    SourceCols = List.Transform(Cols, each _[Source]),

    // поля Type, Pre, Post обчислюються ліниво: без примусового обчислення невідомий DataType чи
    // CustomTransform «вибухав» уже під час читання файла (READ_ERROR замість NO_SCHEMA; аудит 1.1, v4.10)
    Checked = List.Count(List.Select(Cols, (c) =>
                  c[Type] <> null and (c[Pre] = null or c[Pre] is function)
                  and (c[Post] = null or c[Post] is function))),

    // цільова назва не може збігатися з колонкою метаданих (fn_ReadPurpose додає їх після Enforcer)
    MetaClash = List.Intersect({TargetCols, cfgSettings[MetaColumns]}),

    Rules = [
        TargetType     = targetType,
        Columns        = Cols,
        SourceCols     = SourceCols,                 // для Table.SelectColumns
        TargetCols     = TargetCols,                 // канонічний порядок результату
        RequiredSource = List.Transform(List.Select(Cols, each _[Required]), each _[Source]),
        PreSteps       = List.Transform(List.Select(Cols, each _[Pre] <> null),
                             each {_[Source], _[Pre]}),      // для Table.TransformColumns (до перейменування)
        RenamePairs    = List.Transform(List.Select(Cols, each _[Source] <> _[Target]),
                             each {_[Source], _[Target]}),   // для Table.RenameColumns
        TypePairs      = List.Transform(Cols, each {_[Target], _[Type]}),   // для Table.TransformColumnTypes
        PostSteps = List.Transform(List.Select(Cols, each _[Post] <> null),
                        each {_[Target], _[Post], _[Type]}),     // для Table.TransformColumns (після типів)
        RowFilters = if rowFilters = null then null
                     else Table.SelectRows(rowFilters,
                              each Text.Lower([TargetType]) = Text.Lower(targetType) and [Use])
    ]
in
    if Table.IsEmpty(Rows) then
        error Error.Record("СхемаНеЗнайдена",
            "У tblSchemaRegistry немає правил для TargetType '" & targetType & "'.", targetType)
    else if Table.IsEmpty(Used) then
        error Error.Record("НемаВикористовуванихКолонок",
            "Для TargetType '" & targetType & "' усі колонки мають Use = Ні.", targetType)
    else if List.Count(List.Distinct(TargetCols)) <> List.Count(TargetCols)
         or List.Count(List.Distinct(SourceCols)) <> List.Count(SourceCols) then
        error Error.Record("ДублікатиВСхемі",
            "У схемі '" & targetType & "' є дублі назв. Див. pq_SchemaCheck.", targetType)
    else if not List.IsEmpty(MetaClash) then
        error Error.Record("ЗбігЗМетаколонкою",
            "У схемі '" & targetType & "' цільові назви збігаються з колонками метаданих: "
                & Text.Combine(MetaClash, ", ") & ". Див. pq_SchemaCheck.", MetaClash)
    else if Checked <> List.Count(Cols) then
        error Error.Record("НекоректніПравила", "Правила схеми '" & targetType & "' не скомпільовано.", targetType)
    else
        Rules