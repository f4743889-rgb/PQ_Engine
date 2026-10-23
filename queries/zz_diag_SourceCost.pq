// zz_diag_SourceCost: тимчасовий діагностичний запит (Engine не змінює; після вимірювання видалити)
// Час повного звіту якості (fn_RowReport, як у pq_QualitySummary) окремо для КОЖНОГО активного
// джерела, по черзі. Додатково: чи вдається буферизувати нормалізовані дані джерела (BufferOk).
// Якщо BufferOk = FALSE, fn_SchemaReport працює без буфера і кожна перевірка колонки
// розбирає файл заново - підозра на основну частину 200 с.
// Запит відкриває файли, тому сам читає книгу (fn_LoadConfig) і викликає лише функції.
let
    Cfg     = fn_LoadConfig(),
    Active  = fn_ActiveSources(Cfg[Registry]),
    Rows    = Table.ToRecords(Active),
    Reports = fn_QcReports(Cfg),                       // той самий список, що й qc_Reports
    Pairs   = List.Zip({Rows, Reports}),

    // примусово обчислити звіт повністю (як під час завантаження на аркуш)
    Force = (rep as record) as text =>
        let s = rep[Summary],
            c = Table.RowCount(rep[Columns]),
            e = Table.RowCount(rep[Errors])
        in  Text.Combine(List.Transform(Record.FieldValues(s), each Text.From(_) ?? "")) & Text.From(c + e),

    // чи вдається Table.Buffer нормалізованих даних (так само, як Indexed у fn_SchemaReport)
    BufferOk = (row as record) as nullable logical =>
        let
            tt  = Text.Trim(Text.From(row[TargetType] ?? "")),
            raw = try fn_ReadSource(row),
            rul = try fn_GetSchemaRules(Cfg[Schema], tt, Cfg[RowFilters]),
            res = if raw[HasError] or rul[HasError] or raw[Value] = null then null
                  else let ex = fn_ApplyRulesEx(raw[Value], rul[Value], fn_EffectiveSettings(row)[Culture]),
                           b  = try Table.RowCount(Table.Buffer(
                                    Table.AddIndexColumn(ex[Data], "__RowNo", 1, 1, Int64.Type)))
                       in  not b[HasError]
        in  res,

    // по черзі: кожен крок чекає позначку часу попереднього
    Step = (state as record, p as list) as record =>
        let
            t0  = state[T],
            f   = if t0 = null then null else Force(p{1}),
            t1  = if f = null then null else DateTime.LocalNow(),
            s   = if t0 = null then null else p{1}[Summary],   // через t0: позначка часу береться ДО звіту
            rec = [Idx = List.Count(state[Out]), SourceID = s[SourceID], TargetType = s[TargetType],
                   Status = s[Status], Rows = s[Rows], ErrorRows = s[ErrorRows],
                   Seconds = Number.Round(Duration.TotalSeconds(t1 - t0), 1)]
        in
            [T = t1, Out = state[Out] & {rec}],

    Final  = List.Accumulate(Pairs, [T = DateTime.LocalNow(), Out = {}], Step),
    // MissingField.UseNull вимагає, щоб усі колонки типу допускали null
    Typ    = type table [Idx = nullable Int64.Type, SourceID = nullable text, TargetType = nullable text,
                         Status = nullable text, Rows = nullable number, ErrorRows = nullable number,
                         Seconds = nullable number],
    // спершу всі виміри часу (буфер), і лише потім окрема перевірка буфера - щоб вона не потрапила в час
    Timing = Table.Buffer(Table.FromRecords(Final[Out], Typ, MissingField.UseNull)),
    WithBuf = Table.AddColumn(Timing, "BufferOk",
                  each if [Status] = "OK" and ([ErrorRows] ?? 0) = 0 then null else BufferOk(Rows{[Idx]}),
                  type nullable logical)
in
    Table.Sort(Table.RemoveColumns(WithBuf, {"Idx"}), {{"Seconds", Order.Descending}})