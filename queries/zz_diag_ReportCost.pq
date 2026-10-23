// zz_diag_ReportCost: тимчасовий діагностичний запит (Engine не змінює; після вимірювання видалити)
// Розкладає час звіту якості для ОДНОГО джерела на етапи. Кожен етап розбирає файл заново,
// тому етапи можна порівнювати між собою:
//   1 Читання         - лише розбір CSV (Table.RowCount сирих даних)
//   2 Конвеєр         - розбір + фільтри + відбір колонок + типи + буфер + рядки з помилками
//                       (те, що fn_SchemaReport робить до перевірок по колонках)
//   3 Повний звіт     - fn_SchemaReport, обчислюється лише таблиця колонок
//   4 Звіт повністю   - fn_SchemaReport: усі поля зведення + таблиці колонок і помилок
//   5 fn_RowReport    - те саме, що для кожного джерела робить pq_QualitySummary
// Різниця 4 - 3 = зведення й таблиця помилок; 5 - 4 = обгортка fn_RowReport.
// Запит відкриває файли, тому сам читає книгу (fn_LoadConfig) і викликає лише функції.
let
    // >>> вкажіть SourceID найбільшого реального файла <<<
    Sid = "GO3",

    Cfg     = fn_LoadConfig(),
    Row     = Table.SelectRows(fn_ActiveSources(Cfg[Registry]), each Text.Trim(Text.From([SourceID] ?? "")) = Sid){0},
    Rules   = fn_GetSchemaRules(Cfg[Schema], Row[TargetType], Cfg[RowFilters]),
    Culture = fn_EffectiveSettings(Row)[Culture],
    // таблиця M не кешується: кожен етап, що переглядає Raw, розбирає файл заново
    Raw     = fn_ReadSource(Row),

    // Кожен крок залежить від попереднього, тож позначки часу йдуть строго по черзі
    T0 = DateTime.LocalNow(),
    S1 = if T0 = null then null else Table.RowCount(Raw),
    T1 = if S1 = null then null else DateTime.LocalNow(),
    S2 = if T1 = null then null else
            let Ex = fn_ApplyRulesEx(Raw, Rules, Culture),
                I  = Table.Buffer(Table.AddIndexColumn(Ex[Data], "__RowNo", 1, 1, Int64.Type))
            in  Table.RowCount(I) + 0 * Table.RowCount(Table.SelectRowsWithErrors(I)),
    T2 = if S2 = null then null else DateTime.LocalNow(),
    S3 = if T2 = null then null else
            let Rep = fn_SchemaReport(Raw, Rules, Sid, Culture),
                Sum = Rep[Summary]
            in  Table.RowCount(Table.Buffer(Rep[Columns])) + 0 * (Sum[Rows] ?? 0),
    T3 = if S3 = null then null else DateTime.LocalNow(),

    // примусово обчислити звіт повністю: усі поля зведення + таблиці колонок і помилок
    Force = (rep as record) as number =>
        let s = rep[Summary]
        in  Text.Length(Text.Combine(List.Transform(Record.FieldValues(s), each Text.From(_) ?? "")))
            + Table.RowCount(Table.Buffer(rep[Columns])) + Table.RowCount(Table.Buffer(rep[Errors])),
    S4 = if T3 = null then null else Force(fn_SchemaReport(Raw, Rules, Sid, Culture)),
    T4 = if S4 = null then null else DateTime.LocalNow(),
    S5 = if T4 = null then null else Force(fn_RowReport(Row, [HasError = false, Value = Rules])),
    T5 = if S5 = null then null else DateTime.LocalNow(),

    Sec = (a as datetime, b as datetime) as number => Number.Round(Duration.TotalSeconds(b - a), 1),
    FileCols   = List.Count(Table.ColumnNames(Raw)),
    SchemaCols = List.Count(Rules[Columns])
in
    #table(
        type table [Stage = text, Seconds = nullable number, Result = nullable number, Note = nullable text],
        {
            {"1 Читання (розбір CSV)",              Sec(T0, T1), S1, "рядків у файлі; колонок у файлі: " & Text.From(FileCols)},
            {"2 Конвеєр нормалізації",              Sec(T1, T2), S2, "рядків після фільтрів і порожніх"},
            {"3 Повний звіт (з перевірками колонок)", Sec(T2, T3), S3, "колонок схеми: " & Text.From(SchemaCols)},
            {"4 fn_SchemaReport повністю",          Sec(T3, T4), S4, "зведення (усі поля) + колонки + помилки"},
            {"5 fn_RowReport повністю",             Sec(T4, T5), S5, "як у pq_QualitySummary (доступ до файла, try, якість)"},
            {"Разом",                               Sec(T0, T5), null, Sid & " (" & Row[TargetType] & ")"}
        })