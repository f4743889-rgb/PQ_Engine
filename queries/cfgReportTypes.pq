// cfgReportTypes: єдиний опис структури таблиць звітів якості (аудит D4)
// Нова колонка звіту додається тут і там, де її обчислюють (fn_SchemaReport / fn_RowReport);
// pq_QualitySummary, pq_QualityColumns, pq_QualityErrors і порожні таблиці беруть типи звідси.
[
    Summary = type table [
        SourceID = nullable text, TargetType = nullable text, PeriodType = any, PeriodNumber = any,
        SourceRequired = nullable logical, Status = nullable text, Message = nullable text,
        RawRows = nullable Int64.Type, Rows = nullable Int64.Type, EmptyRowsRemoved = nullable Int64.Type,
        RowsDropped = nullable Int64.Type, FilterRulesSkipped = nullable Int64.Type,
        ErrorRows = nullable Int64.Type, ErrorRate = nullable Percentage.Type,
        ErrorThreshold = nullable Percentage.Type, QualityFlag = nullable text,
        ErrorColumns = nullable Int64.Type, MissingRequired = nullable text,
        MissingOptional = nullable text, EmptyColumns = nullable text, DuplicateHeaders = nullable text
    ],
    Columns = type table [
        SourceID = nullable text, TargetType = nullable text, Source = nullable text, Target = nullable text,
        Required = nullable logical, ExpectedType = nullable text, InSource = nullable logical,
        Rows = nullable Int64.Type, NullCount = nullable Int64.Type, ErrorCount = nullable Int64.Type,
        ErrorSamples = nullable text
    ],
    Errors = type table [
        SourceID = nullable text, TargetType = nullable text, Column = nullable text,
        RowNo = nullable Int64.Type, Reason = nullable text, Message = nullable text, Detail = nullable text
    ]
]