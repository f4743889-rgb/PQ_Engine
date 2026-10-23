// fn_ReadSource - Головна функція завантаження даних
(sourceRecord as record, optional template as nullable table) as nullable table =>
let
    IsRequired = fn_ToBool(Record.FieldOrDefault(sourceRecord, "IsRequired", null), true),
    Src        = fn_ResolveSource(sourceRecord),
    Opts       = fn_Options(sourceRecord)
in
    if Src[ErrorCode] <> null then
        if IsRequired
        then error Error.Record(Src[ErrorCode], Src[ErrorMessage], Src[ErrorDetail])
        else template
    else
        fn_ParseSource(Src, Opts)