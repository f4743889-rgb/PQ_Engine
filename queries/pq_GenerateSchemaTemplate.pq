// pq_GenerateSchemaTemplate - Генератор схеми зі статусами
let
    Registry = fn_LoadConfig()[Registry],

    Active = fn_ActiveSources(Registry),
    
    GetSchema = (row as record) as record =>
        let
            Src     = fn_ResolveSource(row),
            Attempt = if Src[ErrorCode] <> null
                      then [HasError = true, Error = [Message = Src[ErrorMessage]]]
                      else try fn_ParseHeaders(Src, fn_Options(row)),
            TplCols = fn_TemplateColumns(row[TargetType]),
            NoPath  = Src[ErrorCode] = "ШляхНеЗадано"
        in
            if not Attempt[HasError] then
                [Origin = "File", Status = "OK", Message = null,
                 Columns = Attempt[Value], SourcePath = Src[Path]]
            else if TplCols <> null then
                // шляху немає = штатний режим шаблону (OK); шлях є, але файл недоступний = WARN
                [Origin = "Template", Status = if NoPath then "OK" else "WARN",
                 Message = Attempt[Error][Message], Columns = TplCols, SourcePath = Src[Path]]
            else
                [Origin = "None", Status = "ERROR",
                 Message = Attempt[Error][Message], Columns = {null}, SourcePath = Src[Path]],

    ToPairs = (cols as list) as list =>
        List.Transform(List.Positions(cols),
            (i) => [ColumnOrder = i + 1, SourceColumnName = cols{i}]),

    Added = Table.AddColumn(Active, "S", each GetSchema(_), type record),
    Open  = Table.ExpandRecordColumn(Added, "S",
                {"Origin", "Status", "Message", "Columns", "SourcePath"},
                {"SchemaOrigin", "SchemaStatus", "StatusMessage", "Columns", "SourcePath"}),
    Pairs = Table.AddColumn(Open, "Col", each ToPairs([Columns]), type list),
    Slim  = Table.SelectColumns(Pairs,
                {"SourceID", "TargetType", "PeriodType", "PeriodNumber", "IsRequired",
                 "SourcePath", "SchemaOrigin", "SchemaStatus", "StatusMessage", "Col"}),
    // IsRequired реєстру джерел = «файл обов'язковий»; щоб не плутати з IsRequired у схемі
    Ren   = Table.RenameColumns(Slim, {{"IsRequired", "SourceRequired"}}),
    Long  = Table.ExpandListColumn(Ren, "Col"),
    Rec   = Table.ExpandRecordColumn(Long, "Col", {"ColumnOrder", "SourceColumnName"}),
    Typed = Table.TransformColumnTypes(Rec,
                {{"ColumnOrder", Int64.Type}, {"SourceColumnName", type text}})
in
    Typed