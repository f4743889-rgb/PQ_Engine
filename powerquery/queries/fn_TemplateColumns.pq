// fn_TemplateColumns - читання шаблонів
(targetType as nullable text) as nullable list =>
let
    T = try Excel.CurrentWorkbook(){[Name = cfgSettings[Tables][DefaultHeaders]]}[Content]
        otherwise #table({"TargetType", "SourceColumnName"}, {}),

    Norm = (v as any) as text => Text.Trim(Text.From(v ?? "")),
    Key  = Norm(targetType),

    M0 = Table.SelectRows(T, each Text.Lower(Norm([TargetType])) = Text.Lower(Key)),   // правило регістру
    // необов'язковий ColumnOrder: якщо колонку додати, порядок стане стійким до сортування таблиці
    M  = if Table.HasColumns(M0, "ColumnOrder") then Table.Sort(M0, {"ColumnOrder"}) else M0,

    Cols = List.Select(
               List.Transform(Table.Column(M, "SourceColumnName"), Norm),
               each _ <> "")
in
    if Key = "" or List.IsEmpty(Cols) then null else List.Distinct(Cols)