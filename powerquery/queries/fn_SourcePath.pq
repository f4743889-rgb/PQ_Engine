// fn_SourcePath
(r as record) as text =>
let
    Txt  = (n as text) as text =>
              Text.Trim(Text.From(Record.FieldOrDefault(r, n, null) ?? "")),
    Full = Txt("FullPath"),
    Dir  = Txt("Path"),
    Name = Txt("FileName")
in
    if Full <> "" then Full
    else if Name = "" then ""
    else if Dir = "" then Name
    else if Text.EndsWith(Dir, "\") then Dir & Name
    else Dir & "\" & Name