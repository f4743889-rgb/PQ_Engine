// fn_ResolveSource - єдине місце, де відкривається файл; помилки як дані
(r as record) as record =>
let
    Path = fn_SourcePath(r),
    Ext  = Text.Lower(Text.AfterDelimiter(Path, ".", {0, RelativePosition.FromEnd}) ?? ""),

    Make = (content as nullable binary, code as nullable text, msg as nullable text) as record =>
              [Path = Path, Ext = Ext, Content = content,
               ErrorCode = code, ErrorMessage = msg, ErrorDetail = Path],
    // один виклик File.Contents (рецензія коду 25.10.2026, CR-03): доступність перевіряється
    // читанням першого байта того самого значення; помилка доступу повертається як дані
    Bin     = if Path = "" then null else File.Contents(Path),
    Probe   = if Path = "" then null
              else try List.Count(Binary.ToList(Binary.Range(Bin, 0, 1))),
    Attempt = if Path = "" then null
              else if Probe[HasError] then Probe
              else [HasError = false, Value = Bin]
in
    if Path = "" then Make(null, "ШляхНеЗадано", "Шлях до файла не задано.")
    else if Attempt[HasError] then Make(null, "ФайлНедоступний", Attempt[Error][Message])
    else Make(Attempt[Value], null, null)