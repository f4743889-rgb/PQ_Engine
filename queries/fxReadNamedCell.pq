/*
fxReadNamedCell
Безпечне зчитування значення з іменованої клітинки Excel.
*/

(cellName as text) as nullable text =>
let
    Attempt = try Excel.CurrentWorkbook(){[Name = cellName]}[Content]{0}[Column1]
in
    if Attempt[HasError] then
        error Error.Record(
            "НемаІменованогоДіапазону",
            "У книзі не знайдено іменований діапазон «" & cellName & "».",
            Attempt[Error][Message]
        )
    else
        let
            TextVal = if Attempt[Value] = null then null else Text.Trim(Text.From(Attempt[Value]))
        in
            if TextVal = "" then null else TextVal