// fn_Opt - значення з реєстру або дефолт; порожня клітинка = дефолт
(r as record, name as text, default as any) as any =>
let
    v = Record.FieldOrDefault(r, name, null)
in
    if v = null then default
    else if v is text then (if Text.Trim(v) = "" then default else Text.Trim(v))
    else v