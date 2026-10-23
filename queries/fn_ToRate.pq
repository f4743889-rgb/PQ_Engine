// fn_ToRate: значення порогу ErrorThreshold -> частка від 0 до 1 або null (поріг вимкнено)
// Приймає: клітинку з відсотковим форматом (5% Excel передає як 0.05) | число 0.05 |
//   текст "0,05" / "0.05" (частка) | текст "5%" (відсотки, лише зі знаком %).
//   Число або текст без % завжди частка від 0 до 1: 5, 50 чи 5000% (=50) дають помилку,
//   щоб помилково введене значення не перетворювалось тихо на інший поріг.
//   null, "", "off", "ні", "no", "-", "false", FALSE = вимкнено.
// Інакше помилка НекоректнийПоріг (її показує pq_RegistryCheck).
(v as any) as nullable number =>
let
    Off   = {"", "off", "ні", "no", "-", "false"},
    Txt   = if v is text then Text.Lower(Text.Trim(v)) else null,
    IsPct = Txt <> null and Text.EndsWith(Txt, "%"),

    Bad = (msg as text) => error Error.Record("НекоректнийПоріг", msg, v),

    Num =
        if v = null then null
        else if v is logical then (if v then Bad("ErrorThreshold = TRUE не є порогом.") else null)
        else if v is number then v
        else if Txt = null then Bad("ErrorThreshold має бути числом або відсотком.")
        else if List.Contains(Off, Txt) then null
        else
            let
                t = Text.Trim(Text.Replace(Txt, "%", "")),
                n = try Number.FromText(t, cfgSettings[Culture])
                    otherwise try Number.FromText(t, "en-US") otherwise null
            in
                if n = null then Bad("ErrorThreshold '" & v & "' не є числом.") else n,

    Rate = if Num = null then null else if IsPct then Num / 100 else Num
in
    if Rate = null then null
    else if Rate < 0 or Rate > 1
        then Bad("ErrorThreshold має бути від 0 до 100%, отримано " & Text.From(v) & ".")
    else Rate