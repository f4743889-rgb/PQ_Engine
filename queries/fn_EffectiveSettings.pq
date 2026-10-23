// fn_EffectiveSettings: ефективні ErrorPolicy, ErrorThreshold і Culture для джерела (аудит D7)
// Пріоритет: колонка рядка реєстру -> options виклику fn_ReadPurpose -> cfgSettings.
// Порожня клітинка реєстру = значення рівнем нижче; "off" у ErrorThreshold вимикає поріг для джерела.
// Threshold обчислюється лише при зверненні (некоректне значення дає помилку НекоректнийПоріг).
(row as nullable record, optional options as nullable record) as record =>
let
    R = row ?? [],
    O = options ?? [],
    Pick = (name as text) as any =>
        let
            fromRow = fn_Opt(R, name, null),
            fromOpt = Record.FieldOrDefault(O, name, null)
        in
            if fromRow <> null then fromRow
            else if fromOpt <> null then fromOpt
            else Record.FieldOrDefault(cfgSettings, name, null)
in
    [
        Policy    = Text.Proper(Text.Trim(Text.From(Pick("ErrorPolicy") ?? "Report"))),
        Threshold = fn_ToRate(Pick("ErrorThreshold")),
        Culture   = Record.FieldOrDefault(O, "Culture", null) ?? cfgSettings[Culture]
    ]