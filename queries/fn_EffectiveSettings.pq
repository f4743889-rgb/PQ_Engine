// fn_EffectiveSettings: ефективні ErrorPolicy, ErrorThreshold і Culture для джерела (аудит D7)
// ErrorPolicy, ErrorThreshold: колонка рядка реєстру -> cfgSettings. options НЕ враховуються (v4.10,
//   аудит 1.4): звіт якості (QualityFlag) і завантаження мають однакові налаштування, тож прогноз
//   не може розійтися з реальною поведінкою. Політику джерела задавайте в реєстрі.
// Culture: колонка Culture реєстру (необов'язкова) -> options -> cfgSettings.
// Порожня клітинка реєстру = значення рівнем нижче; "off" у ErrorThreshold вимикає поріг для джерела.
// Threshold обчислюється лише при зверненні (некоректне значення дає помилку НекоректнийПоріг).
(row as nullable record, optional options as nullable record) as record =>
let
    R = row ?? [],
    O = options ?? [],
    Pick = (name as text, useOptions as logical) as any =>
        let
            fromRow = fn_Opt(R, name, null),
            fromOpt = if useOptions then Record.FieldOrDefault(O, name, null) else null
        in
            if fromRow <> null then fromRow
            else if fromOpt <> null then fromOpt
            else Record.FieldOrDefault(cfgSettings, name, null)
in
    [
        Policy    = Text.Proper(Text.Trim(Text.From(Pick("ErrorPolicy", false) ?? "Report"))),
        Threshold = fn_ToRate(Pick("ErrorThreshold", false)),
        Culture   = Text.From(Pick("Culture", true))
    ]