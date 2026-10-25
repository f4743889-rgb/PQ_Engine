// fn_EffectiveSettings: ефективні ErrorPolicy, ErrorThreshold і Culture для джерела (аудит D7)
// ErrorPolicy, ErrorThreshold: колонка рядка реєстру -> cfgSettings. options НЕ враховуються (v4.10,
//   аудит 1.4): звіт якості (QualityFlag) і завантаження мають однакові налаштування, тож прогноз
//   не може розійтися з реальною поведінкою. Політику джерела задавайте в реєстрі.
// Culture: колонка Culture реєстру (необов'язкова) -> cfgSettings. options теж не враховуються (v4.10):
//   fn_RowReport їх не має, тож інакше типізація у звіті й у завантаженні могла б розійтися.
// Параметр options лишено для сумісності викликів; у розрахунку не використовується (v4.11).
// Порожня клітинка реєстру = значення рівнем нижче; "off" у ErrorThreshold вимикає поріг для джерела.
// Threshold обчислюється лише при зверненні (некоректне значення дає помилку НекоректнийПоріг).
(row as nullable record, optional options as nullable record) as record =>
let
    R = row ?? [],
    // рядок реєстру -> cfgSettings
    Pick = (name as text) as any =>
        let fromRow = fn_Opt(R, name, null)
        in  if fromRow <> null then fromRow else Record.FieldOrDefault(cfgSettings, name, null)
in
    [
        Policy    = Text.Proper(Text.Trim(Text.From(Pick("ErrorPolicy") ?? "Report"))),
        Threshold = fn_ToRate(Pick("ErrorThreshold")),
        Culture   = Text.From(Pick("Culture"))
    ]