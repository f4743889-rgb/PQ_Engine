// fn_LoadVerdict: єдине правило «чи зупиниться завантаження» (аудит D6)
// Ним користуються fn_EnforceSchema (зупиняє запит) і fn_RowReport (QualityFlag у звітах),
// тож прогноз у звіті не може розійтися з реальною поведінкою.
// Параметри-функції викликаються лише тоді, коли потрібні: підрахунок помилок = ще один прохід по даних.
//   policy          - Strict | Report | Lenient (уже нормалізована, fn_EffectiveSettings)
//   missingRequired - список відсутніх обов'язкових колонок
//   hasErrors()     - чи є рядки з помилками
//   rate()          - частка рядків з помилками
//   threshold()     - поріг (частка) або null
// Результат: [Stop, Reason (код помилки або null), Flag (OK | WARN | FAIL), Rate, Threshold]
(policy as text, missingRequired as list, hasErrors as function, rate as function,
 threshold as function) as record =>
let
    Reason =
        if not List.IsEmpty(missingRequired) then "ВідсутніОбов'язковіПоля"
        else if not List.Contains({"Strict", "Report", "Lenient"}, policy) then "НевідомаПолітика"
        else if policy = "Strict" and hasErrors() then "ПомилкиТипізації"
        else if policy <> "Strict" and threshold() <> null and rate() > threshold() then "ПеревищенийПоріг"
        else null
in
    [
        Stop      = Reason <> null,
        Reason    = Reason,
        Flag      = if Reason <> null then "FAIL" else if hasErrors() then "WARN" else "OK",
        Rate      = rate(),
        Threshold = threshold()
    ]