// pq_RegistryCheck: перевірка реєстру джерел. Усі правила - fn_RegistryIssues, ті самі, що зупиняють
// fn_ReadPurpose. Blocking = TRUE: завантаження призначень з колонки TargetTypes зупиниться
// (якщо в options виклику не дозволено AllowDuplicatePeriods / AllowDuplicateFiles).
let
    I = fn_RegistryIssues(fn_LoadConfig()[Registry])
in
    Table.SelectColumns(
        Table.AddColumn(I, "TargetTypes", each Text.Combine([Targets], ", "), type text),
        {"Issue", "Key", "Count", "TargetTypes", "Blocking"})