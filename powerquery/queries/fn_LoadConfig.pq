// fn_LoadConfig: конфігурація Engine з книги одним записом (аудит D10)
// Імена таблиць - cfgSettings[Tables]. Викликати всередині запиту, що відкриває файли (q_..., pq_Quality...,
// test_Load) або перевіряє таблиці: читання книги лишається в тому самому запиті (як у fn_TemplateColumns),
// тому Formula.Firewall не виникає. Поля запису обчислюються лише при зверненні.
//   Запит призначення:  fn_ReadPurpose(fn_LoadConfig(), "Registry")
() as record =>
let
    N   = cfgSettings[Tables],
    Get = (key as text) as table => Excel.CurrentWorkbook(){[Name = Record.Field(N, key)]}[Content],
    FiltersRaw = try Get("RowFilters")
                 otherwise #table({"TargetType", "SourceColumnName", "Operator", "Value", "Use"}, {})
in
    [
        Registry      = Get("Registry"),
        Schema        = fn_NormalizeSchema(Get("Schema")),
        RowFilters    = fn_NormalizeRowFilters(FiltersRaw),
        RowFiltersRaw = FiltersRaw
    ]