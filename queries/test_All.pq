// test_All: підсумок регресійних і модульних тестів Engine. Очікування: у всіх рядках Result = PASS.
// Збирає результати інших запитів, книгу не читає.
let
    Suite = (name as text, passes as list) as record =>
        [Suite  = name,
         Checks = List.Count(passes),
         Failed = List.Count(List.Select(passes, each _ <> true))],

    Safe = (name as text, get as function) as record =>
        let r = try Suite(name, get())
        in  if r[HasError]
            then [Suite = name, Checks = 0, Failed = 1, Error = r[Error][Message]]
            else r[Value] & [Error = null],

    Rows = {
        Safe("test_Summary",    () => test_Summary[Pass]),
        Safe("test_Load",       () => test_Load[Pass]),
        Safe("test_ToRate",     () => test_ToRate[Pass]),
        Safe("test_Canon",      () => test_Canon[Pass]),
        Safe("test_LoadVerdict", () => test_LoadVerdict[Pass]),
        Safe("test_RegistryIssues", () => test_RegistryIssues[Pass]),
        Safe("test_Extensions", () => test_Extensions[Pass]),
        Safe("test_NullCounts", () => test_NullCounts[Pass]),
        Safe("test_RowFilters", () => {test_RowFilters[OK]})
    },
    T = Table.FromRecords(Rows,
            type table [Suite = text, Checks = Int64.Type, Failed = Int64.Type, Error = nullable text])
in
    Table.AddColumn(T, "Result", each if [Failed] = 0 and [Checks] > 0 then "PASS" else "FAIL", type text)