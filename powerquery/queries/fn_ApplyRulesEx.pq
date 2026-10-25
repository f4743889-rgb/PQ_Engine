// fn_ApplyRulesEx: конвеєр без перевірок і політик; повертає дані й лічильники
// [Data = таблиця, Dropped = відсічено службових рядків (tblRowFilters), RulesSkipped = правил пропущено]
// Правила фільтрів надходять у rules[RowFilters] (fn_GetSchemaRules з параметром rowFilters).
(raw as nullable table, rules as record, optional culture as nullable text) as record =>
let
    Culture = culture ?? cfgSettings[Culture],

    // немає даних: порожня таблиця з тими ж колонками, далі той самий конвеєр
    Base0 = if raw = null then #table(rules[SourceCols], {}) else raw,

    // 0. службові рядки (tblRowFilters): за іменами колонок джерела, ДО вибору колонок схеми,
    //    тому правило може посилатися й на колонку, якої немає в схемі
    Filters = Record.FieldOrDefault(rules, "RowFilters", null),

    // fn_ApplyRowFilters буферизує таблицю, тому перед нею лишаємо лише потрібні колонки:
    // колонки схеми + колонки правил фільтрів (v4.9; раніше в пам'ять потрапляв файл на всю ширину).
    // Відсутні в файлі колонки пропускаються: їх далі обробляють SelectColumns і перевірки звіту.
    FilterCols = if Filters = null or not Table.HasColumns(Filters, "SourceColumnName") then {}
                 else List.Select(List.Transform(Table.Column(Filters, "SourceColumnName"),
                          each Text.Trim(Text.From(_ ?? ""))), each _ <> ""),
    Narrow  = if Filters = null or raw = null then Base0
              else Table.SelectColumns(Base0, List.Distinct(rules[SourceCols] & FilterCols), MissingField.Ignore),

    Flt     = if Filters = null or raw = null
              then [Data = Base0, Dropped = 0, RulesApplied = 0, RulesSkipped = 0]
              else fn_ApplyRowFilters(Narrow, Filters, rules[TargetType], Culture),
    Base    = Flt[Data],

    // 1. канонічна форма: усі колонки схеми, у порядку схеми; відсутні необов'язкові = null
    Shaped = Table.SelectColumns(Base, rules[SourceCols], MissingField.UseNull),

    // 2. прибрати повністю порожні рядки (порожній текст теж порожній).
    //    Клітинка-помилка (#N/A з xlsx) вважається непорожньою: інакше помилка в предикаті
    //    ламала весь запит (аудит 1.2, v4.10)
    //    Один try на рядок (аудит 2.2): помилка до першої непорожньої клітинки дає true, тобто рядок
    //    непорожній; клітинка-помилка не може потрапити у try усередині функції (test_AuditFixes 1.2)
    NonEmpty = Table.SelectRows(Shaped,
                   each try List.MatchesAny(Record.FieldValues(_),
                                (v) => v <> null and (not (v is text) or Text.Trim(v) <> ""))
                        otherwise true),

    // 3. до-типові дії (імена джерела), 4. перейменування, 5. типи, 6. після-типові дії
    Pre     = if List.IsEmpty(rules[PreSteps])    then NonEmpty
              else Table.TransformColumns(NonEmpty, rules[PreSteps]),
    Renamed = if List.IsEmpty(rules[RenamePairs]) then Pre
              else Table.RenameColumns(Pre, rules[RenamePairs]),
    Typed   = Table.TransformColumnTypes(Renamed, rules[TypePairs], Culture),
    Post    = if List.IsEmpty(rules[PostSteps])   then Typed
              else Table.TransformColumns(Typed, rules[PostSteps])
in
    [Data = Post, Dropped = Flt[Dropped], RulesSkipped = Flt[RulesSkipped]]