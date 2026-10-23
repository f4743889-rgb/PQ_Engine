// fn_ApplyRules: конвеєр без перевірок і політик (лише дані; лічильники дає fn_ApplyRulesEx)
(raw as nullable table, rules as record, optional culture as nullable text) as table =>
    fn_ApplyRulesEx(raw, rules, culture)[Data]