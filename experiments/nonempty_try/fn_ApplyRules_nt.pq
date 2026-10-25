// ЕКСПЕРИМЕНТ (п. 35 документа): копія fn_ApplyRules з NonEmpty без try. Лише для заміру профайлером;
// після заміру видалити з книги. У робочий код не переносити: без try клітинка-помилка ламає запит.
// fn_ApplyRules_nt: конвеєр без перевірок і політик (лише дані; лічильники дає fn_ApplyRulesEx_nt)
(raw as nullable table, rules as record, optional culture as nullable text) as table =>
    fn_ApplyRulesEx_nt(raw, rules, culture)[Data]