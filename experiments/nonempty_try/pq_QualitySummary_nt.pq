// ЕКСПЕРИМЕНТ (п. 35 документа): копія pq_QualitySummary з NonEmpty без try. Лише для заміру профайлером;
// після заміру видалити з книги. У робочий код не переносити: без try клітинка-помилка ламає запит.
let
    T = fn_SummaryTable(fn_QcReports_nt(fn_LoadConfig()))
in
    T
