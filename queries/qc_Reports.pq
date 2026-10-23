// qc_Reports: звіти якості всіх активних джерел (список записів [Summary, Columns, Errors]).
// Відкриває файли, тому сам читає книгу через fn_LoadConfig. Логіка - у fn_QcReports.
let
    Reports = fn_QcReports(fn_LoadConfig())
in
    Reports