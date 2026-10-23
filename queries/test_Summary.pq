// test_Summary: регресійна перевірка зведення якості за таблицею tblTestExpected
// Один рядок = одна перевірка (джерело + поле). Порожнє очікування не перевіряється,
// "-" = очікується порожнє значення, Status = ABSENT = джерела не має бути у зведенні.
// Текст порівнюється без урахування регістру й пробілів на краях, числа як текст.
// Звіти будуються ЛИШЕ для SourceID з tblTestExpected (fn_QcReports з фільтром), тому реальні
// джерела в реєстрі тестової книги не сповільнюють тести. Запит відкриває файли, тож сам читає
// книгу (tblTestExpected, fn_LoadConfig) і викликає лише функції (правило трьох видів запитів).
let
    Exp    = Excel.CurrentWorkbook(){[Name = "tblTestExpected"]}[Content],
    Ids    = List.Distinct(List.Select(
                 List.Transform(Table.Column(Exp, "SourceID"), each Text.Trim(Text.From(_ ?? ""))),
                 each _ <> "")),
    Act    = Table.Buffer(fn_SummaryTable(fn_QcReports(fn_LoadConfig(), Ids))),
    Fields = {"Status", "QualityFlag", "Rows", "RowsDropped", "ErrorRows",
              "MissingRequired", "MissingOptional", "DuplicateHeaders"},

    Norm = (v as any) as text => if v = null then "" else Text.Trim(Text.From(v)),
    Same = (expected as text, actual as text) as logical =>
               Text.Lower(if expected = "-" then "" else expected) = Text.Lower(actual),

    Mk = (id as text, field as text, e as text, a as text, ok as logical) as record =>
            [SourceID = id, Field = field, Expected = e, Actual = a, Pass = ok],

    Check = (e as record) as list =>
        let
            id    = Norm(Record.FieldOrDefault(e, "SourceID", null)),
            rows  = Table.SelectRows(Act, each Norm([SourceID]) = id),
            found = not Table.IsEmpty(rows),
            a     = if found then rows{0} else [],
            st    = Norm(Record.FieldOrDefault(e, "Status", null))
        in
            if id = "" then {}
            else if Text.Upper(st) = "ABSENT" then
                {Mk(id, "(рядок у зведенні)", "відсутній", if found then "присутній" else "відсутній", not found)}
            else if not found then
                {Mk(id, "(рядок у зведенні)", "присутній", "відсутній", false)}
            else
                List.Transform(
                    List.Select(Fields, (f) => Norm(Record.FieldOrDefault(e, f, null)) <> ""),
                    (f) => let ev = Norm(Record.Field(e, f)), av = Norm(Record.FieldOrDefault(a, f, null))
                           in  Mk(id, f, ev, av, Same(ev, av))),

    All = List.Combine(List.Transform(Table.ToRecords(Exp), Check)),
    Typ = type table [SourceID = text, Field = text, Expected = text, Actual = text, Pass = logical]
in
    if List.IsEmpty(All) then #table(Typ, {})
    else Table.Sort(Table.FromRecords(All, Typ), {{"Pass", Order.Ascending}, {"SourceID", Order.Ascending}})