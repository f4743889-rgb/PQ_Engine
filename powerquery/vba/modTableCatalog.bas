Attribute VB_Name = "modTableCatalog"
Option Explicit

' ============================================================================
' modTableCatalog: каталог таблиць книги, підписи над таблицями і довідка (v4.11e)
' Потребує modEngineCommon (engFindTable, engColumnExists, engCell, engCellOpt, engQueryNameOf,
' engTableName). Потрібен Excel 2021+: формули підписів використовують LET, XLOOKUP, IFS.
'
' UpdateHelp за один прохід:
'   1. Аркуш HELP і таблиця tblTableCatalog на ньому (створюються, якщо їх немає).
'   2. На кожному аркуші, де шапка якоїсь таблиці вище рядка 3, вставляє рядки зверху.
'   3. Якщо таблиці hlpProject немає, створює її на HELP зі стартовими правилами структури книги.
'   4. Оновлює каталог: №, Таблиця (гіперпосилання), Аркуш, Діапазон, Рядків, Колонок, Тип, Запит,
'      У моделі даних, Призначення, Підпис, Довідка, Порушення, Колонки таблиці.
'      Призначення, Підпис (TRUE/FALSE) і Довідка заповнюєте ви - повторний запуск їх зберігає.
'      Порушення - відхилення від правил hlpProject; макрос їх лише показує.
'      Таблиця і Довідка - гіперпосилання на таблицю (Довідка - якщо hlp-таблицю знайдено).
'   5. Над кожною таблицею з Підпис = TRUE ставить підпис у рядку (шапка - 2), у перших трьох
'      колонках: назва (з приміткою «Назва таблиці») | «i Призначення» | «? Як заповнювати»
'      (для hlp-таблиці - «<- до tblX»).
'      Клітинки підпису запам'ятовуються прихованими іменами engLbl_<таблиця>; на наступному
'      запуску вони очищаються і пишуться заново. Чужий вміст у рядку підпису не перезаписується
'      (порушення «рядок підпису зайнятий»).
' Чужі таблиці макрос не пересуває; єдина автоматична зміна розміщення - вставка рядків зверху.
' ============================================================================

Private Const CAT_SHEET As String = "HELP"
Private Const CAT_TABLE As String = "tblTableCatalog"
Private Const PROJECT_HELP As String = "hlpProject"
Private Const LABEL_NAME_PREFIX As String = "engLbl_"
Private Const LABEL_NOTE As String = "Назва таблиці"
Private Const HEADER_ROW_MIN As Long = 3
Private Const LABEL_COLS As Long = 3
Private Const COL_COUNT As Long = 14

' колонки каталогу
Private Const C_NUM As Long = 1, C_NAME As Long = 2, C_SHEET As Long = 3, C_RANGE As Long = 4
Private Const C_ROWS As Long = 5, C_COLS As Long = 6, C_KIND As Long = 7, C_QUERY As Long = 8
Private Const C_MODEL As Long = 9, C_DESC As Long = 10, C_LABEL As Long = 11, C_HELP As Long = 12
Private Const C_ISSUES As Long = 13, C_COLLIST As Long = 14

Private Function CatalogHeader() As Variant
    CatalogHeader = Array("№", "Таблиця", "Аркуш", "Діапазон", "Рядків", "Колонок", "Тип", "Запит", _
                          "У моделі даних", "Призначення", "Підпис", "Довідка", "Порушення", "Колонки таблиці")
End Function

Public Sub UpdateHelp()
    Dim wsCat As Worksheet, cat As ListObject, kept As Object, lo As ListObject
    Dim tbls() As ListObject, n As Long, i As Long, data() As Variant, lblOk() As Boolean
    Dim oldScreen As Boolean, ok As Boolean, lbl As Boolean, rec As Variant, lblV As Variant
    Dim nm As String, key As String, qName As String, desc As String, helpName As String, iss As String
    Dim inserted As Long, failed As String, created As Boolean
    Dim labeled As Long, nIssues As Long, filled As Long, removed As Long

    oldScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Fail

    Set cat = EnsureCatalog(wsCat)
    Set kept = ReadKept(cat)
    ClearAllLabels
    inserted = EnsureHeaderRows(failed)
    EnsureCatalogColumns cat
    created = EnsureProjectHelp(wsCat)

    n = CollectTables(tbls)
    If n > 0 Then
        ReDim data(1 To n, 1 To COL_COUNT)
        ReDim lblOk(1 To n)
        For i = 1 To n
            Set lo = tbls(i)
            nm = lo.Name: key = LCase$(nm)
            qName = QueryOf(lo)
            desc = "": lblV = Empty: helpName = ""
            If kept.Exists(key) Then
                rec = kept(key)
                desc = rec(0): lblV = rec(1): helpName = rec(2)
                kept.Remove key
            End If
            If desc = "" Then
                desc = DefaultDescription(nm, qName)
                If desc <> "" Then filled = filled + 1
            End If
            lbl = ToBoolLabel(lblV, DefaultLabel(nm))
            If helpName = "" Then helpName = DefaultHelp(nm)
            iss = TableIssues(lo, lbl, helpName, desc, ok)
            lblOk(i) = ok
            If iss <> "" Then nIssues = nIssues + 1

            data(i, C_NUM) = i
            data(i, C_NAME) = nm
            data(i, C_SHEET) = lo.Parent.Name
            data(i, C_RANGE) = lo.Range.Address(False, False)
            data(i, C_ROWS) = lo.ListRows.Count
            data(i, C_COLS) = lo.ListColumns.Count
            data(i, C_KIND) = KindOf(lo)
            data(i, C_QUERY) = qName
            data(i, C_MODEL) = IIf(InModel(lo), "так", "")
            data(i, C_DESC) = desc
            data(i, C_LABEL) = lbl
            data(i, C_HELP) = helpName
            data(i, C_ISSUES) = iss
            data(i, C_COLLIST) = ColumnList(lo)
        Next i
    End If
    removed = kept.Count            ' таблиці з каталогу, яких у книзі вже немає

    WriteCatalog cat, wsCat, data, n

    For i = 1 To n
        If data(i, C_LABEL) And lblOk(i) Then
            WriteLabel tbls(i), wsCat
            labeled = labeled + 1
        End If
    Next i

    Application.ScreenUpdating = oldScreen
    MsgBox "Таблиць у каталозі: " & n & vbCrLf & _
           "Підписів над таблицями: " & labeled & vbCrLf & _
           "Таблиць з порушеннями: " & nIssues & IIf(nIssues > 0, " (колонка «Порушення»)", "") & vbCrLf & _
           IIf(inserted > 0, "Аркушів, де вставлено рядки над таблицями: " & inserted & vbCrLf, "") & _
           IIf(failed <> "", "Не вдалося вставити рядки на аркушах: " & failed & vbCrLf, "") & _
           IIf(created, "Створено таблицю " & PROJECT_HELP & " зі стартовими правилами." & vbCrLf, "") & _
           "Стартових описів додано: " & filled & vbCrLf & _
           IIf(removed > 0, "Видалено з каталогу (таблиць більше немає): " & removed & vbCrLf, "") & vbCrLf & _
           "Аркуш «" & wsCat.Name & "», таблиця " & CAT_TABLE & "." & vbCrLf & _
           "Колонки «Призначення», «Підпис», «Довідка» можна редагувати: повторний запуск їх зберігає.", _
           vbInformation, "Довідка і каталог таблиць"
    Exit Sub
Fail:
    Application.ScreenUpdating = oldScreen
    MsgBox "Помилка: " & Err.Description, vbCritical, "Довідка і каталог таблиць"
End Sub

' ================== Аркуш HELP і таблиця каталогу ==================
Private Function SheetExists(nm As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then SheetExists = True: Exit Function
    Next ws
End Function

Private Function EnsureCatalog(ByRef wsCat As Worksheet) As ListObject
    Dim lo As ListObject
    Set lo = engFindTable(CAT_TABLE)
    If Not lo Is Nothing Then
        Set wsCat = lo.Parent
        Set EnsureCatalog = lo
        Exit Function
    End If

    If SheetExists(CAT_SHEET) Then
        Set wsCat = ThisWorkbook.Worksheets(CAT_SHEET)
    Else
        Set wsCat = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        wsCat.Name = CAT_SHEET
    End If

    wsCat.Cells(HEADER_ROW_MIN, 1).Resize(1, COL_COUNT).value = CatalogHeader()
    Set lo = wsCat.ListObjects.Add(xlSrcRange, wsCat.Cells(HEADER_ROW_MIN, 1).Resize(2, COL_COUNT), , xlYes)
    lo.Name = CAT_TABLE
    Set EnsureCatalog = lo
End Function

' назва таблиці (нижній регістр) -> Array(Призначення, Підпис, Довідка) з наявного каталогу
Private Function ReadKept(cat As ListObject) As Object
    Dim d As Object, i As Long, nm As String
    Set d = CreateObject("Scripting.Dictionary")
    Set ReadKept = d
    If cat.DataBodyRange Is Nothing Then Exit Function
    If Not engColumnExists(cat, "Таблиця") Then Exit Function
    For i = 1 To cat.ListRows.Count
        nm = Trim$(CStr(engCell(cat, i, "Таблиця")))
        If nm <> "" Then
            d(LCase$(nm)) = Array(Trim$(CStr(engCellOpt(cat, i, "Призначення", ""))), _
                                  engCellOpt(cat, i, "Підпис", Empty), _
                                  Trim$(CStr(engCellOpt(cat, i, "Довідка", ""))))
        End If
    Next i
End Function

' Приводить каталог до 14 колонок у порядку CatalogHeader. Дані перезаписує WriteCatalog.
Private Sub EnsureCatalogColumns(cat As ListObject)
    Dim hdr As Variant, i As Long, cur As Long, same As Boolean, extra As Range, lo As ListObject
    Dim ws As Worksheet
    hdr = CatalogHeader()
    cur = cat.ListColumns.Count
    If cur = COL_COUNT Then
        same = True
        For i = 1 To COL_COUNT
            If cat.ListColumns(i).Name <> hdr(i - 1) Then same = False: Exit For
        Next i
        If same Then Exit Sub
    End If

    Set ws = cat.Parent
    If cur < COL_COUNT Then
        Set extra = cat.Range.Cells(1, cur + 1).Resize(cat.Range.Rows.Count, COL_COUNT - cur)
        If Application.WorksheetFunction.CountA(extra) > 0 Then RaiseNoRoom extra
        For Each lo In ws.ListObjects
            If lo.Name <> cat.Name Then
                If Not Intersect(lo.Range.EntireColumn, extra.EntireColumn) Is Nothing Then RaiseNoRoom extra
            End If
        Next lo
    End If

    If Not cat.DataBodyRange Is Nothing Then
        cat.DataBodyRange.Hyperlinks.Delete
        cat.DataBodyRange.ClearContents
    End If
    cat.Resize cat.Range.Cells(1, 1).Resize(cat.Range.Rows.Count, COL_COUNT)
    If cur > COL_COUNT Then
        ws.Cells(cat.Range.Row, cat.Range.Column + COL_COUNT).Resize(cat.Range.Rows.Count, cur - COL_COUNT).Clear
    End If
    ' тимчасові назви, щоб нові назви не збіглися зі старими в іншій позиції
    For i = 1 To COL_COUNT
        cat.HeaderRowRange.Cells(1, i).value = "tmp_col_" & i
    Next i
    cat.HeaderRowRange.value = hdr
End Sub

Private Sub RaiseNoRoom(extra As Range)
    Err.Raise vbObjectError + 513, , "Праворуч від каталогу " & CAT_TABLE & " немає місця для нових колонок (" & _
              extra.EntireColumn.Address(False, False) & "). Звільніть ці колонки або перенесіть сусідню " & _
              "таблицю і запустіть UpdateHelp ще раз."
End Sub

Private Sub WriteCatalog(cat As ListObject, wsCat As Worksheet, data() As Variant, n As Long)
    Dim body As Range, i As Long
    If Not cat.DataBodyRange Is Nothing Then
        cat.DataBodyRange.Hyperlinks.Delete
        cat.DataBodyRange.ClearContents
    End If
    cat.Resize cat.HeaderRowRange.Resize(Application.Max(n, 1) + 1)
    If n = 0 Then Exit Sub

    Set body = cat.DataBodyRange
    body.value = data

    ' гіперпосилання: назва таблиці і назва довідки (якщо таблиця знайдена) -> діапазон таблиці
    For i = 1 To n
        AddTableLink wsCat, body.Cells(i, C_NAME), CStr(data(i, C_NAME))
        If CStr(data(i, C_HELP)) <> "" Then AddTableLink wsCat, body.Cells(i, C_HELP), CStr(data(i, C_HELP))
    Next i

    With cat
        .DataBodyRange.VerticalAlignment = xlTop
        .ListColumns(C_DESC).DataBodyRange.WrapText = True
        .ListColumns(C_ISSUES).DataBodyRange.WrapText = True
        .ListColumns(C_COLLIST).DataBodyRange.WrapText = False
        For i = C_NUM To C_MODEL
            .ListColumns(i).Range.EntireColumn.AutoFit
        Next i
        .ListColumns(C_LABEL).Range.EntireColumn.AutoFit
        .ListColumns(C_HELP).Range.EntireColumn.AutoFit
        .ListColumns(C_DESC).Range.EntireColumn.ColumnWidth = 60
        .ListColumns(C_ISSUES).Range.EntireColumn.ColumnWidth = 50
        .ListColumns(C_COLLIST).Range.EntireColumn.ColumnWidth = 60
    End With
End Sub

' Гіперпосилання в клітинці каталогу на таблицю tblName; якщо таблиці немає - лише текст
Private Sub AddTableLink(wsCat As Worksheet, anchor As Range, tblName As String)
    Dim lo As ListObject, target As String
    Set lo = engFindTable(tblName)
    If lo Is Nothing Then Exit Sub
    target = "'" & Replace(lo.Parent.Name, "'", "''") & "'!" & lo.Range.Address
    wsCat.Hyperlinks.Add anchor:=anchor, Address:="", SubAddress:=target, TextToDisplay:=tblName
End Sub

Private Function CollectTables(ByRef tbls() As ListObject) As Long
    Dim ws As Worksheet, lo As ListObject, n As Long
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name <> CAT_TABLE Then
                n = n + 1
                ReDim Preserve tbls(1 To n)
                Set tbls(n) = lo
            End If
        Next lo
    Next ws
    CollectTables = n
End Function

' ================== Розміщення: шапки не вище рядка 3 ==================
' Повертає кількість аркушів, де вставлено рядки; failed - аркуші, де вставити не вдалося.
Private Function EnsureHeaderRows(ByRef failed As String) As Long
    Dim ws As Worksheet, lo As ListObject, minRow As Long, k As Long
    For Each ws In ThisWorkbook.Worksheets
        minRow = 0
        For Each lo In ws.ListObjects
            If minRow = 0 Or lo.Range.Row < minRow Then minRow = lo.Range.Row
        Next lo
        If minRow > 0 And minRow < HEADER_ROW_MIN Then
            k = HEADER_ROW_MIN - minRow
            On Error Resume Next
            ws.Rows("1:" & k).Insert Shift:=xlDown
            If Err.Number <> 0 Then
                failed = failed & IIf(failed = "", "", ", ") & ws.Name
                Err.Clear
            Else
                EnsureHeaderRows = EnsureHeaderRows + 1
            End If
            On Error GoTo 0
        End If
    Next ws
End Function

' ================== hlpProject: загальна довідка ==================
Private Function EnsureProjectHelp(wsCat As Worksheet) As Boolean
    Dim rules As Collection, col As Long, lo As ListObject, i As Long, arr() As Variant, top As Range, r As Variant
    If Not engFindTable(PROJECT_HELP) Is Nothing Then Exit Function

    ' перший вільний блок праворуч від усього на аркуші, через одну порожню колонку
    For Each lo In wsCat.ListObjects
        If lo.Range.Column + Application.Max(lo.Range.Columns.Count, LABEL_COLS) - 1 > col Then
            col = lo.Range.Column + Application.Max(lo.Range.Columns.Count, LABEL_COLS) - 1
        End If
    Next lo
    With wsCat.UsedRange
        If .Column + .Columns.Count - 1 > col Then col = .Column + .Columns.Count - 1
    End With
    col = col + 2

    Set rules = ProjectRules()
    ReDim arr(1 To rules.Count + 1, 1 To 3)
    arr(1, 1) = "№": arr(1, 2) = "Розділ": arr(1, 3) = "Правило"
    i = 1
    For Each r In rules
        i = i + 1
        arr(i, 1) = i - 1: arr(i, 2) = r(0): arr(i, 3) = r(1)
    Next r

    Set top = wsCat.Cells(HEADER_ROW_MIN, col)
    top.Resize(rules.Count + 1, 3).value = arr
    Set lo = wsCat.ListObjects.Add(xlSrcRange, top.Resize(rules.Count + 1, 3), , xlYes)
    lo.Name = PROJECT_HELP
    lo.ListColumns(1).Range.EntireColumn.ColumnWidth = 5
    lo.ListColumns(2).Range.EntireColumn.ColumnWidth = 16
    lo.ListColumns(3).Range.EntireColumn.ColumnWidth = 100
    lo.ListColumns(3).DataBodyRange.WrapText = True
    lo.DataBodyRange.VerticalAlignment = xlTop
    EnsureProjectHelp = True
End Function

Private Sub AddRule(c As Collection, section As String, txt As String)
    c.Add Array(section, Sym(txt))
End Sub

Private Function ProjectRules() As Collection
    Dim c As New Collection
    AddRule c, "Розміщення", "Шапка кожної таблиці - не вище рядка 3. Рядок над шапкою порожній; ще вище (рядок 1) - підпис таблиці, який ставить макрос UpdateHelp."
    AddRule c, "Розміщення", "Таблиці ростуть униз і вправо, тому кілька таблиць на одному аркуші ставлять лише поруч, ніколи одна під одною."
    AddRule c, "Розміщення", "Таблиця з запиту Power Query (q_, pq_) - одна на аркуші. Аркуш бажано називати так само, як запит."
    AddRule c, "Розміщення", "Кожна таблиця займає щонайменше 3 колонки (місце для підпису). Між сусідніми таблицями - щонайменше одна порожня колонка."
    AddRule c, "Розміщення", "Нижче і правіше таблиці нічого не розміщують: ні формул, ні приміток."
    AddRule c, "Назви", "Префікс задає роль таблиці: tbl - ручне введення і налаштування; q_ - дані призначення (результат запиту); pq_ - діагностика і звіти Engine; hlp - довідка; test_ - тести (лише тестова книга)."
    AddRule c, "Назви", "Назви таблиць унікальні і завжди з префіксом. Довідка до tblX за замовчуванням - hlpX."
    AddRule c, "Довідка", "Аркуш HELP: tblTableCatalog, hlpProject (ця таблиця) і всі hlp-таблиці - поруч, шапки в рядку 3."
    AddRule c, "Довідка", "Призначення таблиці описують лише в tblTableCatalog[Призначення]. Правила заповнення колонок - у hlp-таблиці, вказаній у tblTableCatalog[Довідка]."
    AddRule c, "Каталог", "Колонки каталогу, які заповнюєте ви: Призначення; Підпис (TRUE - показувати підпис над таблицею, FALSE - ні); Довідка (назва hlp-таблиці). Решту заповнює UpdateHelp, ваші значення зберігаються між запусками."
    AddRule c, "Каталог", "Колонка «Порушення» показує відхилення від цих правил. Макрос їх не виправляє; єдина автоматична зміна - вставка рядків над таблицею, шапка якої вище рядка 3."
    AddRule c, "Підпис", "Підпис - три клітинки над першими колонками таблиці: назва | «{i} Призначення» | «? Як заповнювати» (для hlp-таблиці - «{<-} до tblX»). Вручну не редагують: UpdateHelp перезаписує їх."
    AddRule c, "Нова таблиця", "Нова таблиця з запиту: завантажити на новий аркуш {->} запустити UpdateHelp (додасть у каталог, зсуне шапку в рядок 3, поставить підпис) {->} заповнити «Призначення», за потреби створити hlp-таблицю і вписати її в «Довідку»."
    AddRule c, "Після змін", "Після вставлення рядків чи колонок перед таблицею, перейменування таблиці чи аркуша - запустити UpdateHelp. Інакше підпис покаже «{!} немає в каталозі»."
    AddRule c, "Службові", "Без підпису за замовчуванням: tblTableCatalog, hlpProject, tblRunLog, test_… і tblTest…."
    AddRule c, "Технічне", "Книга має бути збережена: формула назви в підписі використовує CELL(""filename""). Потрібен Excel 2021 або новіший (LET, XLOOKUP, IFS)."
    Set ProjectRules = c
End Function

' Символи поза Windows-1251 (модуль зберігається в cp1251)
Private Function Sym(ByVal s As String) As String
    s = Replace(s, "{i}", ChrW(&H2139))
    s = Replace(s, "{!}", ChrW(&H26A0))
    s = Replace(s, "{<-}", ChrW(&H2190))
    s = Replace(s, "{->}", ChrW(&H2192))
    Sym = s
End Function

' ================== Правила таблиці ==================
Private Function HasKnownPrefix(nm As String) As Boolean
    Dim t As String: t = LCase$(nm)
    HasKnownPrefix = (Left$(t, 3) = "tbl" Or Left$(t, 3) = "hlp" Or Left$(t, 2) = "q_" Or _
                      Left$(t, 3) = "pq_" Or Left$(t, 5) = "test_")
End Function

Private Function DefaultLabel(nm As String) As Boolean
    Dim t As String: t = LCase$(nm)
    DefaultLabel = Not (t = LCase$(CAT_TABLE) Or t = LCase$(PROJECT_HELP) Or t = "tblrunlog" Or _
                        Left$(t, 5) = "test_" Or Left$(t, 7) = "tbltest")
End Function

Private Function DefaultHelp(nm As String) As String
    Dim lo As ListObject
    If LCase$(Left$(nm, 3)) <> "tbl" Then Exit Function
    Set lo = engFindTable("hlp" & Mid$(nm, 4))
    If Not lo Is Nothing Then DefaultHelp = lo.Name
End Function

' Підпис: логічне як є; число <> 0; текст так/ні, true/false, істина/хибність; інакше dflt
Private Function ToBoolLabel(v As Variant, dflt As Boolean) As Boolean
    Dim t As String
    ToBoolLabel = dflt
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbBoolean Then ToBoolLabel = v: Exit Function
    If VarType(v) <> vbString And IsNumeric(v) Then ToBoolLabel = (v <> 0): Exit Function
    t = LCase$(Trim$(CStr(v)))
    Select Case t
        Case "true", "так", "істина", "1", "yes": ToBoolLabel = True
        Case "false", "ні", "хибність", "0", "no": ToBoolLabel = False
    End Select
End Function

Private Sub AddIssue(ByRef s As String, t As String)
    s = s & IIf(s = "", "", "; ") & t
End Sub

' Колонки блоку таблиці: від першої колонки, ширина не менша за LABEL_COLS
Private Sub BlockCols(lo As ListObject, ByRef c1 As Long, ByRef c2 As Long)
    c1 = lo.Range.Column
    c2 = c1 + Application.Max(lo.Range.Columns.Count, LABEL_COLS) - 1
End Sub

Private Function IsQueryTable(lo As ListObject) As Boolean
    IsQueryTable = (lo.SourceType = xlSrcQuery Or lo.SourceType = xlSrcModel)
End Function

Private Function LabelRange(lo As ListObject) As Range
    If lo.Range.Row - 2 < 1 Then Exit Function
    Set LabelRange = lo.Parent.Cells(lo.Range.Row - 2, lo.Range.Column).Resize(1, LABEL_COLS)
End Function

' Клітинки зайняті: є вміст, об'єднання або перетин з таблицею
Private Function RangeBusy(rng As Range) As Boolean
    Dim c As Range, lo As ListObject
    If Not IsNull(rng.MergeCells) Then
        If rng.MergeCells Then RangeBusy = True: Exit Function
    Else
        RangeBusy = True: Exit Function
    End If
    For Each c In rng.Cells
        If c.formula <> "" Then RangeBusy = True: Exit Function
    Next c
    For Each lo In rng.Worksheet.ListObjects
        If Not Intersect(lo.Range, rng) Is Nothing Then RangeBusy = True: Exit Function
    Next lo
End Function

Private Function TableIssues(lo As ListObject, lbl As Boolean, helpName As String, desc As String, _
                             ByRef labelOk As Boolean) As String
    Dim s As String, other As ListObject, a1 As Long, a2 As Long, b1 As Long, b2 As Long, rng As Range

    If Not HasKnownPrefix(lo.Name) Then AddIssue s, "назва без префікса (tbl, hlp, q_, pq_, test_)"
    If lo.Range.Row < HEADER_ROW_MIN Then AddIssue s, "шапка вище рядка " & HEADER_ROW_MIN

    BlockCols lo, a1, a2
    For Each other In lo.Parent.ListObjects
        If other.Name <> lo.Name Then
            BlockCols other, b1, b2
            If a1 <= b2 And b1 <= a2 Then
                AddIssue s, "перетинається по колонках з " & other.Name & _
                            " (таблиці на аркуші - лише поруч, кожна щонайменше на " & LABEL_COLS & " колонки)"
            ElseIf a2 + 1 = b1 Or b2 + 1 = a1 Then
                AddIssue s, "немає порожньої колонки до " & other.Name
            End If
        End If
    Next other
    If IsQueryTable(lo) And lo.Parent.ListObjects.Count > 1 Then AddIssue s, "таблиця запиту не одна на аркуші"

    If helpName <> "" Then
        If engFindTable(helpName) Is Nothing Then AddIssue s, "довідку " & helpName & " не знайдено"
    End If
    If desc = "" Then AddIssue s, "не заповнено «Призначення»"

    labelOk = False
    If lbl Then
        Set rng = LabelRange(lo)
        If rng Is Nothing Then
            AddIssue s, "немає місця для підпису над шапкою"
        ElseIf RangeBusy(rng) Then
            AddIssue s, "рядок підпису " & rng.Address(False, False) & " зайнятий"
        Else
            labelOk = True
        End If
    End If
    TableIssues = s
End Function

' ================== Підписи над таблицями ==================
' Три формули в клітинках (Formula2, англійською; Excel показує їх мовою інтерфейсу).
' Excel 2021+: LET, XLOOKUP, IFS; LAMBDA і TEXTAFTER не потрібні.
'   1. Назва таблиці: аркуш (текст після останньої «]» у CELL("filename")) і перша клітинка шапки
'      (ADDRESS(...;4)) -> рядок каталогу.
'   2. Посилання на клітинку «Призначення» в каталозі.
'   3. Посилання на hlp-довідку з колонки «Довідка» або, для hlp-таблиці, назад до її таблиці.
Private Sub WriteLabel(lo As ListObject, wsCat As Worksheet)
    Dim rng As Range, t As String, tgt As String, nmRef As String, sh As String
    Dim info As String, warn As String, back As String
    Set rng = LabelRange(lo)
    If rng Is Nothing Then Exit Sub
    t = CAT_TABLE
    tgt = lo.Range.Cells(1, 1).Address(False, False)
    nmRef = rng.Cells(1, 1).Address(False, False)
    sh = Replace(wsCat.Name, "'", "''")
    info = ChrW(&H2139): warn = ChrW(&H26A0): back = ChrW(&H2190)

    rng.Cells(1, 1).Formula2 = _
        "=LET(target," & tgt & ",file,CELL(""filename"",target)," & _
        "sheet,MID(file,FIND(""|"",SUBSTITUTE(file,""]"",""|"",LEN(file)-LEN(SUBSTITUTE(file,""]"",""""))))+1,255)," & _
        "addr,ADDRESS(ROW(target),COLUMN(target),4),ranges," & t & "[Діапазон]," & _
        "firstCell,LEFT(ranges,FIND("":"",ranges&"":"")-1)," & _
        "XLOOKUP(1,(" & t & "[Аркуш]=sheet)*(firstCell=addr)," & t & "[Таблиця],""""))"

    rng.Cells(1, 2).Formula2 = _
        "=LET(key," & nmRef & ",hit,XLOOKUP(key," & t & "[Таблиця]," & t & "[Призначення],"""")," & _
        "IF(OR(key="""",hit=""""),IF(key="""",""" & warn & " немає в каталозі"",""" & info & " Призначення (не заповнено)"")," & _
        "HYPERLINK(""#'" & sh & "'!""&ADDRESS(ROW(hit),COLUMN(hit),4),""" & info & " Призначення"")))"

    rng.Cells(1, 3).Formula2 = _
        "=LET(tn," & nmRef & "," & _
        "hlp,XLOOKUP(tn," & t & "[Таблиця]," & t & "[Довідка],"""")," & _
        "owner,XLOOKUP(tn," & t & "[Довідка]," & t & "[Таблиця],"""")," & _
        "IFS(tn="""",""""," & _
        "hlp<>"""",IFERROR(IF(ISREF(INDIRECT(hlp)),HYPERLINK(""#""&hlp,""? Як заповнювати""),""" & warn & " ""&hlp&"" не знайдено""),""" & warn & " ""&hlp&"" не знайдено"")," & _
        "owner<>"""",HYPERLINK(""#""&owner,""" & back & " до ""&owner)," & _
        "TRUE,""""))"

    rng.Cells(1, 1).Font.Bold = True
    rng.Cells(1, 1).ClearComments
    rng.Cells(1, 1).AddComment LABEL_NOTE
    rng.Cells(1, 1).Comment.Shape.TextFrame.AutoSize = True
    With rng.Cells(1, 2).Resize(1, 2).Font
        .Color = RGB(5, 99, 193)
        .Underline = xlUnderlineStyleSingle
    End With
    ThisWorkbook.names.Add Name:=LABEL_NAME_PREFIX & lo.Name, _
        RefersTo:="='" & Replace(lo.Parent.Name, "'", "''") & "'!" & rng.Address, Visible:=False
End Sub

' Очищає всі підписи попереднього запуску (за прихованими іменами engLbl_…).
' Клітинки, на які тепер заходить таблиця, не чіпає.
Private Sub ClearAllLabels()
    Dim i As Long, nmObj As Name, rng As Range, lo As ListObject, clash As Boolean
    For i = ThisWorkbook.names.Count To 1 Step -1
        Set nmObj = ThisWorkbook.names(i)
        If Left$(nmObj.Name, Len(LABEL_NAME_PREFIX)) = LABEL_NAME_PREFIX Then
            Set rng = Nothing
            On Error Resume Next
            Set rng = nmObj.RefersToRange
            On Error GoTo 0
            If Not rng Is Nothing Then
                clash = False
                For Each lo In rng.Worksheet.ListObjects
                    If Not Intersect(lo.Range, rng) Is Nothing Then clash = True: Exit For
                Next lo
                If Not clash Then
                    rng.ClearContents
                    rng.ClearFormats
                    rng.ClearComments
                End If
            End If
            nmObj.Delete
        End If
    Next i
End Sub

' ================== Властивості таблиці ==================
Private Function KindOf(lo As ListObject) As String
    Select Case lo.SourceType
        Case xlSrcRange: KindOf = "Ручна"
        Case xlSrcQuery: KindOf = "Результат запиту"
        Case xlSrcModel: KindOf = "З моделі даних"
        Case Else: KindOf = "Інше (" & lo.SourceType & ")"
    End Select
End Function

Private Function QueryOf(lo As ListObject) As String
    On Error Resume Next
    If lo.SourceType = xlSrcQuery Then
        QueryOf = engQueryNameOf(lo.QueryTable.WorkbookConnection)
    ElseIf lo.SourceType = xlSrcModel Then
        QueryOf = engQueryNameOf(lo.TableObject.WorkbookConnection)
    End If
    On Error GoTo 0
End Function

' таблиця підключена до моделі даних як зв'язана (підключення WorksheetConnection_…!<таблиця>)
Private Function InModel(lo As ListObject) As Boolean
    Dim cn As WorkbookConnection
    For Each cn In ThisWorkbook.Connections
        If Right$(cn.Name, Len(lo.Name) + 1) = "!" & lo.Name Then
            If InStr(1, cn.Name, "WorksheetConnection", vbTextCompare) > 0 Then InModel = True: Exit Function
        End If
    Next cn
End Function

Private Function ColumnList(lo As ListObject) As String
    Dim lc As ListColumn, s As String
    For Each lc In lo.ListColumns
        s = s & IIf(s = "", "", ", ") & lc.Name
    Next lc
    ColumnList = s
End Function

' ================== Стартові описи ==================
' Лише для відомих таблиць Engine і за правилами іменування; далі опис живе в каталозі.
Private Function DefaultDescription(tblName As String, qName As String) As String
    Dim t As String: t = LCase$(tblName)
    Select Case t
        Case LCase$(engTableName("Registry"))
            DefaultDescription = "Реєстр джерел: один рядок = один файл (SourceID, призначення, шлях, період, політика помилок)."
        Case LCase$(engTableName("Schema"))
            DefaultDescription = "Схема даних: колонки кожного призначення, цільові назви, типи, обов'язковість, перетворення."
        Case LCase$(engTableName("RowFilters"))
            DefaultDescription = "Правила відсікання службових рядків (Всього, х тощо) до типізації."
        Case LCase$(engTableName("DefaultHeaders"))
            DefaultDescription = "Шаблон заголовків призначення на випадок, коли шлях до файла не задано."
        Case LCase$(engTableName("TargetTypes"))
            DefaultDescription = "Довідник призначень (TargetType); список для Data Validation."
        Case LCase$(engTableName("PeriodRules"))
            DefaultDescription = "Довідник типів періодів і максимальних номерів (PeriodType, MaxNumber)."
        Case LCase$(engTableName("Period"))
            DefaultDescription = "Межі звітного періоду (StartDate, EndDate) для календаря Calendar."
        Case "tblschemacolumns"
            DefaultDescription = "Каталог колонок файлів по призначеннях (результат pq_ExtractedColumns); джерело списків Data Validation."
        Case "tbldatatype"
            DefaultDescription = "Довідник типів даних для колонки DataType схеми (Data Validation)."
        Case "tbllogical"
            DefaultDescription = "Довідник логічних значень (TRUE/FALSE тощо) для Data Validation."
        Case "tblerrorpolicy"
            DefaultDescription = "Довідник політик помилок (Strict, Report, Lenient) з описами."
        Case "tbloperator"
            DefaultDescription = "Довідник операторів фільтрів службових рядків (Data Validation)."
        Case "tblcustomtransform"
            DefaultDescription = "Довідник дій CustomTransform схеми (Data Validation)."
        Case "tblnums"
            DefaultDescription = "Номери періодів (N) для Data Validation колонки PeriodNumber."
        Case "tblrunlog"
            DefaultDescription = "Журнал запусків RefreshWithRunLog: час, статус і якість кожного джерела."
        Case "tbltestexpected"
            DefaultDescription = "Тестова книга: очікувані значення зведення якості (test_Summary)."
        Case "tbltestload"
            DefaultDescription = "Тестова книга: сценарії завантаження fn_ReadPurpose (test_Load)."
        Case LCase$(PROJECT_HELP)
            DefaultDescription = "Загальна довідка проєкту: правила структури книги, каталогу і довідки."
    End Select
    If DefaultDescription <> "" Then Exit Function

    If Left$(t, 3) = "hlp" Then
        DefaultDescription = "Довідка: як заповнювати колонки таблиці tbl" & Mid$(tblName, 4) & "."
    ElseIf qName = "pq_QualitySummary" Then
        DefaultDescription = "Зведення якості по джерелах (Status, QualityFlag, кількість рядків і помилок)."
    ElseIf Left$(qName, 2) = "q_" Then
        DefaultDescription = "Дані призначення (результат " & qName & ")."
    ElseIf Left$(qName, 5) = "test_" Then
        DefaultDescription = "Результат тесту " & qName & "."
    ElseIf qName <> "" Then
        DefaultDescription = "Результат запиту " & qName & "."
    End If
End Function
