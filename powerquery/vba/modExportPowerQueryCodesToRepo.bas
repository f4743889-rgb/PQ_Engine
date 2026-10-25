Attribute VB_Name = "modExportPowerQueryCodesToRepo"
Option Explicit

' ============================================================================
' modExportPowerQueryCodesToRepo: експорт коду Power Query, таблиць, модулів VBA і README для Git (v4.11)
' (етап D, аудит D13). Потребує modEngineCommon.
'
' Структура поруч із книгою:
'   powerquery\queries\<ім'я>.pq     код кожного запиту (UTF-8)
'   powerquery\metadata\queries.yml  тека (group, з cfgQueryGroups), опис (Description), тип, таблиця
'                                    завантаження, залежності
'   powerquery\tables\<таблиця>.csv  таблиці книги, які заповнюють вручну (реєстр, схема,
'                                    фільтри, довідники); таблиці-результати запитів не
'                                    експортуються, бо відтворюються оновленням; журнал
'                                    tblRunLog і каталог tblTableCatalog теж ні (EXCLUDED_TABLES:
'                                    історія і мінливі колонки; зміст каталогу - у README)
'   powerquery\vba\<модуль>.bas|.cls|.frm  модулі VBA (VBComponent.Export, ANSI = cp1251;
'                                    .gitattributes зберігає їх у Git в UTF-8)
'   powerquery\README.md             довідка проєкту (генерується, вручну не редагувати):
'                                    правила з hlpProject, таблиці з tblTableCatalog,
'                                    hlp-довідки до таблиць, запити за теками cfgQueryGroups,
'                                    макроси VBA
' Перед експортом старі .pq, .csv і модулі VBA видаляються, тож Git бачить і видалені запити,
' таблиці чи модулі. Модулі VBA потребують «Довіряти доступ до об'єктної моделі проєкту VBA»
' (Центр безпеки -> Параметри макросів); без доступу тека vba не чіпається.
' Клітинки з формулами експортуються формулою (наприклад, Path = =TestDir).
' Тестову книгу (з tblTestExpected) експорт не обробляє: репозиторій дзеркалить робочу книгу.
' Залежності шукаються за точними ідентифікаторами M без коментарів і текстових літералів.
' ============================================================================

' Таблиці, які не експортуються: журнал запусків накопичує історію; каталог таблиць має мінливі
' колонки (Діапазон, Рядків, Порушення), а його зміст потрапляє в README.md
Private Const EXCLUDED_TABLES As String = "|tblRunLog|tblTableCatalog|"
' Ознака тестової книги: її не експортують (репозиторій дзеркалить робочу книгу)
Private Const TEST_MARKER_TABLE As String = "tblTestExpected"

' Карта тек з cfgQueryGroups (заповнює LoadQueryGroups)
Private mGroups As Object        ' ім'я запиту -> тека
Private mPrefixes As Collection  ' елементи Array(префікс, тека)
Private mUnmapped As String      ' запити без теки (для підсумку)
Private mUnmappedN As Long
Private Const GROUPS_QUERY As String = "cfgQueryGroups"
Private Const README_RULES As String = "hlpProject"
Private Const README_CATALOG As String = "tblTableCatalog"
Private mReadmeNote As String    ' що потрапило в README (для підсумку)

Public Sub ExportPowerQueriesToRepo()
    Dim root As String, qDir As String, mDir As String, tDir As String
    Dim q As WorkbookQuery, ws As Worksheet, lo As ListObject, nQ As Long, nT As Long

    root = ThisWorkbook.path
    If root = "" Then
        MsgBox "Спершу збережіть книгу: файли експортуються в теку поруч із нею.", vbExclamation
        Exit Sub
    End If
    ' Репозиторій дзеркалить робочу книгу; тестова (з tblTestExpected) замінила б робочу
    ' конфігурацію (реєстр, схему, фільтри) тестовою. Тестова конфігурація - у tests\test_config.xlsx.
    If Not engFindTable(TEST_MARKER_TABLE) Is Nothing Then
        MsgBox "Це тестова книга (є таблиця " & TEST_MARKER_TABLE & "): експорт зупинено." & vbCrLf & vbCrLf & _
               "Теку powerquery\ у репозиторії заповнює лише робоча книга. Тестова конфігурація " & _
               "зберігається в tests\test_config.xlsx, тестові запити - у tests\queries.", _
               vbExclamation, "Export"
        Exit Sub
    End If

    qDir = root & "\powerquery\queries"
    mDir = root & "\powerquery\metadata"
    tDir = root & "\powerquery\tables"

    On Error GoTo Fail
    engEnsureFolder qDir
    engEnsureFolder mDir
    engEnsureFolder tDir
    ClearFolder qDir, "pq"
    ClearFolder tDir, "csv"

    For Each q In ThisWorkbook.Queries
        engWriteUtf8 qDir & "\" & engSafeFileName(q.Name) & ".pq", q.formula
        nQ = nQ + 1
    Next q
    Dim groupsNote As String
    groupsNote = LoadQueryGroups()
    engWriteUtf8 mDir & "\queries.yml", BuildQueriesYaml()
    engWriteUtf8 root & "\powerquery\README.md", BuildReadme()

    Dim vbaNote As String, nV As Long
    vbaNote = ExportVbaModules(root & "\powerquery\vba", nV)

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            ' лише звичайні таблиці (не результати запитів) і не журнали (історія, а не конфігурація)
            If lo.SourceType = xlSrcRange And InStr(1, EXCLUDED_TABLES, "|" & lo.Name & "|", vbTextCompare) = 0 Then
                engWriteUtf8 tDir & "\" & engSafeFileName(lo.Name) & ".csv", TableToCsv(lo)
                nT = nT + 1
            End If
        Next lo
    Next ws

    MsgBox "Експортовано:" & vbCrLf & _
           "запитів Power Query: " & nQ & vbCrLf & _
           "таблиць конфігурації: " & nT & vbCrLf & _
           "модулів VBA: " & nV & vbCrLf & _
           "README.md: " & mReadmeNote & vbCrLf & _
           IIf(vbaNote <> "", vbCrLf & vbaNote & vbCrLf, "") & _
           IIf(groupsNote <> "", vbCrLf & groupsNote & vbCrLf, "") & _
           IIf(mUnmappedN > 0, vbCrLf & "Без теки в " & GROUPS_QUERY & " (" & mUnmappedN & "):" & mUnmapped & vbCrLf, "") & _
           vbCrLf & root & "\powerquery", IIf(mUnmappedN > 0 Or groupsNote <> "" Or vbaNote <> "", vbExclamation, vbInformation), "Export complete"
    Exit Sub
Fail:
    MsgBox "Помилка експорту: " & Err.Description, vbCritical, "Export"
End Sub

' ================== VBA ==================
' Модулі проєкту -> vbaDir\<ім'я>.bas (стандартні), .cls (класи), .frm/.frx (форми); модулі
' аркушів і книги - як .cls, лише якщо в них є код (крім рядків Option).
' VBComponent.Export пише в ANSI (cp1251); .gitattributes (working-tree-encoding=CP1251) зберігає
' модулі в Git в UTF-8, на диску після checkout - знову cp1251 і CRLF для Import у VBE.
' Тека очищується лише після перевірки доступу: інакше Git побачив би всі модулі видаленими.
' Повертає "" або текст попередження; n - кількість експортованих модулів.
Private Function ExportVbaModules(vbaDir As String, ByRef n As Long) As String
    Dim proj As Object, comp As Object, ext As String

    n = 0
    ExportVbaModules = VbaAccessProblem(proj)
    If ExportVbaModules <> "" Then
        ExportVbaModules = "Модулі VBA не експортовано (тека vba не змінена): " & ExportVbaModules
        Exit Function
    End If

    engEnsureFolder vbaDir
    ClearFolder vbaDir, "bas"
    ClearFolder vbaDir, "cls"
    ClearFolder vbaDir, "frm"
    ClearFolder vbaDir, "frx"

    For Each comp In proj.VBComponents
        Select Case comp.Type
            Case 1: ext = "bas"                       ' vbext_ct_StdModule
            Case 2: ext = "cls"                       ' vbext_ct_ClassModule
            Case 3: ext = "frm"                       ' vbext_ct_MSForm (.frx Export створює сам)
            Case 100: If HasCode(comp) Then ext = "cls" Else ext = ""    ' vbext_ct_Document
            Case Else: ext = ""
        End Select
        If ext <> "" Then
            comp.Export vbaDir & "\" & engSafeFileName(comp.Name) & "." & ext
            n = n + 1
        End If
    Next comp
End Function

' "" - доступ до проєкту VBA є (proj заповнено), інакше причина
Private Function VbaAccessProblem(ByRef proj As Object) As String
    Dim cnt As Long
    Set proj = Nothing
    On Error Resume Next
    Set proj = ThisWorkbook.VBProject
    cnt = proj.VBComponents.Count
    If Err.Number <> 0 Or proj Is Nothing Then
        VbaAccessProblem = "увімкніть «Довіряти доступ до об'єктної моделі проєкту VBA» у Центрі безпеки."
    ElseIf proj.Protection = 1 Then                   ' vbext_pp_locked
        VbaAccessProblem = "проєкт VBA захищено паролем."
    End If
    On Error GoTo 0
End Function

' Чи є в модулі код, крім порожніх рядків і Option ...
Private Function HasCode(comp As Object) As Boolean
    Dim cm As Object, i As Long, s As String
    Set cm = comp.CodeModule
    For i = 1 To cm.CountOfLines
        s = Trim$(cm.Lines(i, 1))
        If s <> "" And StrComp(Left$(s, 7), "Option ", vbTextCompare) <> 0 Then HasCode = True: Exit Function
    Next i
End Function

' ================== queries.yml ==================
' Сортування рядків за абеткою однаково на будь-якій машині: порівняння LCase$ за кодами
' символів (vbBinaryCompare), а не за правилами мови Windows (vbTextCompare) - інакше порядок
' у queries.yml і README залежав би від регіональних налаштувань (рецензія 25.10.2026)
Private Sub SortStrings(ByRef arr() As String)
    Dim i As Long, j As Long, tmp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(LCase$(arr(j)), LCase$(arr(i)), vbBinaryCompare) < 0 Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub

' Запити книги за абеткою: порядок ThisWorkbook.Queries змінюється, коли запит видаляють і
' імпортують знову, і queries.yml та README «змінювалися» б без змін у коді
Private Function SortedQueries() As Collection
    Dim q As WorkbookQuery, gk() As String, i As Long, res As New Collection
    If ThisWorkbook.Queries.Count = 0 Then Set SortedQueries = res: Exit Function
    ReDim gk(0 To ThisWorkbook.Queries.Count - 1)
    For Each q In ThisWorkbook.Queries
        gk(i) = q.Name: i = i + 1
    Next q
    SortStrings gk
    For i = 0 To UBound(gk)
        res.Add ThisWorkbook.Queries(gk(i))
    Next i
    Set SortedQueries = res
End Function

Private Function BuildQueriesYaml() As String
    Dim qNames As Object, q As WorkbookQuery, ids As Object, dep As Variant
    Dim txt As String, deps As String, loadTo As String, lo As ListObject

    Set qNames = CreateObject("Scripting.Dictionary")
    qNames.CompareMode = vbBinaryCompare
    For Each q In ThisWorkbook.Queries
        qNames(q.Name) = True
    Next q

    txt = "queries:" & vbCrLf
    For Each q In SortedQueries()
        Set ids = engMIdentifiers(q.formula)
        deps = ""
        For Each dep In ids.Keys
            If qNames.Exists(CStr(dep)) And CStr(dep) <> q.Name Then
                deps = deps & "      - " & YamlQuote(CStr(dep)) & vbCrLf
            End If
        Next dep

        loadTo = ""
        Set lo = engFindQueryTable(q.Name)
        If Not lo Is Nothing Then loadTo = lo.Name

        txt = txt & "  - name: " & YamlQuote(q.Name) & vbCrLf
        txt = txt & "    file: " & YamlQuote("../queries/" & engSafeFileName(q.Name) & ".pq") & vbCrLf
        txt = txt & "    group: " & YamlQuote(GroupOf(q.Name)) & vbCrLf
        txt = txt & "    description: " & YamlQuote(QueryDescription(q)) & vbCrLf
        txt = txt & "    type: " & YamlQuote(DetectQueryType(q.formula, deps <> "")) & vbCrLf
        txt = txt & "    loads_to_table: " & YamlQuote(loadTo) & vbCrLf
        If deps = "" Then
            txt = txt & "    depends_on: []" & vbCrLf
        Else
            txt = txt & "    depends_on:" & vbCrLf & deps
        End If
        txt = txt & vbCrLf
    Next q
    BuildQueriesYaml = txt
End Function

' ================== README.md ==================
' Довідка проєкту з таблиць книги: правила (hlpProject), таблиці (tblTableCatalog), hlp-довідки
' до таблиць, запити за теками (cfgQueryGroups), макроси VBA. Без дати і без мінливих колонок каталогу
' (Діапазон, Рядків, Колонок, Порушення), щоб Git бачив лише змістовні зміни.
Private Function BuildReadme() As String
    Dim t As String, lo As ListObject, cat As ListObject, ws As Worksheet
    Dim r As Long, k As Long, sec As String, prev As String, first As Boolean
    mReadmeNote = ""

    t = "# " & MdText(FileBaseName(ThisWorkbook.Name)) & vbCrLf & vbCrLf
    t = t & "> Файл створює макрос ExportPowerQueriesToRepo з таблиць книги (" & README_RULES & ", " & _
        README_CATALOG & ", hlp-таблиці, " & GROUPS_QUERY & "). Вручну не редагувати: зміни вносяться в книгу." & _
        vbCrLf & vbCrLf

    ' 1. Правила структури книги
    Set lo = engFindTable(README_RULES)
    If Not lo Is Nothing Then
        If engColumnExists(lo, "Розділ") And engColumnExists(lo, "Правило") And Not lo.DataBodyRange Is Nothing Then
            t = t & "## Правила структури книги" & vbCrLf
            first = True
            For r = 1 To lo.ListRows.Count
                sec = Trim$(CStr(engCell(lo, r, "Розділ")))
                If first Or sec <> prev Then
                    t = t & vbCrLf & "### " & MdText(IIf(sec = "", "Інше", sec)) & vbCrLf & vbCrLf
                    prev = sec: k = 0: first = False
                End If
                k = k + 1
                t = t & k & ". " & MdText(CStr(engCell(lo, r, "Правило"))) & vbCrLf
            Next r
            t = t & vbCrLf
            AddReadmeNote "правила"
        End If
    End If

    ' 2. Таблиці книги
    Set cat = engFindTable(README_CATALOG)
    If Not cat Is Nothing Then
        If Not cat.DataBodyRange Is Nothing And engColumnExists(cat, "Таблиця") Then
            t = t & "## Таблиці книги" & vbCrLf & vbCrLf
            t = t & "| Таблиця | Аркуш | Тип | Запит | Призначення | Довідка |" & vbCrLf
            t = t & "| --- | --- | --- | --- | --- | --- |" & vbCrLf
            For r = 1 To cat.ListRows.Count
                t = t & "| " & MdText(CatText(cat, r, "Таблиця")) & _
                        " | " & MdText(CatText(cat, r, "Аркуш")) & _
                        " | " & MdText(CatText(cat, r, "Тип")) & _
                        " | " & MdText(CatText(cat, r, "Запит")) & _
                        " | " & MdText(CatText(cat, r, "Призначення")) & _
                        " | " & HelpLink(CatText(cat, r, "Довідка")) & " |" & vbCrLf
            Next r
            t = t & vbCrLf
            AddReadmeNote "таблиці"
        End If
    End If

    ' 3. Довідки до таблиць (hlp-таблиці, крім загальної)
    first = True
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If LCase$(Left$(lo.Name, 3)) = "hlp" And StrComp(lo.Name, README_RULES, vbTextCompare) <> 0 Then
                If first Then t = t & "## Довідки до таблиць" & vbCrLf & vbCrLf: first = False
                t = t & "### " & MdText(lo.Name) & vbCrLf & vbCrLf
                If Not cat Is Nothing Then
                    sec = OwnerOfHelp(cat, lo.Name)
                    If sec <> "" Then t = t & "Як заповнювати таблицю **" & MdText(sec) & "**." & vbCrLf & vbCrLf
                End If
                t = t & MdTable(lo) & vbCrLf
            End If
        Next lo
    Next ws
    If Not first Then AddReadmeNote "hlp-довідки"

    ' 4. Запити Power Query за теками
    t = t & QueriesSection()
    AddReadmeNote "запити"

    ' 5. Макроси VBA
    t = t & MacrosSection()
    AddReadmeNote "макроси"
    BuildReadme = t
End Function

Private Sub AddReadmeNote(s As String)
    mReadmeNote = mReadmeNote & IIf(mReadmeNote = "", "", ", ") & s
End Sub

Private Function FileBaseName(fileName As String) As String
    If InStrRev(fileName, ".") > 1 Then FileBaseName = Left$(fileName, InStrRev(fileName, ".") - 1) Else FileBaseName = fileName
End Function

' Текст для Markdown: без розриву таблиці «|» і рядків
Private Function MdText(ByVal s As String) As String
    s = Replace(s, vbCrLf, "<br>")
    s = Replace(s, vbLf, "<br>")
    s = Replace(s, vbCr, "<br>")
    MdText = Trim$(Replace(s, "|", "\|"))
End Function

Private Function CellText(c As Range) As String
    Dim v As Variant
    v = c.value
    If IsError(v) Then
        CellText = c.text
    ElseIf VarType(v) = vbBoolean Then
        CellText = IIf(v, "TRUE", "FALSE")
    ElseIf IsEmpty(v) Then
        CellText = ""
    Else
        CellText = CStr(v)
    End If
End Function

Private Function CatText(cat As ListObject, r As Long, col As String) As String
    If engColumnExists(cat, col) Then CatText = CellText(cat.ListColumns(col).DataBodyRange.Cells(r, 1))
End Function

' Посилання на розділ README з hlp-таблицею (якір GitHub: заголовок у нижньому регістрі)
Private Function HelpLink(hlp As String) As String
    If hlp = "" Then Exit Function
    If engFindTable(hlp) Is Nothing Then
        HelpLink = MdText(hlp)
    Else
        HelpLink = "[" & MdText(hlp) & "](#" & LCase$(hlp) & ")"
    End If
End Function

Private Function OwnerOfHelp(cat As ListObject, hlp As String) As String
    Dim r As Long
    If Not engColumnExists(cat, "Довідка") Then Exit Function
    For r = 1 To cat.ListRows.Count
        If StrComp(CatText(cat, r, "Довідка"), hlp, vbTextCompare) = 0 Then
            OwnerOfHelp = CatText(cat, r, "Таблиця"): Exit Function
        End If
    Next r
End Function

Private Function MdTable(lo As ListObject) As String
    Dim t As String, ln As String, c As Range, r As Long, i As Long
    For Each c In lo.HeaderRowRange.Cells
        ln = ln & " " & MdText(CStr(c.value)) & " |"
    Next c
    t = "|" & ln & vbCrLf & "|"
    For i = 1 To lo.ListColumns.Count
        t = t & " --- |"
    Next i
    t = t & vbCrLf
    If Not lo.DataBodyRange Is Nothing Then
        For r = 1 To lo.DataBodyRange.Rows.Count
            ln = ""
            For Each c In lo.DataBodyRange.Rows(r).Cells
                ln = ln & " " & MdText(CellText(c)) & " |"
            Next c
            t = t & "|" & ln & vbCrLf
        Next r
    End If
    MdTable = t
End Function

' Запити за теками cfgQueryGroups (теки за назвою: «00 …», «01 …»), посилання на .pq
Private Function QueriesSection() As String
    Dim qNames As Object, byGroup As Object, q As WorkbookQuery, ids As Object, dep As Variant
    Dim g As String, hasDeps As Boolean, gk() As String, i As Long, t As String

    Set qNames = CreateObject("Scripting.Dictionary")
    qNames.CompareMode = vbBinaryCompare
    For Each q In ThisWorkbook.Queries
        qNames(q.Name) = True
    Next q

    Set byGroup = CreateObject("Scripting.Dictionary")
    For Each q In SortedQueries()
        g = GroupOf(q.Name, False)
        If g = "" Then g = "Без теки"
        Set ids = engMIdentifiers(q.formula)
        hasDeps = False
        For Each dep In ids.Keys
            If qNames.Exists(CStr(dep)) And CStr(dep) <> q.Name Then hasDeps = True: Exit For
        Next dep
        byGroup(g) = byGroup(g) & "- [" & MdText(q.Name) & "](<queries/" & engSafeFileName(q.Name) & ".pq>) — " & _
                     DetectQueryType(q.formula, hasDeps) & _
                     IIf(QueryDescription(q) <> "", " — " & MdText(QueryDescription(q)), "") & vbCrLf
    Next q
    If byGroup.Count = 0 Then Exit Function

    ReDim gk(0 To byGroup.Count - 1)
    i = 0
    For Each dep In byGroup.Keys
        gk(i) = CStr(dep): i = i + 1
    Next dep
    SortStrings gk

    t = "## Запити Power Query" & vbCrLf & vbCrLf & _
        "Тека — як у редакторі Power Query (" & GROUPS_QUERY & "); тип і залежності — у metadata/queries.yml." & vbCrLf
    For i = 0 To UBound(gk)
        t = t & vbCrLf & "### " & MdText(gk(i)) & vbCrLf & vbCrLf & byGroup(gk(i))
    Next i
    QueriesSection = t
End Function

' Розділ «Макроси»: стандартні модулі, їхній опис (рядок «' <модуль>: опис» у заголовку) і
' макроси - публічні процедури Sub без параметрів (те, що видно в Alt+F8)
Private Function MacrosSection() As String
    Dim proj As Object, comp As Object, cm As Object, i As Long, s As String, t As String
    Dim desc As String, subs As String, nm As String, why As String
    Dim macroRows As Object, gk() As String, mk As Variant

    t = "## Макроси VBA" & vbCrLf & vbCrLf
    why = VbaAccessProblem(proj)
    If why <> "" Then
        MacrosSection = t & "_Список недоступний під час експорту: " & MdText(why) & "_" & vbCrLf & vbCrLf
        Exit Function
    End If

    t = t & "Код модулів — у теці vba/. Імпорт у книгу: редактор VBA, File > Import File (спершу modEngineCommon)." & _
        vbCrLf & vbCrLf & "| Модуль | Макроси (Alt+F8) | Опис |" & vbCrLf & "| --- | --- | --- |" & vbCrLf
    ' рядки збираються за назвою модуля і виводяться за абеткою: порядок VBComponents змінюється
    ' після кожного видалення й імпорту модуля, і README «змінювався» б без змін у коді
    Set macroRows = CreateObject("Scripting.Dictionary")
    For Each comp In proj.VBComponents
        If comp.Type = 1 Then
            Set cm = comp.CodeModule
            desc = "": subs = ""
            For i = 1 To cm.CountOfLines
                s = cm.Lines(i, 1)
                If desc = "" And Left$(s, Len(comp.Name) + 3) = "' " & comp.Name & ":" Then
                    desc = Trim$(Mid$(s, Len(comp.Name) + 4))
                End If
                nm = MacroName(s)
                If nm <> "" Then subs = subs & IIf(subs = "", "", ", ") & nm
            Next i
            macroRows(comp.Name) = "| [" & MdText(comp.Name) & "](<vba/" & engSafeFileName(comp.Name) & ".bas>) | " & _
                MdText(subs) & " | " & MdText(desc) & " |" & vbCrLf
        End If
    Next comp
    If macroRows.Count > 0 Then
        ReDim gk(0 To macroRows.Count - 1)
        i = 0
        For Each mk In macroRows.Keys
            gk(i) = CStr(mk): i = i + 1
        Next mk
        SortStrings gk
        For i = 0 To UBound(gk)
            t = t & macroRows(gk(i))
        Next i
    End If
    MacrosSection = t & vbCrLf
End Function

' Ім'я макроса з рядка «Sub X()» або «Public Sub X()»; інакше ""
Private Function MacroName(ByVal s As String) As String
    Dim p As Long
    If Left$(s, 11) = "Public Sub " Then
        s = Mid$(s, 12)
    ElseIf Left$(s, 4) = "Sub " Then
        s = Mid$(s, 5)
    Else
        Exit Function
    End If
    p = InStr(s, "(")
    If p < 2 Then Exit Function
    If Left$(Trim$(Mid$(s, p + 1)), 1) <> ")" Then Exit Function    ' з параметрами - не макрос
    MacroName = Trim$(Left$(s, p - 1))
End Function

' ================== Теки (cfgQueryGroups) ==================
' Читає Groups (#"тека" = {"запит", ...}) і Prefixes (префікс = "тека") з тексту cfgQueryGroups.
' Повертає "" або попередження.
Private Function LoadQueryGroups() As String
    Dim f As String, re As Object, rq As Object, rp As Object, it As Object, nm As Object, grp As String

    Set mGroups = CreateObject("Scripting.Dictionary")
    mGroups.CompareMode = vbBinaryCompare
    Set mPrefixes = New Collection
    mUnmapped = "": mUnmappedN = 0

    On Error Resume Next
    f = ThisWorkbook.Queries(GROUPS_QUERY).formula
    On Error GoTo 0
    If f = "" Then
        LoadQueryGroups = "Запит " & GROUPS_QUERY & " не знайдено: поле group у queries.yml порожнє."
        Exit Function
    End If
    f = engStripMComments(f)

    Set rq = CreateObject("VBScript.RegExp")
    rq.Global = True
    rq.Pattern = """((?:[^""]|"""")*)"""

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "#""((?:[^""]|"""")*)""\s*=\s*\{([^}]*)\}"
    For Each it In re.Execute(f)
        grp = engUnescapeM(it.SubMatches(0))
        For Each nm In rq.Execute(it.SubMatches(1))
            mGroups(engUnescapeM(nm.SubMatches(0))) = grp
        Next nm
    Next it

    re.Global = False
    re.Pattern = "\bPrefixes\s*=\s*\[([^\]]*)\]"
    If re.Test(f) Then
        Set rp = CreateObject("VBScript.RegExp")
        rp.Global = True
        rp.Pattern = "(\w+)\s*=\s*""((?:[^""]|"""")*)"""
        For Each it In rp.Execute(re.Execute(f)(0).SubMatches(0))
            mPrefixes.Add Array(CStr(it.SubMatches(0)), engUnescapeM(it.SubMatches(1)))
        Next it
    End If

    If mGroups.Count = 0 Then LoadQueryGroups = "У " & GROUPS_QUERY & " не знайдено жодної теки (Groups)."
End Function

' Тека запиту: точне ім'я в Groups, інакше перший відповідний префікс; "" = не задано
Private Function GroupOf(queryName As String, Optional countUnmapped As Boolean = True) As String
    Dim p As Variant
    If Not mGroups Is Nothing Then
        If mGroups.Exists(queryName) Then GroupOf = mGroups(queryName): Exit Function
    End If
    If Not mPrefixes Is Nothing Then
        For Each p In mPrefixes
            If Left$(queryName, Len(p(0))) = p(0) Then GroupOf = p(1): Exit Function
        Next p
    End If
    If Not countUnmapped Then Exit Function
    mUnmappedN = mUnmappedN + 1
    If mUnmappedN <= 10 Then mUnmapped = mUnmapped & vbCrLf & "  * " & queryName
End Function

' Function | Config | Source | Workbook | Transform | Standalone (за кодом без коментарів):
'   Config    - запис або список-літерал (cfg*): код починається з [ або {
'   Source    - відкриває файли, бази чи веб
'   Workbook  - читає таблиці чи іменовані клітинки книги (Excel.CurrentWorkbook)
Private Function DetectQueryType(mText As String, hasDeps As Boolean) As String
    Dim f As String, fn As Variant
    f = LCase$(Trim$(Replace(Replace(Replace(engStripMComments(mText), vbCr, " "), vbLf, " "), vbTab, " ")))

    If Left$(f, 1) = "(" Then DetectQueryType = "Function": Exit Function
    If Left$(f, 1) = "[" Or Left$(f, 1) = "{" Then DetectQueryType = "Config": Exit Function

    For Each fn In Array("sql.database(", "sql.databases(", "access.database(", "oracle.database(", _
                         "odbc.datasource(", "oledb.datasource(", "folder.files(", "folder.contents(", _
                         "file.contents(", "excel.workbook(", "csv.document(", "xml.tables(", _
                         "json.document(", "sharepoint.files(", "sharepoint.contents(", "web.contents(")
        If InStr(f, CStr(fn)) > 0 Then DetectQueryType = "Source": Exit Function
    Next fn

    If InStr(f, "excel.currentworkbook(") > 0 Then DetectQueryType = "Workbook": Exit Function

    If hasDeps Then DetectQueryType = "Transform" Else DetectQueryType = "Standalone"
End Function

' Опис запиту (властивість Description: підказка в редакторі Power Query); "" якщо немає
Private Function QueryDescription(q As WorkbookQuery) As String
    On Error Resume Next
    QueryDescription = Trim$(Replace(Replace(q.Description, vbCr, " "), vbLf, " "))
    On Error GoTo 0
End Function

Private Function YamlQuote(txtValue As String) As String
    YamlQuote = """" & Replace(Replace(txtValue, "\", "\\"), """", "\""") & """"
End Function

' ================== CSV ==================
' Роздільник ";", рядок заголовків = назви колонок, числа з крапкою, дати ISO, формули як формули
Private Function TableToCsv(lo As ListObject) As String
    Dim txt As String, ln As String, c As Range, r As Long
    For Each c In lo.HeaderRowRange.Cells
        ln = ln & IIf(ln = "", "", ";") & CsvQuote(CStr(c.value))
    Next c
    txt = ln & vbCrLf
    If Not lo.DataBodyRange Is Nothing Then
        For r = 1 To lo.DataBodyRange.Rows.Count
            ln = ""
            For Each c In lo.DataBodyRange.Rows(r).Cells
                ln = ln & IIf(c.Column = lo.DataBodyRange.Column, "", ";") & CsvField(c)
            Next c
            txt = txt & ln & vbCrLf
        Next r
    End If
    TableToCsv = txt
End Function

Private Function CsvField(c As Range) As String
    Dim v As Variant, s As String
    If c.HasFormula Then
        s = c.formula
    Else
        v = c.value
        If IsError(v) Then
            s = c.text
        ElseIf IsEmpty(v) Then
            s = ""
        ElseIf VarType(v) = vbBoolean Then
            s = IIf(v, "TRUE", "FALSE")
        ElseIf VarType(v) = vbDate Then
            If v = Int(v) Then s = Format$(v, "yyyy-mm-dd") Else s = Format$(v, "yyyy-mm-dd hh:nn:ss")
        ElseIf VarType(v) <> vbString And IsNumeric(v) Then
            s = Replace(CStr(v), Application.International(xlDecimalSeparator), ".")
        Else
            s = CStr(v)
        End If
    End If
    CsvField = CsvQuote(s)
End Function

Private Function CsvQuote(s As String) As String
    If InStr(s, ";") > 0 Or InStr(s, """") > 0 Or InStr(s, vbLf) > 0 Or InStr(s, vbCr) > 0 Then
        CsvQuote = """" & Replace(s, """", """""") & """"
    Else
        CsvQuote = s
    End If
End Function

' Видалення файлів із розширенням ext у теці (без підтек)
Private Sub ClearFolder(folderPath As String, ext As String)
    Dim fso As Object, f As Object, toDel As New Collection, p As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then Exit Sub
    For Each f In fso.GetFolder(folderPath).Files
        If LCase$(fso.GetExtensionName(f.Name)) = LCase$(ext) Then toDel.Add f.path
    Next f
    For Each p In toDel
        fso.DeleteFile CStr(p), True
    Next p
End Sub
