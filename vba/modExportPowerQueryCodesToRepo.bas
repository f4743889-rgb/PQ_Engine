Attribute VB_Name = "modExportPowerQueryCodesToRepo"
Option Explicit

' ============================================================================
' modExportPowerQueryCodesToRepo: експорт коду Power Query і таблиць конфігурації для Git
' (етап D, аудит D13). Потребує modEngineCommon.
'
' Структура поруч із книгою:
'   powerquery\queries\<ім'я>.pq     код кожного запиту (UTF-8)
'   powerquery\metadata\queries.yml  тека (group, з cfgQueryGroups), тип запиту, таблиця
'                                    завантаження, залежності
'   powerquery\tables\<таблиця>.csv  таблиці книги, які заповнюють вручну (реєстр, схема,
'                                    фільтри, довідники); таблиці-результати запитів не
'                                    експортуються, бо відтворюються оновленням; журнал
'                                    tblRunLog теж ні (EXCLUDED_TABLES)
' Перед експортом старі .pq і .csv видаляються, тож Git бачить і видалені запити чи таблиці.
' Клітинки з формулами експортуються формулою (наприклад, Path = =TestDir).
'
' Відмінності від попередньої версії: залежності шукаються за точними ідентифікаторами M
' без коментарів і текстових літералів (раніше fn_Opt «знаходився» в fn_Options, а все після
' першого // у рядку з vbLf губилося); додано тип Function і вивантаження таблиць у CSV.
' ============================================================================

' Таблиці, які не експортуються: журнал запусків накопичує історію і змінюється щоразу
Private Const EXCLUDED_TABLES As String = "|tblRunLog|"

' Карта тек з cfgQueryGroups (заповнює LoadQueryGroups)
Private mGroups As Object        ' ім'я запиту -> тека
Private mPrefixes As Collection  ' елементи Array(префікс, тека)
Private mUnmapped As String      ' запити без теки (для підсумку)
Private mUnmappedN As Long
Private Const GROUPS_QUERY As String = "cfgQueryGroups"

Public Sub ExportPowerQueriesToRepo()
    Dim root As String, qDir As String, mDir As String, tDir As String
    Dim q As WorkbookQuery, ws As Worksheet, lo As ListObject, nQ As Long, nT As Long

    root = ThisWorkbook.path
    If root = "" Then
        MsgBox "Спершу збережіть книгу: файли експортуються в теку поруч із нею.", vbExclamation
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
           IIf(groupsNote <> "", vbCrLf & groupsNote & vbCrLf, "") & _
           IIf(mUnmappedN > 0, vbCrLf & "Без теки в " & GROUPS_QUERY & " (" & mUnmappedN & "):" & mUnmapped & vbCrLf, "") & _
           vbCrLf & root & "\powerquery", IIf(mUnmappedN > 0 Or groupsNote <> "", vbExclamation, vbInformation), "Export complete"
    Exit Sub
Fail:
    MsgBox "Помилка експорту: " & Err.Description, vbCritical, "Export"
End Sub

' ================== queries.yml ==================
Private Function BuildQueriesYaml() As String
    Dim qNames As Object, q As WorkbookQuery, ids As Object, dep As Variant
    Dim txt As String, deps As String, loadTo As String, lo As ListObject

    Set qNames = CreateObject("Scripting.Dictionary")
    qNames.CompareMode = vbBinaryCompare
    For Each q In ThisWorkbook.Queries
        qNames(q.Name) = True
    Next q

    txt = "queries:" & vbCrLf
    For Each q In ThisWorkbook.Queries
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
Private Function GroupOf(queryName As String) As String
    Dim p As Variant
    If Not mGroups Is Nothing Then
        If mGroups.Exists(queryName) Then GroupOf = mGroups(queryName): Exit Function
    End If
    If Not mPrefixes Is Nothing Then
        For Each p In mPrefixes
            If Left$(queryName, Len(p(0))) = p(0) Then GroupOf = p(1): Exit Function
        Next p
    End If
    mUnmappedN = mUnmappedN + 1
    If mUnmappedN <= 10 Then mUnmapped = mUnmapped & vbCrLf & "  * " & queryName
End Function

' Function | Source | Parameter | Transform | Standalone (за кодом без коментарів)
Private Function DetectQueryType(formula As String, hasDeps As Boolean) As String
    Dim f As String, fn As Variant
    f = LCase$(Trim$(Replace(Replace(engStripMComments(formula), vbCr, " "), vbLf, " ")))

    If Left$(f, 1) = "(" Then DetectQueryType = "Function": Exit Function

    For Each fn In Array("sql.database(", "sql.databases(", "access.database(", "oracle.database(", _
                         "odbc.datasource(", "oledb.datasource(", "folder.files(", "folder.contents(", _
                         "file.contents(", "excel.workbook(", "csv.document(", "xml.tables(", _
                         "json.document(", "sharepoint.files(", "sharepoint.contents(", "web.contents(")
        If InStr(f, CStr(fn)) > 0 Then DetectQueryType = "Source": Exit Function
    Next fn

    If InStr(f, "excel.currentworkbook(") > 0 Then
        If InStr(Replace(f, " ", ""), "[name=") > 0 Then DetectQueryType = "Parameter" Else DetectQueryType = "Source"
        Exit Function
    End If

    If hasDeps Then DetectQueryType = "Transform" Else DetectQueryType = "Standalone"
End Function

Private Function YamlQuote(value As String) As String
    YamlQuote = """" & Replace(Replace(value, "\", "\\"), """", "\""") & """"
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
