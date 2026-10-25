Attribute VB_Name = "modEngineCommon"
Option Explicit

' ============================================================================
' modEngineCommon: спільні процедури макросів Power Query Engine (аудит D13; v4.11)
'
' Імпортуйте разом із modSelectSourceFiles, modExportPowerQueryCodesToRepo,
' modPOWERQUERYPROFILER, modRunLog, modCleanExportSelectedSheets: вони викликають
' процедури eng* звідси. Самодостатніми навмисно лишаються modImportPowerQueryFromFolder
' і modSetupTestEngine (ними збирають книгу з нуля).
'
' Параметри Engine мають одне джерело - запит cfgSettings. engLoadSettings читає з його
' тексту CsvDelimiter, CsvEncoding, TrueValues, Tables (імена таблиць) і Extensions;
' engSetting("ключ") повертає значення, engTableName("Registry") - ім'я таблиці.
' ============================================================================

Public Const ENG_SETTINGS_QUERY As String = "cfgSettings"

Private mSettings As Object      ' Scripting.Dictionary: ключ -> значення
Private mNote As String          ' попередження останнього читання cfgSettings

' ================== ПАРАМЕТРИ З cfgSettings ==================
' Повертає "" або текст попередження (які ключі не знайдено; для них діють аварійні значення)
Public Function engLoadSettings() As String
    Dim f As String, re As Object, rk As Object, it As Object, inner As String, lst As String, miss As String

    Set mSettings = CreateObject("Scripting.Dictionary")
    mSettings.CompareMode = vbTextCompare
    ' аварійні значення: діють, лише якщо cfgSettings не вдалося прочитати
    mSettings("CsvDelimiter") = ";"
    mSettings("CsvEncoding") = 1251
    mSettings("TrueValues") = "|true|1|так|yes|y|+|"
    mSettings("Tables.Registry") = "tblSourceRegistry"
    mSettings("Tables.Schema") = "tblSchemaRegistry"
    mSettings("Tables.RowFilters") = "tblRowFilters"
    mSettings("Tables.DefaultHeaders") = "tblDefaultHeaders"
    mSettings("Tables.TargetTypes") = "tblTargetType"
    mSettings("Tables.PeriodRules") = "tblPeriodRules"
    mSettings("Extensions.Csv") = "|csv|txt|"
    mSettings("Extensions.Excel") = "|xlsx|xlsm|xls|"
    mNote = ""

    On Error Resume Next
    f = ThisWorkbook.Queries(ENG_SETTINGS_QUERY).formula
    On Error GoTo 0
    If f = "" Then
        mNote = "Запит " & ENG_SETTINGS_QUERY & " не знайдено в книзі; використано аварійні значення."
        engLoadSettings = mNote
        Exit Function
    End If
    f = engStripMComments(f)

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = False

    re.Pattern = "\bCsvDelimiter\s*=\s*""((?:[^""]|"""")*)"""
    If re.Test(f) Then
        mSettings("CsvDelimiter") = engUnescapeM(re.Execute(f)(0).SubMatches(0))
    Else
        miss = miss & " CsvDelimiter"
    End If

    re.Pattern = "\bCsvEncoding\s*=\s*(\d+)"
    If re.Test(f) Then
        mSettings("CsvEncoding") = CLng(re.Execute(f)(0).SubMatches(0))
    Else
        miss = miss & " CsvEncoding"
    End If

    lst = QuotedListAfter(f, "TrueValues")
    If lst <> "" Then mSettings("TrueValues") = LCase$(lst) Else miss = miss & " TrueValues"

    ' Tables = [Registry = "tblSourceRegistry", ...]
    re.Pattern = "\bTables\s*=\s*\[([^\]]*)\]"
    If re.Test(f) Then
        inner = re.Execute(f)(0).SubMatches(0)
        Set rk = CreateObject("VBScript.RegExp")
        rk.Global = True
        rk.Pattern = "(\w+)\s*=\s*""((?:[^""]|"""")*)"""
        For Each it In rk.Execute(inner)
            mSettings("Tables." & it.SubMatches(0)) = engUnescapeM(it.SubMatches(1))
        Next it
    Else
        miss = miss & " Tables"
    End If

    ' Extensions = [Csv = {"csv", "txt"}, Excel = {"xlsx", "xlsm", "xls"}]
    re.Pattern = "\bExtensions\s*=\s*\[([^\]]*)\]"
    If re.Test(f) Then
        inner = re.Execute(f)(0).SubMatches(0)
        lst = QuotedListAfter(inner, "Csv")
        If lst <> "" Then mSettings("Extensions.Csv") = LCase$(lst)
        lst = QuotedListAfter(inner, "Excel")
        If lst <> "" Then mSettings("Extensions.Excel") = LCase$(lst)
    Else
        miss = miss & " Extensions"
    End If

    If mSettings("CsvDelimiter") = "" Then mSettings("CsvDelimiter") = ";": miss = miss & " CsvDelimiter(порожній)"
    If miss <> "" Then mNote = "У запиті " & ENG_SETTINGS_QUERY & " не знайдено:" & miss & _
                               " (для них використано аварійні значення)."
    engLoadSettings = mNote
End Function

' Значення параметра; при першому зверненні читає cfgSettings
Public Function engSetting(keyName As String) As Variant
    If mSettings Is Nothing Then engLoadSettings
    If mSettings.Exists(keyName) Then engSetting = mSettings(keyName) Else engSetting = Empty
End Function

' Ім'я таблиці книги з cfgSettings[Tables]: engTableName("Registry") -> "tblSourceRegistry"
Public Function engTableName(keyName As String) As String
    engTableName = CStr(engSetting("Tables." & keyName))
End Function

' Попередження останнього engLoadSettings ("" = усе прочитано)
Public Function engSettingsNote() As String
    engSettingsNote = mNote
End Function

' Розширення, які читає Engine: kind = "Csv", "Excel" або "" (будь-яке з них)
Public Function engIsExtension(filePath As String, Optional kind As String = "") As Boolean
    Dim ext As String
    ext = "|" & LCase$(Mid$(filePath, InStrRev(filePath, ".") + 1)) & "|"
    If kind = "" Then
        engIsExtension = InStr(1, CStr(engSetting("Extensions.Csv")) & CStr(engSetting("Extensions.Excel")), ext) > 0
    Else
        engIsExtension = InStr(1, CStr(engSetting("Extensions." & kind)), ext) > 0
    End If
End Function

' Як fn_ToBool: порожнє = dflt, логічне як є, число <> 0, текст з cfgSettings[TrueValues]
Public Function engToBool(v As Variant, dflt As Boolean) As Boolean
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then engToBool = dflt: Exit Function
    If VarType(v) = vbBoolean Then engToBool = v: Exit Function
    If VarType(v) <> vbString And IsNumeric(v) Then engToBool = (v <> 0): Exit Function
    Dim t As String: t = LCase$(Trim$(CStr(v)))
    If t = "" Then
        engToBool = dflt
    Else
        engToBool = (InStr(1, CStr(engSetting("TrueValues")), "|" & t & "|") > 0)
    End If
End Function

' Список рядків у лапках після «ключ = {...}» у вигляді "|a|b|"; "" якщо ключа немає
Private Function QuotedListAfter(src As String, keyName As String) As String
    Dim re As Object, rq As Object, it As Object, res As String
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\b" & keyName & "\s*=\s*\{([^}]*)\}"
    If Not re.Test(src) Then Exit Function
    Set rq = CreateObject("VBScript.RegExp")
    rq.Global = True
    rq.Pattern = """((?:[^""]|"""")*)"""
    For Each it In rq.Execute(re.Execute(src)(0).SubMatches(0))
        res = res & "|" & Trim$(engUnescapeM(it.SubMatches(0)))
    Next it
    If res <> "" Then QuotedListAfter = res & "|"
End Function

' ================== ТЕКСТ M ==================
' Видаляє коментарі M (// до кінця рядка, /* ... */), не чіпаючи текст у лапках
Public Function engStripMComments(src As String) As String
    Dim i As Long, n As Long, p As Long, ch As String, res As String, inQ As Boolean
    n = Len(src)
    i = 1
    Do While i <= n
        ch = Mid$(src, i, 1)
        If inQ Then
            res = res & ch
            If ch = """" Then
                If Mid$(src, i + 1, 1) = """" Then res = res & """": i = i + 1 Else inQ = False
            End If
        ElseIf ch = """" Then
            inQ = True
            res = res & ch
        ElseIf Mid$(src, i, 2) = "//" Then
            Do While i <= n
                If Mid$(src, i, 1) = vbLf Then Exit Do
                i = i + 1
            Loop
            res = res & vbLf
        ElseIf Mid$(src, i, 2) = "/*" Then
            p = InStr(i + 2, src, "*/")
            If p = 0 Then i = n Else i = p + 1
        Else
            res = res & ch
        End If
        i = i + 1
    Loop
    engStripMComments = res
End Function

' Розекранування текстового літерала M: "" -> ", #(tab), #(lf), #(cr), #(#)
Public Function engUnescapeM(ByVal t As String) As String
    t = Replace(t, """""", """")
    t = Replace(t, "#(tab)", vbTab, , , vbTextCompare)
    t = Replace(t, "#(lf)", vbLf, , , vbTextCompare)
    t = Replace(t, "#(cr)", vbCr, , , vbTextCompare)
    t = Replace(t, "#(#)", "#")
    engUnescapeM = t
End Function

' Ідентифікатори коду M (без коментарів і текстових літералів): звичайні і #"..."
' Повертає Dictionary ім'я -> True; регістр важливий, як у M
Public Function engMIdentifiers(mCode As String) As Object
    Dim res As Object, re As Object, it As Object
    Set res = CreateObject("Scripting.Dictionary")
    res.CompareMode = vbBinaryCompare
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "#""((?:[^""]|"""")*)""|""(?:[^""]|"""")*""|([A-Za-z_\u0400-\u04FF][A-Za-z0-9_\u0400-\u04FF]*)"
    For Each it In re.Execute(engStripMComments(mCode))
        If Len(it.SubMatches(0)) > 0 Then
            res(Replace(it.SubMatches(0), """""", """")) = True
        ElseIf Len(it.SubMatches(1)) > 0 Then
            res(it.SubMatches(1)) = True
        End If
    Next it
    Set engMIdentifiers = res
End Function

' ================== ТАБЛИЦІ ==================
Public Function engFindTable(tblName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tblName, vbTextCompare) = 0 Then Set engFindTable = lo: Exit Function
        Next lo
    Next ws
End Function

Public Function engColumnExists(lo As ListObject, colName As String) As Boolean
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, colName, vbTextCompare) = 0 Then engColumnExists = True: Exit Function
    Next lc
End Function

' Значення клітинки рядка r таблиці; клітинка-помилка (#N/A тощо) = Empty
Public Function engCell(lo As ListObject, r As Long, col As String) As Variant
    Dim v As Variant
    v = lo.ListColumns(col).DataBodyRange.Cells(r, 1).value
    If IsError(v) Then engCell = Empty Else engCell = v
End Function

' Значення необов'язкової колонки; якщо колонки немає, dflt
Public Function engCellOpt(lo As ListObject, r As Long, col As String, dflt As Variant) As Variant
    If engColumnExists(lo, col) Then engCellOpt = engCell(lo, r, col) Else engCellOpt = dflt
End Function

Public Sub engSetCell(lo As ListObject, r As Long, col As String, v As Variant)
    lo.ListColumns(col).DataBodyRange.Cells(r, 1).value = v
End Sub

' Шлях як у fn_SourcePath: FullPath має пріоритет, інакше Path + FileName
Public Function engSourcePath(lo As ListObject, r As Long) As String
    Dim full As String, dir_ As String, nm As String
    full = Trim$(CStr(engCellOpt(lo, r, "FullPath", "")))
    dir_ = Trim$(CStr(engCellOpt(lo, r, "Path", "")))
    nm = Trim$(CStr(engCellOpt(lo, r, "FileName", "")))
    If full <> "" Then
        engSourcePath = full
    ElseIf nm = "" Then
        engSourcePath = ""
    ElseIf dir_ = "" Then
        engSourcePath = nm
    ElseIf Right$(dir_, 1) = "\" Then
        engSourcePath = dir_ & nm
    Else
        engSourcePath = dir_ & "\" & nm
    End If
End Function

' ================== ЗАПИТИ І ПІДКЛЮЧЕННЯ ==================
' Ім'я запиту Power Query за підключенням: параметр Location рядка підключення
' (не залежить від мови Excel, на відміну від назви «Query - X» / «Запит – X»)
Public Function engQueryNameOf(cn As WorkbookConnection) As String
    Dim s As String, p As Long, q As Long
    On Error Resume Next
    s = CStr(cn.OLEDBConnection.Connection)
    On Error GoTo 0
    p = InStr(1, s, "Location=", vbTextCompare)
    If p = 0 Then Exit Function
    p = p + Len("Location=")
    If Mid$(s, p, 1) = """" Then
        q = InStr(p + 1, s, """")
        If q > p Then engQueryNameOf = Mid$(s, p + 1, q - p - 1)
    Else
        q = InStr(p, s, ";")
        If q = 0 Then q = Len(s) + 1
        engQueryNameOf = Mid$(s, p, q - p)
    End If
End Function

' Таблиця на аркуші, завантажена з запиту queryName (за Location підключення, інакше за CommandText)
Public Function engFindQueryTable(queryName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject, qn As String, cmd As String
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            qn = "": cmd = ""
            On Error Resume Next
            If lo.SourceType = xlSrcQuery Then
                qn = engQueryNameOf(lo.QueryTable.WorkbookConnection)
                cmd = CStr(lo.QueryTable.CommandText)
            ElseIf lo.SourceType = xlSrcModel Then
                ' таблиця з моделі даних: підключення моделі не містить Location;
                ' Excel називає таку таблицю за запитом (пробіли -> "_")
                qn = engQueryNameOf(lo.TableObject.WorkbookConnection)
                If qn = "" And StrComp(lo.Name, Replace(queryName, " ", "_"), vbTextCompare) = 0 Then qn = queryName
            End If
            On Error GoTo 0
            If StrComp(qn, queryName, vbTextCompare) = 0 _
               Or InStr(1, cmd, "[" & queryName & "]", vbTextCompare) > 0 Then
                Set engFindQueryTable = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

' Завантажує запит на аркуш як таблицю (без оновлення: дані з'являться після «Оновити все»
' або lo.QueryTable.Refresh). Якщо запит уже завантажено, повертає наявну таблицю.
' Правила hlpProject: таблиця називається як запит (q_, pq_, test_ - роль видно з назви),
' шапка не вище рядка 3 (topLeft за замовчуванням A3).
' refreshWithAll: позначка підключення «Оновлювати під час оновлення всіх». Звіти якості й Calendar
' з v4.13 вимкнено з «Оновити все» (оновлюються на вимогу), тому для них передають False.
Public Function engLoadQueryToSheet(queryName As String, ws As Worksheet, Optional topLeft As String = "A3", _
                                    Optional refreshWithAll As Boolean = True) As ListObject
    Dim lo As ListObject
    Set lo = engFindQueryTable(queryName)
    If Not lo Is Nothing Then Set engLoadQueryToSheet = lo: Exit Function
    Set lo = ws.ListObjects.Add(SourceType:=0, _
        Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & queryName & _
                ";Extended Properties=""""", Destination:=ws.Range(topLeft))
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & queryName & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SaveData = True
        .AdjustColumnWidth = True
        .PreserveColumnInfo = True
    End With
    On Error Resume Next
    lo.QueryTable.WorkbookConnection.RefreshWithRefreshAll = refreshWithAll
    lo.Name = queryName
    On Error GoTo 0
    Set engLoadQueryToSheet = lo
End Function

' Оновлює таблицю запиту, якщо її підключення вимкнено з «Оновити все» (звіти якості, v4.13).
' Викликати після engRefreshAllSync, коли потрібні свіжі дані такої таблиці. Повертає "" або текст
' помилки; якщо таблиці немає або вона й так оновлюється з усіма - нічого не робить.
Public Function engRefreshIfExcluded(queryName As String) As String
    Dim lo As ListObject, excluded As Boolean
    Set lo = engFindQueryTable(queryName)
    If lo Is Nothing Then Exit Function
    On Error Resume Next
    excluded = Not lo.QueryTable.WorkbookConnection.RefreshWithRefreshAll
    If Err.Number <> 0 Then excluded = True: Err.Clear
    If excluded Then
        lo.QueryTable.Refresh BackgroundQuery:=False
        If Err.Number <> 0 Then engRefreshIfExcluded = queryName & ": " & Err.Description
    End If
    On Error GoTo 0
End Function

' Синхронне «Оновити все»: фонове оновлення вимикається і потім відновлюється.
' Повертає "" або текст помилки. Помилки окремих запитів Excel показує власним вікном,
' а не через Err, тому їх видно в pq_QualitySummary і журналі, а не тут.
Public Function engRefreshAllSync() As String
    Dim cn As WorkbookConnection, bg As Object, nm As Variant, b As Boolean
    Set bg = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    For Each cn In ThisWorkbook.Connections
        Err.Clear
        b = cn.OLEDBConnection.BackgroundQuery          ' не всі підключення мають OLEDB
        If Err.Number = 0 Then bg(cn.Name) = b: cn.OLEDBConnection.BackgroundQuery = False
    Next cn
    Err.Clear
    ThisWorkbook.RefreshAll
    If Err.Number <> 0 Then engRefreshAllSync = Err.Description
    Err.Clear
    Application.CalculateUntilAsyncQueriesDone
    For Each nm In bg.Keys
        ThisWorkbook.Connections(CStr(nm)).OLEDBConnection.BackgroundQuery = bg(nm)
    Next nm
    On Error GoTo 0
End Function

' ================== ФАЙЛИ ==================
Public Function engFileExists(filePath As String) As Boolean
    On Error Resume Next
    engFileExists = (Len(filePath) > 0 And Dir(filePath) <> "")
    On Error GoTo 0
End Function

' Недопустимі в імені файла символи замінюються на "_"
Public Function engSafeFileName(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, CStr(ch), "_")
    Next ch
    engSafeFileName = s
End Function

' Створює теку разом з усіма проміжними
Public Sub engEnsureFolder(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Right$(folderPath, 1) = "\" Then folderPath = Left$(folderPath, Len(folderPath) - 1)
    If folderPath = "" Or fso.FolderExists(folderPath) Then Exit Sub
    engEnsureFolder fso.GetParentFolderName(folderPath)
    fso.CreateFolder folderPath
End Sub

Public Function engReadUtf8(filePath As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.charset = "utf-8"
    st.Open
    st.LoadFromFile filePath
    engReadUtf8 = st.ReadText
    st.Close
End Function

' Запис тексту в UTF-8 (з BOM, як і раніше в експорті запитів)
Public Sub engWriteUtf8(filePath As String, txt As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.charset = "utf-8"
    st.Open
    st.WriteText txt
    st.SaveToFile filePath, 2
    st.Close
End Sub
